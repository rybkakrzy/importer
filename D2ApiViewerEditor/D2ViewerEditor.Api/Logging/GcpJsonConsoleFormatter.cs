using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Logging.Console;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Logging;

// Console formatter emitting each log as a single JSON line readable by BOTH Google Cloud Logging
// (severity, logging.googleapis.com/* trace fields, httpRequest, labels) and Elasticsearch/ELK
// (Elastic Common Schema: @timestamp, log.*, service.*, trace.*, error.*, http.*, url.*).
//
// Design notes:
//  - Trace/span come from Activity.Current (ASP.NET W3C trace context) — no extra tracing dependency.
//  - HTTP context is optional (IHttpContextAccessor); the formatter works in background workers too.
//  - Exceptions are written as structured error.* fields (type/message/stack_trace/inner) AND appended
//    to `message` so GCP Error Reporting still groups by the stack trace. The Exception object itself is
//    never handed to the serializer (that yields huge, unstable JSON).
//  - Known-sensitive property/query names are redacted to "[REDACTED]".
//  - Scope/state properties never overwrite system fields.
public sealed class GcpJsonConsoleFormatter : ConsoleFormatter
{
    public const string FormatterName = "gcp-json";
    private const string Redacted = "[REDACTED]";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        // System.Text.Json refuses to serialize reflection metadata (Type, MethodBase, Assembly, …) as a
        // deserialization-gadget guard. Log payloads can carry such values (e.g. a destructured {@Request}
        // with a Type property) — emit them as text instead of throwing.
        Converters = { new TypeJsonConverter(), new ReflectionMetadataJsonConverterFactory() }
    };

    // Top-level keys owned by the formatter — scope/state properties matching these are dropped so they
    // can never overwrite a system field.
    private static readonly HashSet<string> ReservedKeys = new(StringComparer.Ordinal)
    {
        "timestamp", "@timestamp", "severity", "level", "message", "category", "service", "environment",
        "traceId", "spanId", "eventId", "exceptionType", "correlation_id", "labels", "log", "trace",
        "span", "transaction", "event", "error", "http", "url", "user", "httpRequest",
        "logging.googleapis.com/trace", "logging.googleapis.com/spanId",
        "logging.googleapis.com/trace_sampled", "logging.googleapis.com/sourceLocation"
    };

    private readonly StructuredLogFormatterOptions _options;
    private readonly IHttpContextAccessor? _httpContextAccessor;
    private readonly HashSet<string> _redactedKeys;

    public GcpJsonConsoleFormatter(
        IOptions<StructuredLogFormatterOptions> options,
        IHttpContextAccessor? httpContextAccessor = null)
        : base(FormatterName)
    {
        _options = options.Value;
        _httpContextAccessor = httpContextAccessor;
        // Pre-normalize the redact list once so "access_token", "accessToken" and "access-token" all match.
        _redactedKeys = new HashSet<string>(
            _options.RedactedPropertyNames.Select(NormalizeKey), StringComparer.OrdinalIgnoreCase);
    }

    public override void Write<TState>(
        in LogEntry<TState> logEntry, IExternalScopeProvider? scopeProvider, TextWriter textWriter)
    {
        var message = logEntry.Formatter?.Invoke(logEntry.State, logEntry.Exception) ?? string.Empty;
        if (string.IsNullOrEmpty(message) && logEntry.Exception is null)
            return;

        var props = new Dictionary<string, object?>(StringComparer.Ordinal);
        if (_options.IncludeScopes)
            scopeProvider?.ForEachScope((scope, target) => CollectPairs(scope, target), props);
        var template = CollectStateWithTemplate(logEntry.State, props);

        var http = _options.IncludeHttpRequest ? _httpContextAccessor?.HttpContext : null;
        var activity = Activity.Current;
        var correlationId = ResolveCorrelationId(props, http);
        var timestamp = DateTimeOffset.UtcNow.ToString("o");
        var fullMessage = BuildMessage(message, logEntry.Exception);

        var payload = new Dictionary<string, object?>(StringComparer.Ordinal);
        AddCoreFields(payload, logEntry, fullMessage, timestamp, correlationId, activity);

        if (_options.IncludeGoogleCloudFields)
            AddGoogleCloudFields(payload, logEntry, activity, http);
        if (_options.IncludeElasticCommonSchemaFields)
            AddElasticFields(payload, logEntry, template, timestamp, correlationId, activity, http);
        if (logEntry.Exception is not null && _options.IncludeElasticCommonSchemaFields)
            payload["error"] = BuildError(logEntry.Exception);

        AddLabels(payload, correlationId, props);
        MergeRemainingProps(payload, props);

        textWriter.WriteLine(Serialize(payload, logEntry, fullMessage));
    }

    private void AddCoreFields<TState>(
        Dictionary<string, object?> payload, in LogEntry<TState> logEntry, string fullMessage,
        string timestamp, string? correlationId, Activity? activity)
    {
        payload["timestamp"] = timestamp;
        payload["severity"] = MapSeverity(logEntry.LogLevel);
        payload["level"] = logEntry.LogLevel.ToString();
        payload["message"] = fullMessage;
        payload["category"] = logEntry.Category;

        // Flat service/environment kept only when ECS is off; under ECS they live in the service object.
        if (!_options.IncludeElasticCommonSchemaFields)
        {
            if (!string.IsNullOrEmpty(_options.ResolvedServiceName))
                payload["service"] = _options.ResolvedServiceName;
            if (!string.IsNullOrEmpty(_options.ResolvedEnvironment))
                payload["environment"] = _options.ResolvedEnvironment;
        }

        if (activity is not null)
        {
            payload["traceId"] = activity.TraceId.ToString();
            payload["spanId"] = activity.SpanId.ToString();
        }

        if (_options.IncludeEventId && logEntry.EventId.Id != 0)
            payload["eventId"] = logEntry.EventId.Id;
        if (logEntry.Exception is not null)
            payload["exceptionType"] = logEntry.Exception.GetType().FullName;
        if (!string.IsNullOrEmpty(correlationId))
            payload["correlation_id"] = correlationId;
    }

    private void AddGoogleCloudFields<TState>(
        Dictionary<string, object?> payload, in LogEntry<TState> logEntry, Activity? activity, HttpContext? http)
    {
        if (activity is not null)
        {
            if (!string.IsNullOrEmpty(_options.ProjectId))
                payload["logging.googleapis.com/trace"] =
                    $"projects/{_options.ProjectId}/traces/{activity.TraceId}";
            payload["logging.googleapis.com/spanId"] = activity.SpanId.ToString();
            payload["logging.googleapis.com/trace_sampled"] =
                activity.ActivityTraceFlags.HasFlag(ActivityTraceFlags.Recorded);
        }

        if (_options.IncludeSourceLocation && logEntry.Exception is not null
            && TryBuildSourceLocation(logEntry.Exception, out var sourceLocation))
            payload["logging.googleapis.com/sourceLocation"] = sourceLocation;

        if (http is not null)
            payload["httpRequest"] = BuildGcpHttpRequest(http);
    }

    private void AddElasticFields<TState>(
        Dictionary<string, object?> payload, in LogEntry<TState> logEntry, string? template,
        string timestamp, string? correlationId, Activity? activity, HttpContext? http)
    {
        payload["@timestamp"] = timestamp;

        var log = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["level"] = logEntry.LogLevel.ToString(),
            ["logger"] = logEntry.Category
        };
        if (!string.IsNullOrEmpty(template))
            log["template"] = template;
        payload["log"] = log;

        payload["service"] = BuildService();

        if (activity is not null)
        {
            payload["trace"] = new Dictionary<string, object?> { ["id"] = activity.TraceId.ToString() };
            payload["span"] = new Dictionary<string, object?> { ["id"] = activity.SpanId.ToString() };
        }
        if (!string.IsNullOrEmpty(correlationId))
            payload["transaction"] = new Dictionary<string, object?> { ["id"] = correlationId };

        payload["event"] = BuildEvent(logEntry.Exception, http);

        if (http is not null)
        {
            payload["http"] = BuildEcsHttp(http);
            payload["url"] = BuildEcsUrl(http);
            var userId = ResolveUserId(http);
            if (userId is not null)
                payload["user"] = new Dictionary<string, object?> { ["id"] = userId };
        }
    }

    private Dictionary<string, object?> BuildService()
    {
        var service = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["name"] = _options.ResolvedServiceName
        };
        if (!string.IsNullOrEmpty(_options.ServiceVersion))
            service["version"] = _options.ServiceVersion;
        if (!string.IsNullOrEmpty(_options.ResolvedEnvironment))
            service["environment"] = _options.ResolvedEnvironment;
        if (!string.IsNullOrEmpty(_options.ServiceInstanceId))
            service["instance"] = new Dictionary<string, object?> { ["id"] = _options.ServiceInstanceId };
        return service;
    }

    private Dictionary<string, object?> BuildEvent(Exception? exception, HttpContext? http)
    {
        var ev = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["kind"] = "event",
            ["type"] = new[] { exception is null ? "info" : "error" }
        };
        if (!string.IsNullOrEmpty(_options.ResolvedServiceName))
            ev["dataset"] = $"{_options.ResolvedServiceName}.log";
        if (http is not null)
            ev["category"] = new[] { "web" };
        if (exception is not null)
            // Low-cardinality classification so a 500 can be triaged from the log alone.
            ev["reason"] = ClassifyError(exception);
        return ev;
    }

    private static Dictionary<string, object?> BuildError(Exception exception)
    {
        var error = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["type"] = exception.GetType().FullName,
            ["message"] = exception.Message,
            ["stack_trace"] = exception.ToString()
        };

        var inner = BuildInnerExceptions(exception);
        if (inner.Count > 0)
            error["inner"] = inner;
        return error;
    }

    private static List<Dictionary<string, object?>> BuildInnerExceptions(Exception exception)
    {
        var inner = new List<Dictionary<string, object?>>();

        if (exception is AggregateException aggregate)
        {
            foreach (var item in aggregate.InnerExceptions)
                inner.Add(DescribeException(item));
            return inner;
        }

        for (var current = exception.InnerException; current is not null; current = current.InnerException)
            inner.Add(DescribeException(current));
        return inner;
    }

    private static Dictionary<string, object?> DescribeException(Exception exception) => new(StringComparer.Ordinal)
    {
        ["type"] = exception.GetType().FullName,
        ["message"] = exception.Message
    };

    private Dictionary<string, object?> BuildGcpHttpRequest(HttpContext http)
    {
        var request = http.Request;
        var result = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["requestMethod"] = request.Method,
            ["requestUrl"] = BuildPath(request) + MaskedQuery(request),
            ["protocol"] = request.Protocol
        };
        if (http.Response.HasStarted || http.Response.StatusCode != 0)
            result["status"] = http.Response.StatusCode;
        if (request.Headers.TryGetValue("User-Agent", out var ua))
            result["userAgent"] = ua.ToString();
        return result;
    }

    private Dictionary<string, object?> BuildEcsHttp(HttpContext http) => new(StringComparer.Ordinal)
    {
        ["request"] = new Dictionary<string, object?> { ["method"] = http.Request.Method },
        ["response"] = new Dictionary<string, object?> { ["status_code"] = http.Response.StatusCode }
    };

    private Dictionary<string, object?> BuildEcsUrl(HttpContext http)
    {
        var url = new Dictionary<string, object?>(StringComparer.Ordinal) { ["path"] = BuildPath(http.Request) };
        if (http.Request.QueryString.HasValue)
            url["query"] = MaskQuery(http.Request.QueryString.Value!.TrimStart('?'));
        return url;
    }

    private static string BuildPath(HttpRequest request) => request.PathBase.Add(request.Path).Value ?? "/";

    private string MaskedQuery(HttpRequest request) =>
        request.QueryString.HasValue ? "?" + MaskQuery(request.QueryString.Value!.TrimStart('?')) : string.Empty;

    private string MaskQuery(string query)
    {
        if (string.IsNullOrEmpty(query))
            return query;

        var parts = query.Split('&');
        for (var i = 0; i < parts.Length; i++)
        {
            var eq = parts[i].IndexOf('=');
            if (eq <= 0)
                continue;
            var name = parts[i][..eq];
            if (IsRedacted(Uri.UnescapeDataString(name)))
                parts[i] = name + "=" + Redacted;
        }
        return string.Join('&', parts);
    }

    private string? ResolveCorrelationId(Dictionary<string, object?> props, HttpContext? http)
    {
        if (http is not null
            && http.Items.TryGetValue(Middleware.RequestObservabilityMiddleware.CorrelationIdItemKey, out var item)
            && item is string fromItems && !string.IsNullOrWhiteSpace(fromItems))
            return fromItems;

        if (http is not null
            && http.Request.Headers.TryGetValue(Middleware.RequestObservabilityMiddleware.CorrelationIdHeader, out var header)
            && !string.IsNullOrWhiteSpace(header))
            return header.ToString();

        if (props.TryGetValue("correlationId", out var fromScope) && fromScope is string s && !string.IsNullOrWhiteSpace(s))
            return s;

        return Activity.Current?.GetBaggageItem("correlation_id");
    }

    private static string? ResolveUserId(HttpContext http)
    {
        var user = http.User;
        if (user?.Identity?.IsAuthenticated != true)
            return null;

        // Technical identifiers only — never name/email (PII).
        return user.FindFirst("oid")?.Value
               ?? user.FindFirst("sub")?.Value
               ?? user.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
    }

    private void AddLabels(Dictionary<string, object?> payload, string? correlationId, Dictionary<string, object?> props)
    {
        var labels = new Dictionary<string, object?>(StringComparer.Ordinal);
        if (!string.IsNullOrEmpty(correlationId))
            labels["correlation_id"] = correlationId;
        CopyLabel(props, labels, "operation", "operation");
        CopyLabel(props, labels, "tenantId", "tenant_id");
        CopyLabel(props, labels, "tenant_id", "tenant_id");

        if (labels.Count > 0)
            payload["labels"] = labels;
    }

    private static void CopyLabel(
        Dictionary<string, object?> props, Dictionary<string, object?> labels, string sourceKey, string labelKey)
    {
        if (!labels.ContainsKey(labelKey) && props.TryGetValue(sourceKey, out var value) && value is not null)
            labels[labelKey] = value.ToString();
    }

    private void MergeRemainingProps(Dictionary<string, object?> payload, Dictionary<string, object?> props)
    {
        foreach (var pair in props)
        {
            if (ReservedKeys.Contains(pair.Key) || payload.ContainsKey(pair.Key))
                continue;
            payload[pair.Key] = pair.Value;
        }
    }

    private void CollectPairs(object? source, Dictionary<string, object?> target)
    {
        if (source is not IEnumerable<KeyValuePair<string, object>> pairs)
            return;

        foreach (var pair in pairs)
        {
            if (pair.Key == "{OriginalFormat}")
                continue;
            target[pair.Key] = IsRedacted(pair.Key) ? Redacted : pair.Value;
        }
    }

    private string? CollectStateWithTemplate(object? state, Dictionary<string, object?> target)
    {
        if (state is not IEnumerable<KeyValuePair<string, object>> pairs)
            return null;

        string? template = null;
        foreach (var pair in pairs)
        {
            if (pair.Key == "{OriginalFormat}")
            {
                template = pair.Value as string;
                continue;
            }
            target[pair.Key] = IsRedacted(pair.Key) ? Redacted : pair.Value;
        }
        return template;
    }

    private bool IsRedacted(string key) => _redactedKeys.Contains(NormalizeKey(key));

    private static string NormalizeKey(string key) => key.Replace("_", string.Empty).Replace("-", string.Empty);

    private static bool TryBuildSourceLocation(Exception exception, out Dictionary<string, object?> location)
    {
        var frame = new StackTrace(exception, fNeedFileInfo: true).GetFrame(0);
        var file = frame?.GetFileName();
        if (frame is null || string.IsNullOrEmpty(file))
        {
            location = default!;
            return false;
        }

        location = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["file"] = file,
            ["line"] = frame.GetFileLineNumber().ToString(),
            ["function"] = frame.GetMethod()?.Name
        };
        return true;
    }

    private static string ClassifyError(Exception exception) => exception switch
    {
        OperationCanceledException => "cancellation",
        UnauthorizedAccessException => "authorization",
        System.Net.Http.HttpRequestException => "dependency",
        System.Net.Sockets.SocketException => "dependency",
        TimeoutException => "dependency",
        System.Data.Common.DbException => "database",
        _ => ClassifyByName(exception)
    };

    private static string ClassifyByName(Exception exception)
    {
        var name = exception.GetType().FullName ?? string.Empty;
        if (name.Contains("Validation", StringComparison.OrdinalIgnoreCase)) return "validation";
        if (name.Contains("Npgsql", StringComparison.OrdinalIgnoreCase)
            || name.Contains("Postgres", StringComparison.OrdinalIgnoreCase)
            || name.Contains("SqlException", StringComparison.OrdinalIgnoreCase)) return "database";
        if (name.Contains("Authentication", StringComparison.OrdinalIgnoreCase)
            || name.Contains("Forbidden", StringComparison.OrdinalIgnoreCase)) return "authorization";
        return "code";
    }

    private static string BuildMessage(string message, Exception? exception) =>
        exception is null ? message : $"{message}\n{exception}";

    private string Serialize<TState>(
        Dictionary<string, object?> payload, in LogEntry<TState> logEntry, string fullMessage)
    {
        try
        {
            return JsonSerializer.Serialize(payload, JsonOptions);
        }
        catch (Exception ex)
        {
            // A logger must never throw. Fall back to a minimal line that still carries the message.
            return JsonSerializer.Serialize(new Dictionary<string, object?>
            {
                ["timestamp"] = DateTimeOffset.UtcNow.ToString("o"),
                ["severity"] = MapSeverity(logEntry.LogLevel),
                ["level"] = logEntry.LogLevel.ToString(),
                ["message"] = fullMessage,
                ["category"] = logEntry.Category,
                ["serializationError"] = ex.Message
            }, JsonOptions);
        }
    }

    private static string MapSeverity(LogLevel level) => level switch
    {
        LogLevel.Trace => "DEBUG",
        LogLevel.Debug => "DEBUG",
        LogLevel.Information => "INFO",
        LogLevel.Warning => "WARNING",
        LogLevel.Error => "ERROR",
        LogLevel.Critical => "CRITICAL",
        _ => "DEFAULT"
    };

    // Serializes Type (and its runtime subtype System.RuntimeType) as its full name, since
    // JsonSerializer otherwise throws NotSupportedException.
    private sealed class TypeJsonConverter : JsonConverter<Type>
    {
        public override bool CanConvert(Type typeToConvert) => typeof(Type).IsAssignableFrom(typeToConvert);

        public override Type Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            => throw new NotSupportedException();

        public override void Write(Utf8JsonWriter writer, Type value, JsonSerializerOptions options)
            => writer.WriteStringValue(value.FullName ?? value.Name);
    }

    // Serializes the remaining reflection metadata (MemberInfo/MethodBase, Assembly, Module,
    // ParameterInfo) as text; these also throw NotSupportedException in JsonSerializer and surface when an
    // object such as an Exception (whose TargetSite is a MethodBase) is logged as a structured value.
    private sealed class ReflectionMetadataJsonConverterFactory : JsonConverterFactory
    {
        public override bool CanConvert(Type typeToConvert) =>
            !typeof(Type).IsAssignableFrom(typeToConvert) &&
            (typeof(System.Reflection.MemberInfo).IsAssignableFrom(typeToConvert) ||
             typeof(System.Reflection.Assembly).IsAssignableFrom(typeToConvert) ||
             typeof(System.Reflection.Module).IsAssignableFrom(typeToConvert) ||
             typeof(System.Reflection.ParameterInfo).IsAssignableFrom(typeToConvert));

        public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options) =>
            (JsonConverter)Activator.CreateInstance(
                typeof(ToStringConverter<>).MakeGenericType(typeToConvert))!;

        private sealed class ToStringConverter<T> : JsonConverter<T>
        {
            public override T Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
                => throw new NotSupportedException();

            public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
                => writer.WriteStringValue(value?.ToString());
        }
    }
}

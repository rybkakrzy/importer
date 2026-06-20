using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Logging.Console;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Logging;

/// <summary>
/// Options carrying the static enrichment fields (service + environment) added to every log line.
/// </summary>
public sealed class StructuredLogFormatterOptions : ConsoleFormatterOptions
{
    public string Service { get; set; } = string.Empty;
    public string Environment { get; set; } = string.Empty;
}

/// <summary>
/// Console formatter emitting each log as a single JSON line suitable for both ELK (Elasticsearch /
/// Logstash / Kibana) and Google Cloud Logging.
///
/// <para>Structured fields: timestamp, severity (GCP) + level (text), message, category, service,
/// environment, traceId/spanId (from <see cref="Activity.Current"/>), plus the key/value pairs from
/// logging scopes (correlationId, masterId, …) and the message's structured arguments. Exceptions
/// carry their type and full stack trace in <c>message</c> (GCP Error Reporting groups by it).</para>
///
/// <para>Sensitive data (file content, tokens, secrets, PII) must never be passed to the logger —
/// this formatter emits whatever scope/state it is given.</para>
/// </summary>
public sealed class GcpJsonConsoleFormatter : ConsoleFormatter
{
    public const string FormatterName = "gcp-json";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        // System.Text.Json refuses to serialize Type/RuntimeType (deserialization-gadget guard) and
        // throws NotSupportedException. Log payloads can carry Type values anywhere in the object graph
        // (e.g. a destructured {@Request} with a Type property) — emit them as their name instead.
        Converters = { new TypeJsonConverter() }
    };

    private static readonly HashSet<string> ReservedKeys = new(StringComparer.Ordinal)
    {
        "severity", "level", "message", "category", "timestamp",
        "eventId", "exceptionType", "service", "environment", "traceId", "spanId"
    };

    private readonly StructuredLogFormatterOptions _options;

    public GcpJsonConsoleFormatter(IOptions<StructuredLogFormatterOptions> options) : base(FormatterName)
        => _options = options.Value;

    public override void Write<TState>(
        in LogEntry<TState> logEntry, IExternalScopeProvider? scopeProvider, TextWriter textWriter)
    {
        var message = logEntry.Formatter?.Invoke(logEntry.State, logEntry.Exception) ?? string.Empty;
        if (string.IsNullOrEmpty(message) && logEntry.Exception is null)
            return;

        // GCP Error Reporting groups by the stack trace in the entry body — append the full exception.
        var fullMessage = logEntry.Exception is null ? message : $"{message}\n{logEntry.Exception}";

        var payload = new Dictionary<string, object?>
        {
            ["timestamp"] = DateTimeOffset.UtcNow.ToString("o"),
            ["severity"] = MapSeverity(logEntry.LogLevel),
            ["level"] = logEntry.LogLevel.ToString(),
            ["message"] = fullMessage,
            ["category"] = logEntry.Category,
        };

        if (!string.IsNullOrEmpty(_options.Service))
            payload["service"] = _options.Service;
        if (!string.IsNullOrEmpty(_options.Environment))
            payload["environment"] = _options.Environment;

        var activity = Activity.Current;
        if (activity is not null)
        {
            payload["traceId"] = activity.TraceId.ToString();
            payload["spanId"] = activity.SpanId.ToString();
        }

        if (logEntry.EventId.Id != 0)
            payload["eventId"] = logEntry.EventId.Id;

        if (logEntry.Exception is not null)
            payload["exceptionType"] = logEntry.Exception.GetType().FullName;

        // Scope key/values (correlationId, business context …) — enables ELK correlation/filtering.
        scopeProvider?.ForEachScope(static (scope, state) => AppendPairs(scope, state), payload);

        // Structured arguments of the message template (e.g. StatusCode, ElapsedMs) as their own fields.
        AppendPairs(logEntry.State, payload);

        string line;
        try
        {
            line = JsonSerializer.Serialize(payload, JsonOptions);
        }
        catch (Exception ex)
        {
            // A logger must never throw. If a value in the payload cannot be serialized, fall back to a
            // minimal line that still carries the message so the original log is not lost.
            line = JsonSerializer.Serialize(new Dictionary<string, object?>
            {
                ["timestamp"] = payload["timestamp"],
                ["severity"] = payload["severity"],
                ["level"] = payload["level"],
                ["message"] = fullMessage,
                ["category"] = logEntry.Category,
                ["serializationError"] = ex.Message
            }, JsonOptions);
        }

        textWriter.WriteLine(line);
    }

    /// <summary>
    /// Serializes <see cref="Type"/> (and its runtime subtype <c>System.RuntimeType</c>) as its full
    /// name, since <see cref="JsonSerializer"/> otherwise throws <see cref="NotSupportedException"/>.
    /// </summary>
    private sealed class TypeJsonConverter : JsonConverter<Type>
    {
        public override bool CanConvert(Type typeToConvert) => typeof(Type).IsAssignableFrom(typeToConvert);

        public override Type Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            => throw new NotSupportedException();

        public override void Write(Utf8JsonWriter writer, Type value, JsonSerializerOptions options)
            => writer.WriteStringValue(value.FullName ?? value.Name);
    }

    private static void AppendPairs(object? source, Dictionary<string, object?> payload)
    {
        if (source is not IEnumerable<KeyValuePair<string, object>> pairs)
            return;

        foreach (var pair in pairs)
        {
            if (pair.Key == "{OriginalFormat}" || ReservedKeys.Contains(pair.Key))
                continue;
            payload[pair.Key] = pair.Value;
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
}

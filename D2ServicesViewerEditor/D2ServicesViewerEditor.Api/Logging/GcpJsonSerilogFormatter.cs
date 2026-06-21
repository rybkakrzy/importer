using System.Text.Json;
using Serilog.Events;
using Serilog.Formatting;

namespace D2ServicesViewerEditor.Api.Logging;

/// <summary>
/// Serilogowy formatter konsoli emitujący JSON zgodny z Google Cloud Logging (pole <c>severity</c>).
///
/// Powód: domyślny `WriteTo.Console()` Seriloga pisze zwykły tekst na stdout → Cloud Logging nadaje
/// każdemu wpisowi severity INFO, więc wyjątki (`Error`/`Fatal`) widać w Logs Explorer jako INFO.
/// Mapujemy <see cref="LogEventLevel"/> na severity GCP i dołączamy stack trace do treści (Error
/// Reporting grupuje błędy po stack trace w `message`).
/// </summary>
public sealed class GcpJsonSerilogFormatter : ITextFormatter
{
    private static readonly HashSet<string> ReservedKeys = new(StringComparer.Ordinal)
    {
        "severity", "level", "message", "category", "timestamp",
        "eventId", "exceptionType", "service", "environment", "traceId", "spanId"
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly string? _service;
    private readonly string? _environment;

    public GcpJsonSerilogFormatter(string? service = null, string? environment = null)
    {
        _service = service;
        _environment = environment;
    }

    public void Format(LogEvent logEvent, TextWriter output)
    {
        var message = logEvent.RenderMessage();
        var fullMessage = logEvent.Exception is null ? message : $"{message}\n{logEvent.Exception}";

        var payload = new Dictionary<string, object?>
        {
            ["severity"] = MapSeverity(logEvent.Level),
            ["level"] = logEvent.Level.ToString(),
            ["message"] = fullMessage,
            ["timestamp"] = logEvent.Timestamp.UtcDateTime.ToString("o"),
        };

        if (!string.IsNullOrWhiteSpace(_service))
            payload["service"] = _service;

        if (!string.IsNullOrWhiteSpace(_environment))
            payload["environment"] = _environment;

        if (logEvent.Properties.TryGetValue("SourceContext", out var source) && source is ScalarValue { Value: string ctx })
            payload["category"] = ctx;

        if (logEvent.Properties.TryGetValue("TraceId", out var traceId))
            payload["traceId"] = ToPlain(traceId);

        if (logEvent.Properties.TryGetValue("SpanId", out var spanId))
            payload["spanId"] = ToPlain(spanId);

        foreach (var property in logEvent.Properties)
        {
            if (property.Key == "SourceContext" || ReservedKeys.Contains(property.Key))
                continue;
            payload[property.Key] = ToPlain(property.Value);
        }

        if (logEvent.Exception is not null)
            payload["exceptionType"] = logEvent.Exception.GetType().FullName;

        output.WriteLine(JsonSerializer.Serialize(payload, JsonOptions));
    }

    private static object? ToPlain(LogEventPropertyValue value) => value switch
    {
        ScalarValue scalar => scalar.Value is Type t ? t.FullName ?? t.Name : scalar.Value,
        SequenceValue sequence => sequence.Elements.Select(ToPlain).ToArray(),
        StructureValue structure => structure.Properties.ToDictionary(
            p => p.Name,
            p => ToPlain(p.Value),
            StringComparer.Ordinal),
        DictionaryValue dictionary => dictionary.Elements.ToDictionary(
            e => ToPlainDictionaryKey(e.Key),
            e => ToPlain(e.Value),
            StringComparer.Ordinal),
        _ => value.ToString()
    };

    private static string ToPlainDictionaryKey(ScalarValue key)
        => key.Value?.ToString() ?? string.Empty;

    private static string MapSeverity(LogEventLevel level) => level switch
    {
        LogEventLevel.Verbose => "DEBUG",
        LogEventLevel.Debug => "DEBUG",
        LogEventLevel.Information => "INFO",
        LogEventLevel.Warning => "WARNING",
        LogEventLevel.Error => "ERROR",
        LogEventLevel.Fatal => "CRITICAL",
        _ => "DEFAULT"
    };
}

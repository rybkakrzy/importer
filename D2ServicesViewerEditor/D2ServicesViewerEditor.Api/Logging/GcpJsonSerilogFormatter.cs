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
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public void Format(LogEvent logEvent, TextWriter output)
    {
        var message = logEvent.RenderMessage();
        var fullMessage = logEvent.Exception is null ? message : $"{message}\n{logEvent.Exception}";

        var payload = new Dictionary<string, object?>
        {
            ["severity"] = MapSeverity(logEvent.Level),
            ["message"] = fullMessage,
            ["timestamp"] = logEvent.Timestamp.UtcDateTime.ToString("o"),
        };

        if (logEvent.Properties.TryGetValue("SourceContext", out var source) && source is ScalarValue { Value: string ctx })
            payload["category"] = ctx;

        if (logEvent.Exception is not null)
            payload["exceptionType"] = logEvent.Exception.GetType().FullName;

        output.WriteLine(JsonSerializer.Serialize(payload, JsonOptions));
    }

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

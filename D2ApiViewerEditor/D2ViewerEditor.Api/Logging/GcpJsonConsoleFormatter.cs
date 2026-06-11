using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Logging.Console;

namespace D2ViewerEditor.Api.Logging;

/// <summary>
/// Formatter konsoli emitujący logi jako jedną linię JSON zgodną z Google Cloud Logging.
///
/// Powód: domyślny formatter pisze zwykły tekst na stdout, a Cloud Logging przypisuje wtedy KAŻDEMU
/// wpisowi severity INFO (DEFAULT) — przez co wyjątki logowane `LogError`/`LogCritical` widać w Logs
/// Explorer jako INFO, a nie ERROR. Cloud Logging odczytuje pole <c>severity</c> ze strukturalnego
/// JSON-a, więc mapujemy <see cref="LogLevel"/> na severity GCP i wyjątek trafia do treści (Error
/// Reporting grupuje po stack trace w `message`).
/// </summary>
public sealed class GcpJsonConsoleFormatter : ConsoleFormatter
{
    public const string FormatterName = "gcp-json";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public GcpJsonConsoleFormatter() : base(FormatterName) { }

    public override void Write<TState>(
        in LogEntry<TState> logEntry, IExternalScopeProvider? scopeProvider, TextWriter textWriter)
    {
        var message = logEntry.Formatter?.Invoke(logEntry.State, logEntry.Exception) ?? string.Empty;
        if (string.IsNullOrEmpty(message) && logEntry.Exception is null)
            return;

        // Error Reporting w GCP grupuje błędy po stack trace zawartym w treści wpisu — dołączamy
        // pełny wyjątek do `message`, gdy istnieje.
        var fullMessage = logEntry.Exception is null
            ? message
            : $"{message}\n{logEntry.Exception}";

        var payload = new Dictionary<string, object?>
        {
            ["severity"] = MapSeverity(logEntry.LogLevel),
            ["message"] = fullMessage,
            ["category"] = logEntry.Category,
            ["timestamp"] = DateTimeOffset.UtcNow.ToString("o"),
        };

        if (logEntry.EventId.Id != 0)
            payload["eventId"] = logEntry.EventId.Id;

        if (logEntry.Exception is not null)
            payload["exceptionType"] = logEntry.Exception.GetType().FullName;

        textWriter.WriteLine(JsonSerializer.Serialize(payload, JsonOptions));
    }

    /// <summary>Mapowanie LogLevel → LogSeverity GCP (https://cloud.google.com/logging/docs/reference/v2/rest/v2/LogEntry#LogSeverity).</summary>
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

using D2ViewerEditor.Api.Logging;
using Microsoft.Extensions.Logging.Console;

namespace D2ViewerEditor.Api.Extensions;

public static class LoggingExtensions
{
    /// <summary>
    /// Włącza strukturalne logowanie JSON zgodne z Google Cloud Logging (pole `severity`), żeby
    /// wyjątki (`LogError`/`LogCritical`) były widoczne w Logs Explorer jako ERROR/CRITICAL, a nie
    /// INFO. Poza GCP/kontenerem (lokalny dev) zostaje czytelny domyślny formatter konsoli.
    /// Sterowanie jawne configiem `Logging:UseGcpFormat` (gdy brak — domyślnie poza Development).
    /// </summary>
    public static WebApplicationBuilder AddGcpStructuredLogging(this WebApplicationBuilder builder)
    {
        var explicitFlag = builder.Configuration.GetValue<bool?>("Logging:UseGcpFormat");
        var useGcp = explicitFlag ?? !builder.Environment.IsDevelopment();
        if (!useGcp)
            return builder;

        builder.Logging.AddConsole(options => options.FormatterName = GcpJsonConsoleFormatter.FormatterName);
        builder.Logging.AddConsoleFormatter<GcpJsonConsoleFormatter, ConsoleFormatterOptions>();
        return builder;
    }
}

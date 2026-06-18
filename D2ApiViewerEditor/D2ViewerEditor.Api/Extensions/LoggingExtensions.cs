using D2ViewerEditor.Api.Logging;

namespace D2ViewerEditor.Api.Extensions;

public static class LoggingExtensions
{
    /// <summary>
    /// Włącza strukturalne logowanie JSON (jedna linia/wpis na stdout) przyjazne dla ELK
    /// (Elasticsearch/Logstash/Kibana) oraz Google Cloud Logging: pole `severity` (GCP) + `level`,
    /// `service`, `environment`, `traceId`, pola ze scope'ów (correlationId, business context).
    /// Poza kontenerem (lokalny dev) zostaje czytelny domyślny formatter konsoli. Sterowanie jawne
    /// configiem `Logging:UseGcpFormat` (gdy brak — domyślnie poza Development).
    /// </summary>
    public static WebApplicationBuilder AddGcpStructuredLogging(this WebApplicationBuilder builder)
    {
        var explicitFlag = builder.Configuration.GetValue<bool?>("Logging:UseGcpFormat");
        var useGcp = explicitFlag ?? !builder.Environment.IsDevelopment();
        if (!useGcp)
            return builder;

        builder.Logging.AddConsole(options => options.FormatterName = GcpJsonConsoleFormatter.FormatterName);
        builder.Logging.AddConsoleFormatter<GcpJsonConsoleFormatter, StructuredLogFormatterOptions>(options =>
        {
            options.Service = builder.Environment.ApplicationName;
            options.Environment = builder.Environment.EnvironmentName;
        });
        return builder;
    }
}

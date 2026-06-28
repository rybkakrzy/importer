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

        // The formatter reads optional HTTP context (method/path/status) — works without it too.
        builder.Services.AddHttpContextAccessor();

        builder.Logging.AddConsole(options => options.FormatterName = GcpJsonConsoleFormatter.FormatterName);
        builder.Logging.AddConsoleFormatter<GcpJsonConsoleFormatter, StructuredLogFormatterOptions>(options =>
        {
            options.ServiceName = builder.Environment.ApplicationName;
            options.EnvironmentName = builder.Environment.EnvironmentName;
            options.ServiceVersion = ResolveServiceVersion(builder.Configuration);
            options.ProjectId = ResolveProjectId(builder.Configuration);
            options.ServiceInstanceId = ResolveInstanceId();
            // Back-compat aliases for any consumer still reading the flat fields.
            options.Service = options.ServiceName;
            options.Environment = options.EnvironmentName;
        });
        return builder;
    }

    private static string ResolveServiceVersion(IConfiguration configuration) =>
        configuration["BuildInfo:Version"]
        ?? System.Reflection.Assembly.GetEntryAssembly()?.GetName().Version?.ToString()
        ?? string.Empty;

    // GCP injects the project id as GOOGLE_CLOUD_PROJECT on Cloud Run/GKE; allow a config override.
    private static string ResolveProjectId(IConfiguration configuration) =>
        configuration["Logging:GcpProjectId"]
        ?? System.Environment.GetEnvironmentVariable("GOOGLE_CLOUD_PROJECT")
        ?? string.Empty;

    // Cloud Run sets K_REVISION; otherwise fall back to the container hostname.
    private static string ResolveInstanceId() =>
        System.Environment.GetEnvironmentVariable("K_REVISION")
        ?? System.Environment.GetEnvironmentVariable("HOSTNAME")
        ?? string.Empty;
    }
}

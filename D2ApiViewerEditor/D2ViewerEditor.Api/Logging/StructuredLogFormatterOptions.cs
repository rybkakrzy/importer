using Microsoft.Extensions.Logging.Console;

namespace D2ViewerEditor.Api.Logging;

// Options for GcpJsonConsoleFormatter. Carries the static enrichment fields (service/environment/
// version) plus toggles that select which field families are emitted (Google Cloud Logging, Elastic
// Common Schema, HTTP context, …). Bound from configuration in AddGcpStructuredLogging.
public sealed class StructuredLogFormatterOptions : ConsoleFormatterOptions
{
    // Back-compat aliases (kept so existing wiring/tests keep working). ServiceName/EnvironmentName
    // take precedence; these are the fallback when the canonical fields are not set.
    public string Service { get; set; } = string.Empty;
    public string Environment { get; set; } = string.Empty;

    public string ProjectId { get; set; } = string.Empty;
    public string ServiceName { get; set; } = string.Empty;
    public string ServiceVersion { get; set; } = string.Empty;
    public string EnvironmentName { get; set; } = string.Empty;
    public string ServiceInstanceId { get; set; } = string.Empty;

    public bool IncludeScopes { get; set; } = true;
    public bool IncludeEventId { get; set; } = true;
    public bool IncludeSourceLocation { get; set; }
    public bool IncludeHttpRequest { get; set; } = true;
    public bool IncludeElasticCommonSchemaFields { get; set; } = true;
    public bool IncludeGoogleCloudFields { get; set; } = true;

    // Property names (case/separator-insensitive) whose values are replaced with "[REDACTED]" wherever
    // they appear as a scope/state property or an http query parameter.
    public HashSet<string> RedactedPropertyNames { get; set; } = new(StringComparer.OrdinalIgnoreCase)
    {
        "password", "passwd", "secret", "token", "access_token", "refresh_token",
        "authorization", "cookie", "api_key", "apikey", "client_secret", "card_number"
    };

    public string ResolvedServiceName => string.IsNullOrEmpty(ServiceName) ? Service : ServiceName;
    public string ResolvedEnvironment => string.IsNullOrEmpty(EnvironmentName) ? Environment : EnvironmentName;
}

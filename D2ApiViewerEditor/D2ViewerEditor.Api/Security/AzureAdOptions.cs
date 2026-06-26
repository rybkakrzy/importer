namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Entra ID configuration for validating the API's access tokens, mapping claims and (when a
/// client secret is supplied via GCP Secret Manager) acquiring downstream tokens for Microsoft
/// Graph. Bound from the "AzureAd" configuration section; values are per-environment, supplied at
/// deploy time. The secret is never committed — empty means token-validation-only / local dev.
/// </summary>
public sealed class AzureAdOptions
{
    public const string SectionName = "AzureAd";

    /// <summary>e.g. https://login.microsoftonline.com/</summary>
    public string Instance { get; set; } = "https://login.microsoftonline.com/";

    public string TenantId { get; set; } = string.Empty;

    /// <summary>The API's application (client) id. Required by Microsoft.Identity.Web for token
    /// validation and downstream token acquisition (Graph).</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Expected audience — the API's application (client) id or api://{clientId}.</summary>
    public string Audience { get; set; } = string.Empty;

    /// <summary>App client secret (injected from GCP Secret Manager at deploy; empty = no Graph/local).</summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Claim type carrying the user's CorporateKey in the access token (configurable).</summary>
    public string CorporateKeyClaim { get; set; } = "corpKey";

    /// <summary>App role value for a standard application user.</summary>
    public string OperatorRole { get; set; } = "Operator";

    /// <summary>App role value for an application administrator.</summary>
    public string AdminRole { get; set; } = "Administrator";

    /// <summary>Delegated scopes requested for downstream calls (per environment, e.g. ["User.Read"]).
    /// Config surface for ecosystem parity (D2WebCore); empty = none.</summary>
    public string[] Scopes { get; set; } = [];

    /// <summary>Outbound HTTP proxy for reaching Entra / Graph in enterprise environments. Optional;
    /// empty <see cref="ProxyOptions.Url"/> = no proxy.</summary>
    public ProxyOptions Proxy { get; set; } = new();

    /// <summary>v2.0 authority built from Instance + TenantId.</summary>
    public string Authority => $"{Instance.TrimEnd('/')}/{TenantId}/v2.0";
}

/// <summary>Outbound proxy configuration for enterprise environments behind a forward proxy
/// (D2WebCore <c>BusinessProxy</c>). Explicit credentials are optional — empty username falls back to
/// the process default credentials.</summary>
public sealed class ProxyOptions
{
    /// <summary>Proxy URL, e.g. http://localhost:3128. Empty = no proxy.</summary>
    public string Url { get; set; } = string.Empty;

    /// <summary>Proxy username. Empty → use default credentials.</summary>
    public string Username { get; set; } = string.Empty;

    /// <summary>Proxy password (supply via secret store/env, not committed config).</summary>
    public string Password { get; set; } = string.Empty;
}

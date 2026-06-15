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
    public string CorporateKeyClaim { get; set; } = "ck";

    /// <summary>App role value for a standard application user.</summary>
    public string EmployeeRole { get; set; } = "APP_Pracownik";

    /// <summary>App role value for an application administrator.</summary>
    public string AdminRole { get; set; } = "APP_Admin";

    /// <summary>v2.0 authority built from Instance + TenantId.</summary>
    public string Authority => $"{Instance.TrimEnd('/')}/{TenantId}/v2.0";
}

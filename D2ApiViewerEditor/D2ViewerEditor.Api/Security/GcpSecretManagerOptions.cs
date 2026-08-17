namespace D2ViewerEditor.Api.Security;

/// <summary>
/// GCP Secret Manager configuration (Doc2 pattern). When enabled, the Entra client secret is
/// fetched at startup from <c>projects/{ProjectId}/secrets/{EntraSecretName}/versions/latest</c>
/// and injected into <c>AzureAd:ClientSecret</c>. Bound from the "GCPSecretManager" section.
/// Disabled by default so local/dev runs without GCP credentials.
/// </summary>
public sealed class GcpSecretManagerOptions
{
    public const string SectionName = "GCPSecretManager";

    public bool Enabled { get; set; }

    public string ProjectId { get; set; } = string.Empty;

    /// <summary>Doc2 naming: <c>SMC01{ENV}2_APP00404_entra_secret</c>.</summary>
    public string EntraSecretName { get; set; } = string.Empty;
}

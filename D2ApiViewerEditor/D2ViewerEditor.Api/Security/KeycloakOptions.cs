namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Keycloak (legacy) authentication configuration (Doc2 dual-auth). When enabled, a second
/// JWT bearer scheme validates Keycloak-issued tokens alongside Entra ID; the active scheme is
/// chosen per request by the token issuer. Disabled by default — Entra is the only scheme.
/// </summary>
public sealed class KeycloakOptions
{
    public const string SectionName = "Keycloak";

    public bool Enabled { get; set; }

    /// <summary>Keycloak realm issuer/authority, e.g. https://kc.example/realms/doc2.</summary>
    public string Authority { get; set; } = string.Empty;

    public string Audience { get; set; } = string.Empty;
}

namespace D2ViewerEditor.Application.Common.Security;

/// <summary>
/// Resolves the current user's identity for access decisions. Implementations:
/// claims-based (Entra ID access token) in the user-facing API, or a no-user system
/// provider in hosts without an interactive user (e.g. external ingest API).
/// </summary>
public interface ICurrentUserProvider
{
    /// <summary>CorporateKey of the user, or null when unavailable.</summary>
    string? CorporateKey { get; }

    /// <summary>True when the user holds the application admin role (Administrator).</summary>
    bool IsAdmin { get; }
}

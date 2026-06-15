namespace D2ViewerEditor.Application.Common.Security;

/// <summary>
/// No-user provider for hosts without an interactive user (e.g. the external ingest API,
/// which authenticates app-to-app, not per user). Keeps DI valid where the document access
/// guard is registered but no user identity exists. Never grants admin or a CorporateKey.
/// </summary>
public sealed class SystemCurrentUserProvider : ICurrentUserProvider
{
    public string? CorporateKey => null;
    public bool IsAdmin => false;
}

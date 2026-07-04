namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Names of the application authorization policies. Backend is the source of truth for
/// access — Angular guards are UX only.
/// </summary>
public static class AuthorizationPolicies
{
    /// <summary>Standard application access (Operator or Administrator). NOTE: this is an API-level
    /// authorization gate; frontend AREA access (which module a user may open) is governed separately
    /// by <see cref="ResourcesProvider"/> — an admin-only user is not shown editor/viewer, but the admin
    /// module still reads document/version data through these endpoints.</summary>
    public const string RequireAppOperator = "RequireAppOperator";

    /// <summary>Administrative module access (Administrator only).</summary>
    public const string RequireAppAdmin = "RequireAppAdmin";
}

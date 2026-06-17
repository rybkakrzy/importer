namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Names of the application authorization policies. Backend is the source of truth for
/// access — Angular guards are UX only.
/// </summary>
public static class AuthorizationPolicies
{
    /// <summary>Standard application access (Operator or Administrator).</summary>
    public const string RequireAppOperator = "RequireAppOperator";

    /// <summary>Administrative module access (Administrator only).</summary>
    public const string RequireAppAdmin = "RequireAppAdmin";
}

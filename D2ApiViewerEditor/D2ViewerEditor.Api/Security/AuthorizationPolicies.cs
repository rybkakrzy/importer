namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Names of the application authorization policies. Backend is the source of truth for
/// access — Angular guards are UX only.
/// </summary>
public static class AuthorizationPolicies
{
    /// <summary>Standard application access (APP_Pracownik or APP_Admin).</summary>
    public const string RequireAppEmployee = "RequireAppEmployee";

    /// <summary>Administrative module access (APP_Admin only).</summary>
    public const string RequireAppAdmin = "RequireAppAdmin";
}

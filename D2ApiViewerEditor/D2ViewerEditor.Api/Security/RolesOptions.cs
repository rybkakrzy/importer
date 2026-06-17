namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Group→role mapping configuration (Doc2 / D2WebCore pattern), bound from the "Roles" section.
/// Entra ID emits the user's group memberships in the <c>groups</c> claim; the claims transformer
/// (<see cref="Doc2ClaimsTransformer"/>) maps those group identifiers to application role names
/// used by the authorization policies. Group identifiers and role names are environment-specific
/// and supplied via configuration (no secrets — only group/role names).
/// </summary>
public sealed class RolesOptions
{
    public const string SectionName = "Roles";

    /// <summary>Optional common prefix of the application's AD groups (informational/filtering).</summary>
    public string GroupPrefix { get; set; } = string.Empty;

    /// <summary>Each application role and the AD groups (names or object-ids) that grant it.</summary>
    public List<RoleGroupMapping> Roles { get; set; } = new();
}

/// <summary>One application role and the Entra ID groups that map onto it.</summary>
public sealed class RoleGroupMapping
{
    /// <summary>Application role name added as a role claim (e.g. <c>Administrator</c>, <c>Operator</c>).</summary>
    public string RoleName { get; set; } = string.Empty;

    /// <summary>AD group identifiers (display names or object ids) whose membership grants <see cref="RoleName"/>.</summary>
    public List<string> GroupNames { get; set; } = new();
}

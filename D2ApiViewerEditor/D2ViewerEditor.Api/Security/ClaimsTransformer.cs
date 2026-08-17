using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Maps Entra ID group memberships (the <c>groups</c> claim) to application role claims
/// (Doc2 / D2WebCore pattern). Runs on every authenticated request and is idempotent.
/// Native Entra <c>app role</c> claims (also surfaced as <c>roles</c>) are preserved, so
/// App Roles and group→role mapping can coexist during migration.
/// <para>
/// LIMITATION — group overage: when a user is a member of &gt; ~200 groups Entra omits the
/// <c>groups</c> claim and emits <c>_claim_names</c>/<c>_claim_sources</c> pointing at Graph
/// instead. This transformer does NOT resolve overage via Graph, so such users get no
/// group-derived roles — rely on App Roles (which coexist) for them, or add Graph overage
/// resolution later. Also note: <c>groups</c> carries object-ids by default; configure the
/// "groups" optional claim (cloud display name / on-prem sAMAccountName) or put object-ids in
/// <see cref="RolesOptions"/> so the configured GroupNames match the token's values.
/// </para>
/// </summary>
public sealed class ClaimsTransformer : IClaimsTransformation
{
    private const string GroupsClaimType = "groups";

    /// <summary>Must match the JwtBearer <c>RoleClaimType</c> configured in Program.cs.</summary>
    private const string RolesClaimType = "roles";

    private readonly RolesOptions _roles;

    public ClaimsTransformer(IOptions<RolesOptions> roles) => _roles = roles.Value;

    public Task<ClaimsPrincipal> TransformAsync(ClaimsPrincipal principal)
    {
        if (principal.Identity is not ClaimsIdentity identity || !identity.IsAuthenticated)
            return Task.FromResult(principal);

        if (_roles.Roles.Count == 0)
            return Task.FromResult(principal);

        // Match against group/role identifiers regardless of which claim type Entra/federation used
        // (D2WebCore pattern): `groups`, `roles`/native app roles, the WS-Fed role URI.
        var candidateValues = principal.Claims
            .Where(c => c.Type == GroupsClaimType
                     || c.Type == RolesClaimType
                     || c.Type == ClaimTypes.Role
                     || c.Type.Contains("identity/claims/role", StringComparison.OrdinalIgnoreCase))
            .Select(c => c.Value.Trim())
            .Where(v => v.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (candidateValues.Count == 0)
            return Task.FromResult(principal);

        foreach (var mapping in _roles.Roles)
        {
            if (string.IsNullOrWhiteSpace(mapping.RoleName))
                continue;

            var granted = mapping.GroupNames?.Any(g => candidateValues.Contains(g.Trim())) is true;
            if (!granted)
                continue;

            // Idempotent: add the role claim only if the principal does not already carry it
            // (the transformer can run more than once per request).
            if (!principal.HasClaim(RolesClaimType, mapping.RoleName))
                identity.AddClaim(new Claim(RolesClaimType, mapping.RoleName));
        }

        return Task.FromResult(principal);
    }
}

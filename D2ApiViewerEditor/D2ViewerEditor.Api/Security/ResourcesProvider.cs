using System.Security.Claims;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Maps the signed-in user's application roles to the set of frontend resources (route names) they
/// may access. Backend is the source of truth; the Angular resource guard calls
/// <c>GET /api/identity/resources</c> and gates navigation by route path. Mirrors the D2WebCore
/// <c>ResourcesProvider</c> pattern.
/// <para>Access matrix (roles are NOT nested — Operator and Administrator are disjoint areas):</para>
/// <list type="bullet">
///   <item>Operator → dashboard, editor, viewer (NOT admin)</item>
///   <item>Administrator → admin (NOT dashboard/editor/viewer)</item>
///   <item>both roles → union of the above</item>
///   <item>no application role → empty (no access anywhere)</item>
/// </list>
/// </summary>
public sealed class ResourcesProvider
{
    private readonly AzureAdOptions _options;

    public ResourcesProvider(IOptions<AzureAdOptions> options) => _options = options.Value;

    /// <summary>Operator resources: home dashboard + DOCX editor + PDF viewer route names.</summary>
    public static readonly IReadOnlyList<string> OperatorResources = ["dashboard", "editor", "viewer"];

    /// <summary>Admin module route name.</summary>
    public const string AdminResource = "admin";

    /// <summary>
    /// Resource names the principal may access, derived from its app roles (claim "roles").
    /// See the access matrix on the class. No app role → empty list.
    /// </summary>
    public IReadOnlyList<string> GetForUser(ClaimsPrincipal user)
    {
        var resources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (user.IsInRole(_options.OperatorRole))
        {
            foreach (var resource in OperatorResources)
            {
                resources.Add(resource);
            }
        }

        if (user.IsInRole(_options.AdminRole))
        {
            resources.Add(AdminResource);
        }

        return resources.ToList();
    }
}

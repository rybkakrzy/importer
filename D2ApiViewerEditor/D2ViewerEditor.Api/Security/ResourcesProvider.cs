using System.Security.Claims;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Maps the signed-in user's application roles to the set of frontend resources (route names) they
/// may access. Backend is the source of truth; the Angular resource guard calls
/// <c>GET /api/identity/resources</c> and gates navigation by route path. Mirrors the D2WebCore
/// <c>ResourcesProvider</c> pattern. Admin is a superset of the document resources.
/// </summary>
public sealed class ResourcesProvider
{
    private readonly AzureAdOptions _options;

    public ResourcesProvider(IOptions<AzureAdOptions> options) => _options = options.Value;

    /// <summary>Document resources: DOCX editor + PDF viewer route names.</summary>
    public static readonly IReadOnlyList<string> DocumentResources = ["editor", "viewer"];

    /// <summary>Admin module route name.</summary>
    public const string AdminResource = "admin";

    /// <summary>
    /// Resource names the principal may access, derived from its app roles (claim "roles").
    /// Operator → documents; Administrator → documents + admin (superset). No app role → empty.
    /// </summary>
    public IReadOnlyList<string> GetForUser(ClaimsPrincipal user)
    {
        var resources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var isAdmin = user.IsInRole(_options.AdminRole);
        var isOperator = user.IsInRole(_options.OperatorRole);

        if (isOperator || isAdmin)
        {
            foreach (var resource in DocumentResources)
            {
                resources.Add(resource);
            }
        }

        if (isAdmin)
        {
            resources.Add(AdminResource);
        }

        return resources.ToList();
    }
}

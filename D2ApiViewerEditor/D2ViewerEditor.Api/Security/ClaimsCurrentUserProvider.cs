using System.Security.Claims;
using D2ViewerEditor.Application.Common.Security;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Production identity provider: resolves CorporateKey and admin role from the validated
/// Entra ID access token (ClaimsPrincipal). CorporateKey comes from a single configurable
/// claim; missing/blank → null (restricted documents then deny). Admin is derived from the
/// app role claim ("roles").
/// </summary>
public sealed class ClaimsCurrentUserProvider : ICurrentUserProvider
{
    private readonly IHttpContextAccessor _httpContextAccessor;
    private readonly AzureAdOptions _options;

    public ClaimsCurrentUserProvider(IHttpContextAccessor httpContextAccessor, IOptions<AzureAdOptions> options)
    {
        _httpContextAccessor = httpContextAccessor;
        _options = options.Value;
    }

    private ClaimsPrincipal? User => _httpContextAccessor.HttpContext?.User;

    public string? CorporateKey
    {
        get
        {
            var value = User?.FindFirst(_options.CorporateKeyClaim)?.Value;
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }

    // Relies on JwtBearer RoleClaimType = "roles" (Entra app roles) configured in Program.cs.
    public bool IsAdmin => User?.IsInRole(_options.AdminRole) ?? false;
}

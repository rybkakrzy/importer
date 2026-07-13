using System.Security.Claims;
using D2ViewerEditor.Application.Common.Security;
using Microsoft.Extensions.Logging;
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
    private readonly ILogger<ClaimsCurrentUserProvider> _logger;

    public ClaimsCurrentUserProvider(
        IHttpContextAccessor httpContextAccessor,
        IOptions<AzureAdOptions> options,
        ILogger<ClaimsCurrentUserProvider> logger)
    {
        _httpContextAccessor = httpContextAccessor;
        _options = options.Value;
        _logger = logger;
    }

    private ClaimsPrincipal? User => _httpContextAccessor.HttpContext?.User;

    public string? CorporateKey
    {
        get
        {
            var user = User;
            // UWAGA: Microsoft.Identity.Web używa CaseSensitiveClaimsIdentity, więc
            // ClaimsPrincipal.FindFirst("corpKey") NIE dopasuje claimu "corpkey". Token z Qutasator-AD
            // niesie klucz pod nazwą "corpkey" — szukamy więc po Type bez rozróżniania wielkości liter,
            // żeby nazwa w konfiguracji ("corpKey"/"corpkey"/"CorpKey") nie decydowała o sukcesie.
            var value = user?.Claims
                .FirstOrDefault(c => string.Equals(c.Type, _options.CorporateKeyClaim, StringComparison.OrdinalIgnoreCase))
                ?.Value;
            if (string.IsNullOrWhiteSpace(value))
            {
                // DIAGNOSTYKA (tymczasowe): gdy oczekiwanego claimu brak, wypisujemy NAZWY claimów
                // obecnych w tokenie (bez wartości — bez PII), by ustalić właściwą nazwę i ustawić
                // AzureAd:CorporateKeyClaim. Po znalezieniu nazwy ten log można usunąć.
                var claimTypes = user?.Claims.Select(c => c.Type).Distinct() ?? [];
                _logger.LogWarning(
                    "CorporateKey claim '{ExpectedClaim}' nieobecny lub pusty. IsAuthenticated={IsAuth}. Dostępne typy claimów: {ClaimTypes}",
                    _options.CorporateKeyClaim,
                    user?.Identity?.IsAuthenticated ?? false,
                    string.Join(", ", claimTypes));
                return null;
            }

            return value.Trim();
        }
    }

    // Relies on JwtBearer RoleClaimType = "roles" (Entra app roles) configured in Program.cs.
    public bool IsAdmin => User?.IsInRole(_options.AdminRole) ?? false;
}

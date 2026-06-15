using System.IdentityModel.Tokens.Jwt;
using Microsoft.AspNetCore.Authentication.JwtBearer;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Authentication scheme names + issuer-based scheme selection for the Entra ID / Keycloak
/// dual-auth setup (Doc2). The composite scheme forwards each request to the right JWT validator
/// based on the bearer token's issuer.
/// </summary>
public static class AuthSchemes
{
    /// <summary>Entra ID scheme (Microsoft.Identity.Web registers it under the default "Bearer").</summary>
    public const string Entra = JwtBearerDefaults.AuthenticationScheme;

    public const string Keycloak = "Keycloak";

    /// <summary>Composite policy scheme used as the default; forwards by token issuer.</summary>
    public const string EntraOrKeycloak = "EntraOrKeycloak";

    /// <summary>
    /// Returns <see cref="Keycloak"/> when the bearer token's <c>iss</c> matches the configured
    /// Keycloak authority, otherwise <see cref="Entra"/> (also the fallback for missing/unreadable tokens).
    /// </summary>
    public static string SelectByIssuer(HttpContext context, string keycloakAuthority)
    {
        var authorization = context.Request.Headers.Authorization.ToString();
        if (authorization.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            var token = authorization["Bearer ".Length..].Trim();
            try
            {
                var handler = new JwtSecurityTokenHandler();
                if (handler.CanReadToken(token))
                {
                    var issuer = handler.ReadJwtToken(token).Issuer;
                    if (!string.IsNullOrWhiteSpace(keycloakAuthority)
                        && issuer.StartsWith(keycloakAuthority.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    {
                        return Keycloak;
                    }
                }
            }
            catch
            {
                // Unreadable/garbage token → fall through to Entra, whose validator rejects it (401).
            }
        }

        return Entra;
    }
}

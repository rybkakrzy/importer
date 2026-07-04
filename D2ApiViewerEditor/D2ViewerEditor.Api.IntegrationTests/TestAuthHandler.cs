using System.Security.Claims;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.IntegrationTests;

/// <summary>
/// Test authentication scheme driven by request headers, so a single host can exercise every access
/// path against the REAL authorization policies (RequireAppOperator / RequireAppAdmin):
/// <list type="bullet">
/// <item>no <c>X-Test-Auth</c> header → <see cref="AuthenticateResult.NoResult"/> → challenge → 401.</item>
/// <item><c>X-Test-Auth: true</c> with no <c>X-Test-Roles</c> → authenticated principal with an empty
/// role set → policies that require a role → 403.</item>
/// <item><c>X-Test-Roles: Operator</c> (comma-separated) → roles surfaced in the <c>roles</c> claim,
/// matching the JwtBearer <c>RoleClaimType</c> the app configures.</item>
/// </list>
/// </summary>
public sealed class TestAuthHandler : AuthenticationHandler<AuthenticationSchemeOptions>
{
    public const string SchemeName = "Test";
    public const string AuthHeader = "X-Test-Auth";
    public const string RolesHeader = "X-Test-Roles";

    /// <summary>Must match the RoleClaimType the API configures for JwtBearer (see ConfigureAuthentication).</summary>
    private const string RolesClaimType = "roles";

    public TestAuthHandler(
        IOptionsMonitor<AuthenticationSchemeOptions> options,
        ILoggerFactory logger,
        UrlEncoder encoder)
        : base(options, logger, encoder)
    {
    }

    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        if (!Request.Headers.ContainsKey(AuthHeader))
            return Task.FromResult(AuthenticateResult.NoResult());

        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, "test-user"),
            new("name", "Test User"),
        };

        if (Request.Headers.TryGetValue(RolesHeader, out var roles))
        {
            foreach (var role in roles.ToString()
                         .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                claims.Add(new Claim(RolesClaimType, role));
            }
        }

        var identity = new ClaimsIdentity(claims, SchemeName, nameType: "name", roleType: RolesClaimType);
        var ticket = new AuthenticationTicket(new ClaimsPrincipal(identity), SchemeName);
        return Task.FromResult(AuthenticateResult.Success(ticket));
    }
}

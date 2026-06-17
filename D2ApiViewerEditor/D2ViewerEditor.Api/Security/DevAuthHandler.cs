using System.Security.Claims;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// LOCAL-DEV ONLY authentication handler. Auto-authenticates every request as a fake user with both
/// app roles, so `[Authorize]` / RequireAppAdmin pass without a real Entra tenant. Enabled exclusively
/// via `Auth:DevBypass=true` AND a non-Production environment (see Program.cs) — never on a server.
/// </summary>
public sealed class DevAuthHandler : AuthenticationHandler<AuthenticationSchemeOptions>
{
    public const string SchemeName = "DevAuth";

    public DevAuthHandler(
        IOptionsMonitor<AuthenticationSchemeOptions> options,
        ILoggerFactory logger,
        UrlEncoder encoder)
        : base(options, logger, encoder)
    {
    }

    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        var claims = new[]
        {
            new Claim(ClaimTypes.NameIdentifier, "dev-local"),
            new Claim("name", "Local Dev"),
            new Claim("roles", "Operator"),
            new Claim("roles", "Administrator"),
        };
        var identity = new ClaimsIdentity(claims, SchemeName, nameType: "name", roleType: "roles");
        var ticket = new AuthenticationTicket(new ClaimsPrincipal(identity), SchemeName);
        return Task.FromResult(AuthenticateResult.Success(ticket));
    }
}

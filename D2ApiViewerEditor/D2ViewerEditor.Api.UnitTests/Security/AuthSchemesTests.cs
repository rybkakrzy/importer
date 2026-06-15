using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Security;

/// <summary>
/// Entra ID / Keycloak dual-auth: the composite scheme picks the validator by the bearer token's
/// issuer. Entra is the safe fallback for missing/garbage tokens (its validator then returns 401).
/// </summary>
[TestFixture]
public class AuthSchemesTests
{
    private const string KeycloakAuthority = "https://kc.example/realms/doc2";

    private static HttpContext ContextWithIssuer(string issuer)
    {
        var token = new JwtSecurityToken(issuer: issuer, audience: null,
            claims: Array.Empty<Claim>(), notBefore: null, expires: null);
        var jwt = new JwtSecurityTokenHandler().WriteToken(token);
        var ctx = new DefaultHttpContext();
        ctx.Request.Headers.Authorization = $"Bearer {jwt}";
        return ctx;
    }

    [Test]
    public void Selects_keycloak_for_keycloak_issuer()
    {
        AuthSchemes.SelectByIssuer(ContextWithIssuer(KeycloakAuthority), KeycloakAuthority)
            .Should().Be(AuthSchemes.Keycloak);
    }

    [Test]
    public void Matches_keycloak_authority_ignoring_trailing_slash()
    {
        AuthSchemes.SelectByIssuer(ContextWithIssuer(KeycloakAuthority), KeycloakAuthority + "/")
            .Should().Be(AuthSchemes.Keycloak);
    }

    [Test]
    public void Selects_entra_for_entra_issuer()
    {
        AuthSchemes.SelectByIssuer(ContextWithIssuer("https://login.microsoftonline.com/abc/v2.0"), KeycloakAuthority)
            .Should().Be(AuthSchemes.Entra);
    }

    [Test]
    public void Selects_entra_when_no_authorization_header()
    {
        AuthSchemes.SelectByIssuer(new DefaultHttpContext(), KeycloakAuthority).Should().Be(AuthSchemes.Entra);
    }

    [Test]
    public void Selects_entra_for_malformed_token()
    {
        var ctx = new DefaultHttpContext();
        ctx.Request.Headers.Authorization = "Bearer not-a-jwt";
        AuthSchemes.SelectByIssuer(ctx, KeycloakAuthority).Should().Be(AuthSchemes.Entra);
    }
}

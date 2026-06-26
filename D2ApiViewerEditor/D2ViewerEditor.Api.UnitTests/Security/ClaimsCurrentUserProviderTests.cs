using System.Security.Claims;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Security;

/// <summary>
/// Produkcyjny dostawca tożsamości: CorporateKey i rola admina pochodzą WYŁĄCZNIE ze
/// zweryfikowanego tokenu Entra (ClaimsPrincipal), nie z body żądania. Te testy pilnują
/// rozdziału warstw: brak claimu <c>corpKey</c> daje <c>null</c> (reguła biznesowa → 400
/// w handlerze), a NIE odrzucenie uwierzytelnienia (401 powstaje w pipeline JwtBearer).
/// </summary>
[TestFixture]
public class ClaimsCurrentUserProviderTests
{
    private const string CorpKeyClaim = "corpKey";

    private static ClaimsCurrentUserProvider ProviderWith(ClaimsPrincipal user)
    {
        var accessor = new HttpContextAccessor
        {
            HttpContext = new DefaultHttpContext { User = user }
        };
        var options = Options.Create(new AzureAdOptions
        {
            CorporateKeyClaim = CorpKeyClaim,
            AdminRole = "Administrator"
        });
        return new ClaimsCurrentUserProvider(accessor, options, NullLogger<ClaimsCurrentUserProvider>.Instance);
    }

    private static ClaimsPrincipal AuthenticatedUser(params Claim[] claims)
    {
        var identity = new ClaimsIdentity(claims, authenticationType: "TestAuth", nameType: "name", roleType: "roles");
        return new ClaimsPrincipal(identity);
    }

    [Test]
    public void CorporateKey_ReadsConfiguredClaim_FromValidatedToken()
    {
        var provider = ProviderWith(AuthenticatedUser(new Claim(CorpKeyClaim, "CORP-42")));

        provider.CorporateKey.Should().Be("CORP-42");
    }

    [Test]
    public void CorporateKey_MatchesClaimCaseInsensitively()
    {
        // Realny przypadek ING-AD: token niesie claim "corpkey" (małe litery), konfiguracja to "corpKey".
        // CaseSensitiveClaimsIdentity (Microsoft.Identity.Web) nie dopasowałby tego przez FindFirst —
        // dlatego provider szuka po Type z OrdinalIgnoreCase. Bez tego CorporateKey jest null → 400.
        var provider = ProviderWith(AuthenticatedUser(new Claim("corpkey", "XI81XE")));

        provider.CorporateKey.Should().Be("XI81XE");
    }

    [Test]
    public void CorporateKey_TrimsSurroundingWhitespace()
    {
        var provider = ProviderWith(AuthenticatedUser(new Claim(CorpKeyClaim, "  CORP-42  ")));

        provider.CorporateKey.Should().Be("CORP-42");
    }

    [Test]
    public void CorporateKey_ReturnsNull_WhenClaimMissing()
    {
        // Token ważny, ale bez claimu corpKey → null. Handler zwróci Result.Failure (400),
        // nie wolno tego mylić z 401 (brak/nieważny token).
        var provider = ProviderWith(AuthenticatedUser(new Claim("name", "Jan Kowalski")));

        provider.CorporateKey.Should().BeNull();
    }

    [Test]
    public void CorporateKey_ReturnsNull_WhenClaimIsBlank()
    {
        var provider = ProviderWith(AuthenticatedUser(new Claim(CorpKeyClaim, "   ")));

        provider.CorporateKey.Should().BeNull();
    }

    [Test]
    public void CorporateKey_ReturnsNull_WhenNoHttpContext()
    {
        var provider = new ClaimsCurrentUserProvider(
            new HttpContextAccessor { HttpContext = null },
            Options.Create(new AzureAdOptions { CorporateKeyClaim = CorpKeyClaim }),
            NullLogger<ClaimsCurrentUserProvider>.Instance);

        provider.CorporateKey.Should().BeNull();
    }

    [Test]
    public void IsAdmin_True_WhenAdminRoleClaimPresent()
    {
        var provider = ProviderWith(AuthenticatedUser(new Claim("roles", "Administrator")));

        provider.IsAdmin.Should().BeTrue();
    }

    [Test]
    public void IsAdmin_False_WhenOnlyNonAdminRole()
    {
        var provider = ProviderWith(AuthenticatedUser(new Claim("roles", "Operator")));

        provider.IsAdmin.Should().BeFalse();
    }
}

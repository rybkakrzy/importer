using D2ViewerEditor.Application.Common.Security;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common.Security;

[TestFixture]
public class DocumentAccessGuardTests
{
    private sealed class StubCurrentUser : ICurrentUserProvider
    {
        public string? CorporateKey { get; init; }
        public bool IsAdmin { get; init; }
    }

    private static DocumentAccessGuard Guard(string? userKey, bool isAdmin = false) =>
        new(new StubCurrentUser { CorporateKey = userKey, IsAdmin = isAdmin });

    [Test]
    public void NoMetadata_IsPublic()
    {
        Guard(null).IsViewAllowed(null).Should().BeTrue();
        Guard(null).IsViewAllowed("   ").Should().BeTrue();
    }

    [Test]
    public void MetadataWithoutAllowList_IsPublic()
    {
        Guard(null).IsViewAllowed("{\"returnUrl\":\"https://x\",\"classification\":\"C2\"}").Should().BeTrue();
    }

    [Test]
    public void NullOrEmptyAllowList_IsPublic()
    {
        Guard(null).IsViewAllowed("{\"allowedCorporateKeys\":null}").Should().BeTrue();
        Guard(null).IsViewAllowed("{\"allowedCorporateKeys\":[]}").Should().BeTrue();
    }

    [Test]
    public void RestrictedList_UserOnList_IsAllowed()
    {
        Guard("CK-2").IsViewAllowed("{\"allowedCorporateKeys\":[\"CK-1\",\"CK-2\"]}").Should().BeTrue();
    }

    [Test]
    public void RestrictedList_UserNotOnList_IsDenied()
    {
        Guard("CK-9").IsViewAllowed("{\"allowedCorporateKeys\":[\"CK-1\",\"CK-2\"]}").Should().BeFalse();
    }

    [Test]
    public void RestrictedList_NoUserKey_IsDenied()
    {
        Guard(null).IsViewAllowed("{\"allowedCorporateKeys\":[\"CK-1\"]}").Should().BeFalse();
    }

    [Test]
    public void RestrictedList_TrimAndCaseInsensitive()
    {
        Guard(" ck-1 ").IsViewAllowed("{\"allowedCorporateKeys\":[\"CK-1\"]}").Should().BeTrue();
    }

    [Test]
    public void MalformedMetadata_DoesNotThrow_TreatedAsPublic()
    {
        Guard(null).IsViewAllowed("{ not valid json ").Should().BeTrue();
    }

    [Test]
    public void Admin_BypassesRestrictedAllowList()
    {
        // APP_Admin sees every document regardless of allowedCorporateKeys (business decision).
        Guard(userKey: null, isAdmin: true)
            .IsViewAllowed("{\"allowedCorporateKeys\":[\"CK-1\",\"CK-2\"]}")
            .Should().BeTrue();
    }
}

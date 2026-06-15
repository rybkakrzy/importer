using D2ViewerEditor.Domain.Security;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Security;

[TestFixture]
public class DocumentAccessPolicyTests
{
    [Test]
    public void NullAllowList_IsPublic()
    {
        DocumentAccessPolicy.IsViewAllowed(null, null).Should().BeTrue();
        DocumentAccessPolicy.IsViewAllowed(null, "anyone").Should().BeTrue();
    }

    [Test]
    public void EmptyAllowList_IsPublic()
    {
        DocumentAccessPolicy.IsViewAllowed(Array.Empty<string>(), null).Should().BeTrue();
        DocumentAccessPolicy.IsViewAllowed(new[] { "  ", "" }, "anyone").Should().BeTrue(
            "blank-only entries are treated as no list");
    }

    [Test]
    public void RestrictedList_UserOnList_IsAllowed()
    {
        DocumentAccessPolicy.IsViewAllowed(new[] { "CK-1", "CK-2" }, "CK-2").Should().BeTrue();
    }

    [Test]
    public void RestrictedList_UserNotOnList_IsDenied()
    {
        DocumentAccessPolicy.IsViewAllowed(new[] { "CK-1", "CK-2" }, "CK-9").Should().BeFalse();
    }

    [Test]
    public void RestrictedList_NoUserKey_IsDenied()
    {
        DocumentAccessPolicy.IsViewAllowed(new[] { "CK-1" }, null).Should().BeFalse();
        DocumentAccessPolicy.IsViewAllowed(new[] { "CK-1" }, "   ").Should().BeFalse();
    }

    [Test]
    public void Comparison_IsTrimAndCaseInsensitive()
    {
        DocumentAccessPolicy.IsViewAllowed(new[] { "  CK-1  " }, "ck-1").Should().BeTrue();
        DocumentAccessPolicy.IsViewAllowed(new[] { "CK-1" }, "  CK-1 ").Should().BeTrue();
    }
}

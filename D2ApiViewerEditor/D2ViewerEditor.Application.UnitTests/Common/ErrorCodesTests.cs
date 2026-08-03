using D2ViewerEditor.Application.Common;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common;

/// <summary>
/// Kody są kontraktem z GUI (rozpoznaje przypadek po `code`, nie po treści komunikatu),
/// więc ich WARTOŚCI muszą przetrwać każdą zmianę nazw stałych po stronie backendu.
/// </summary>
[TestFixture]
public class ErrorCodesTests
{
    [TestCase("DOCUMENT_CONTENT_EMPTY")]
    [TestCase("PASSWORD_REQUIRED")]
    [TestCase("WRONG_PASSWORD")]
    [TestCase("UNSUPPORTED_LEGACY_DOC")]
    public void IsKnown_ForPublishedCode_ReturnsTrue(string code)
    {
        ErrorCodes.IsKnown(code).Should().BeTrue();
    }

    [Test]
    public void PublishedCodes_KeepTheirWireValues()
    {
        ErrorCodes.DocumentContentEmpty.Should().Be("DOCUMENT_CONTENT_EMPTY");
        ErrorCodes.DocumentProtected.Should().Be("PASSWORD_REQUIRED");
        ErrorCodes.DocumentUnlockFailed.Should().Be("WRONG_PASSWORD");
        ErrorCodes.UnsupportedLegacyDoc.Should().Be("UNSUPPORTED_LEGACY_DOC");
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase("password_required")]
    [TestCase("NotAValidationCode")]
    public void IsKnown_ForUnknownOrMiscasedCode_ReturnsFalse(string? code)
    {
        ErrorCodes.IsKnown(code).Should().BeFalse();
    }
}

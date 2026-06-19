using D2ViewerEditor.Application.Features.Documents.Common;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Common;

[TestFixture]
public class ExternalDocumentMetadataTests
{
    [Test]
    public void Parse_NullOrWhitespace_ReturnsEmpty()
    {
        ExternalDocumentMetadata.Parse(null).Should().Be(ExternalDocumentMetadata.Empty);
        ExternalDocumentMetadata.Parse("").Should().Be(ExternalDocumentMetadata.Empty);
        ExternalDocumentMetadata.Parse("   ").Should().Be(ExternalDocumentMetadata.Empty);
    }

    [Test]
    public void Parse_MalformedJson_ReturnsEmptyWithoutThrowing()
    {
        var act = () => ExternalDocumentMetadata.Parse("{ this is not json ");

        act.Should().NotThrow();
        act().Should().Be(ExternalDocumentMetadata.Empty);
    }

    [Test]
    public void Parse_JsonNullLiteral_ReturnsEmpty()
    {
        ExternalDocumentMetadata.Parse("null").Should().Be(ExternalDocumentMetadata.Empty);
    }

    [Test]
    public void Parse_AllFieldsPresent_MapsEveryProperty()
    {
        const string json =
            "{\"returnUrl\":\"https://app.example.com/cb\",\"classification\":\"C2\",\"userDownload\":true}";

        var result = ExternalDocumentMetadata.Parse(json);

        result.ReturnUrl.Should().Be("https://app.example.com/cb");
        result.Classification.Should().Be("C2");
        result.UserDownload.Should().BeTrue();
    }

    [Test]
    public void Parse_PartialJson_LeavesMissingFieldsNull()
    {
        var result = ExternalDocumentMetadata.Parse("{\"classification\":\"C1\"}");

        result.ReturnUrl.Should().BeNull();
        result.Classification.Should().Be("C1");
        result.UserDownload.Should().BeNull();
    }

    [Test]
    public void Parse_IsCaseInsensitiveForPropertyNames()
    {
        // JsonSerializerDefaults.Web => camelCase + case-insensitive matching.
        var result = ExternalDocumentMetadata.Parse("{\"ReturnUrl\":\"https://x/y\",\"UserDownload\":true}");

        result.ReturnUrl.Should().Be("https://x/y");
        result.UserDownload.Should().BeTrue();
    }

    [TestCase(true, true, TestName = "explicit true allows download")]
    [TestCase(false, false, TestName = "explicit false forbids download")]
    [TestCase(null, false, TestName = "missing flag forbids download")]
    public void IsUserDownloadAllowed_OnlyExplicitTrueAllows(bool? flag, bool expected)
    {
        var metadata = new ExternalDocumentMetadata(null, null, flag);

        metadata.IsUserDownloadAllowed.Should().Be(expected);
    }

    [TestCase(true, true, TestName = "explicit true keeps save-state visible")]
    [TestCase(null, true, TestName = "missing flag keeps save-state visible (default)")]
    [TestCase(false, false, TestName = "explicit false hides save-state")]
    public void IsSaveStateVisible_HiddenOnlyOnExplicitFalse(bool? flag, bool expected)
    {
        var metadata = new ExternalDocumentMetadata(null, null, null, flag);

        metadata.IsSaveStateVisible.Should().Be(expected);
    }

    [Test]
    public void Parse_ShowSaveStateFalse_IsParsedAndHidesSaveState()
    {
        var result = ExternalDocumentMetadata.Parse("{\"showSaveState\":false}");

        result.ShowSaveState.Should().BeFalse();
        result.IsSaveStateVisible.Should().BeFalse();
    }

    [Test]
    public void Parse_ShowSaveStateAbsent_DefaultsToVisible()
    {
        ExternalDocumentMetadata.Parse("{\"classification\":\"C2\"}").IsSaveStateVisible.Should().BeTrue();
    }

    [Test]
    public void Empty_HasAllNullFieldsAndForbidsDownload()
    {
        ExternalDocumentMetadata.Empty.ReturnUrl.Should().BeNull();
        ExternalDocumentMetadata.Empty.Classification.Should().BeNull();
        ExternalDocumentMetadata.Empty.UserDownload.Should().BeNull();
        ExternalDocumentMetadata.Empty.IsUserDownloadAllowed.Should().BeFalse();
        ExternalDocumentMetadata.Empty.ShowSaveState.Should().BeNull();
        ExternalDocumentMetadata.Empty.IsSaveStateVisible.Should().BeTrue("missing flag defaults to visible");
    }

    [Test]
    public void Serialize_EmitsCamelCaseKeys()
    {
        var json = new ExternalDocumentMetadata("https://x/y", "C3", true).Serialize();

        json.Should().Contain("\"returnUrl\":\"https://x/y\"");
        json.Should().Contain("\"classification\":\"C3\"");
        json.Should().Contain("\"userDownload\":true");
    }

    [Test]
    public void SerializeThenParse_RoundTripsAllValues()
    {
        var original = new ExternalDocumentMetadata("https://app/cb", "Secret", true);

        var roundTripped = ExternalDocumentMetadata.Parse(original.Serialize());

        roundTripped.Should().Be(original);
    }
}

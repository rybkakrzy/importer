using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using D2ViewerEditor.Application.Validators.Documents;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class IngestExternalDocumentCommandValidatorTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private IngestExternalDocumentCommandValidator _validator = null!;

    [SetUp]
    public void SetUp() => _validator = new IngestExternalDocumentCommandValidator();

    private static IngestExternalDocumentCommand Valid(string? metadata = null) =>
        new(new byte[] { 1, 2, 3 }, "ext.docx", DocxMime, "ExternalApp", metadata);

    [Test]
    public void Validate_ValidCommand_Passes()
    {
        _validator.Validate(Valid()).IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_EmptyFileName_Fails()
    {
        var result = _validator.Validate(Valid() with { FileName = "" });
        result.IsValid.Should().BeFalse();
    }

    [Test]
    public void Validate_UnsupportedMime_Fails()
    {
        var result = _validator.Validate(Valid() with { MimeType = "text/plain" });
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Wspierane są tylko pliki DOCX i PDF");
    }

    [Test]
    public void Validate_EmptyContent_Fails()
    {
        var result = _validator.Validate(Valid() with { Content = System.Array.Empty<byte>() });
        result.IsValid.Should().BeFalse();
    }

    [Test]
    public void Validate_EmptyCreatedBy_Fails()
    {
        var result = _validator.Validate(Valid() with { CreatedBy = "" });
        result.IsValid.Should().BeFalse();
    }

    [Test]
    public void Validate_InvalidJsonMetadata_Fails()
    {
        var result = _validator.Validate(Valid("{ bad json"));
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Metadata musi być poprawnym JSON-em");
    }

    [Test]
    public void Validate_ValidJsonMetadata_Passes()
    {
        _validator.Validate(Valid("{\"returnUrl\":\"https://x\"}")).IsValid.Should().BeTrue();
    }
}

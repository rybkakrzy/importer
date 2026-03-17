using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;
using D2ViewerEditor.Application.Validators;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class SaveDocumentValidatorTests
{
    private SaveDocumentValidator _validator;

    [SetUp]
    public void SetUp()
    {
        _validator = new SaveDocumentValidator();
    }

    [Test]
    public void Validate_WithValidHtml_ShouldPass()
    {
        // Arrange
        var command = new SaveDocumentCommand(
            Html: "<p>Treść dokumentu</p>",
            OriginalFileName: "test.docx",
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
        result.Errors.Should().BeEmpty();
    }

    [Test]
    public void Validate_WithEmptyHtml_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentCommand(
            Html: "",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Treść HTML dokumentu jest wymagana");
    }

    [Test]
    public void Validate_WithNullHtml_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentCommand(
            Html: null!,
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Treść HTML dokumentu jest wymagana");
    }

    [Test]
    public void Validate_WithWhitespaceHtml_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentCommand(
            Html: "   ",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
    }
}

using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Application.Validators.Documents;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class UploadDocumentCommandValidatorTests
{
    private UploadDocumentCommandValidator _validator;

    [SetUp]
    public void SetUp()
    {
        _validator = new UploadDocumentCommandValidator();
    }

    [Test]
    public void Validate_WithValidCommand_ShouldPass()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1, 2, 3 },
            FileName: "dokument.pdf",
            MimeType: "application/pdf",
            CreatedBy: "JanKowalski"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_WithEmptyFileName_ShouldFail()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: "",
            MimeType: "application/pdf",
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Nazwa dokumentu jest wymagana");
    }

    [Test]
    public void Validate_WithFileNameTooLong_ShouldFail()
    {
        // Arrange
        var longName = new string('a', 501);
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: longName,
            MimeType: "application/pdf",
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Nazwa dokumentu nie może przekraczać 500 znaków");
    }

    [Test]
    public void Validate_WithEmptyMimeType_ShouldFail()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: "doc.pdf",
            MimeType: "",
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Typ MIME jest wymagany");
    }

    [Test]
    public void Validate_WithInvalidMimeType_ShouldFail()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: "doc.pdf",
            MimeType: "notvalid",
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Nieprawidłowy format typu MIME");
    }

    [Test]
    public void Validate_WithEmptyContent_ShouldFail()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: Array.Empty<byte>(),
            FileName: "doc.pdf",
            MimeType: "application/pdf",
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Zawartość dokumentu nie może być pusta");
    }

    [Test]
    public void Validate_WithEmptyCreatedBy_ShouldFail()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: "doc.pdf",
            MimeType: "application/pdf",
            CreatedBy: ""
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Pole 'CreatedBy' jest wymagane");
    }

    [Test]
    public void Validate_WithValidMimeTypeWithSubtype_ShouldPass()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1, 2 },
            FileName: "doc.docx",
            MimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            CreatedBy: "Admin"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }
}

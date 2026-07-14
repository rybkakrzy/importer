using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Application.Validators.Documents;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class SaveDocumentVersionCommandValidatorTests
{
    private SaveDocumentVersionCommandValidator _validator;

    [SetUp]
    public void SetUp()
    {
        _validator = new SaveDocumentVersionCommandValidator();
    }

    [Test]
    public void Validate_WithValidCommand_ShouldPass()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            Content: new byte[] { 1, 2, 3 },
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_WithEmptyMasterId_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.Empty,
            Content: new byte[] { 1 },
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "GUID mastera jest wymagany");
    }

    [Test]
    public void Validate_WithNullContent_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            Content: null!,
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Zawartość dokumentu jest wymagana");
    }

    [Test]
    public void Validate_WithEmptyContent_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            Content: Array.Empty<byte>(),
            CreatedBy: "User"
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Zawartość dokumentu nie może być pusta"
            && e.ErrorCode == "DOCUMENT_CONTENT_EMPTY");
    }

    [Test]
    public void Validate_WithEmptyCreatedBy_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            Content: new byte[] { 1 },
            CreatedBy: ""
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Pole 'CreatedBy' jest wymagane");
    }

    [Test]
    public void Validate_WithCreatedByTooLong_ShouldFail()
    {
        // Arrange
        var command = new SaveDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            Content: new byte[] { 1 },
            CreatedBy: new string('a', 256)
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Pole 'CreatedBy' nie może przekraczać 255 znaków");
    }
}

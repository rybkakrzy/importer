using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Application.Validators.Documents;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class RestoreDocumentVersionCommandValidatorTests
{
    private RestoreDocumentVersionCommandValidator _validator;

    [SetUp]
    public void SetUp()
    {
        _validator = new RestoreDocumentVersionCommandValidator();
    }

    [Test]
    public void Validate_WithValidCommand_ShouldPass()
    {
        // Arrange
        var command = new RestoreDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            VersionId: Guid.NewGuid()
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
        var command = new RestoreDocumentVersionCommand(
            MasterId: Guid.Empty,
            VersionId: Guid.NewGuid()
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "GUID mastera jest wymagany");
    }

    [Test]
    public void Validate_WithEmptyVersionId_ShouldFail()
    {
        // Arrange
        var command = new RestoreDocumentVersionCommand(
            MasterId: Guid.NewGuid(),
            VersionId: Guid.Empty
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "GUID wersji jest wymagany");
    }

    [Test]
    public void Validate_WithBothEmptyGuids_ShouldReturnTwoErrors()
    {
        // Arrange
        var command = new RestoreDocumentVersionCommand(
            MasterId: Guid.Empty,
            VersionId: Guid.Empty
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().HaveCount(2);
        result.Errors.Should().Contain(e => e.ErrorMessage == "GUID mastera jest wymagany");
        result.Errors.Should().Contain(e => e.ErrorMessage == "GUID wersji jest wymagany");
    }
}

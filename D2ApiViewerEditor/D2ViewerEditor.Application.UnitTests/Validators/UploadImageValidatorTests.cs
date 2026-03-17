using D2ViewerEditor.Application.Features.Documents.Commands.UploadImage;
using D2ViewerEditor.Application.Validators;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class UploadImageValidatorTests
{
    private UploadImageValidator _validator;

    [SetUp]
    public void SetUp()
    {
        _validator = new UploadImageValidator();
    }

    [Test]
    public void Validate_WithValidCommand_ShouldPass()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "photo.jpg",
            ContentType: "image/jpeg",
            FileSize: 1024
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
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "",
            ContentType: "image/jpeg",
            FileSize: 1024
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Nazwa pliku jest wymagana");
    }

    [Test]
    public void Validate_WithEmptyContentType_ShouldFail()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "photo.jpg",
            ContentType: "",
            FileSize: 1024
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Typ pliku jest wymagany");
    }

    [Test]
    public void Validate_WithNonImageContentType_ShouldFail()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "file.pdf",
            ContentType: "application/pdf",
            FileSize: 1024
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Plik musi być obrazem (image/*)");
    }

    [Test]
    public void Validate_WithFileSizeExceedingLimit_ShouldFail()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "big.png",
            ContentType: "image/png",
            FileSize: 11 * 1024 * 1024 // 11 MB
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.ErrorMessage == "Rozmiar obrazu nie może przekraczać 10 MB");
    }

    [Test]
    public void Validate_WithFileSizeAtLimit_ShouldPass()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "big.png",
            ContentType: "image/png",
            FileSize: 10 * 1024 * 1024 // exactly 10 MB
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_WithImagePngContentType_ShouldPass()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "image.png",
            ContentType: "image/png",
            FileSize: 512
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_ContentTypeCheckIsCaseInsensitive()
    {
        // Arrange
        var command = new UploadImageCommand(
            FileStream: Stream.Null,
            FileName: "photo.jpg",
            ContentType: "Image/JPEG",
            FileSize: 512
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }
}

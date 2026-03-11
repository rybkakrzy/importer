using D2ViewerEditor.Application.Features.Barcode.Commands.GenerateBarcode;
using D2ViewerEditor.Application.Validators;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class GenerateBarcodeValidatorTests
{
    private GenerateBarcodeValidator _validator = null!;

    [SetUp]
    public void Setup()
    {
        _validator = new GenerateBarcodeValidator();
    }

    [Test]
    public void Validate_WithValidCommand_ShouldPass()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "Code128",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
        result.Errors.Should().BeEmpty();
    }

    [Test]
    public void Validate_WithEmptyContent_ShouldFail()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "",
            BarcodeType: "Code128",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Treść kodu kreskowego jest wymagana");
    }

    [Test]
    public void Validate_WithContentTooLong_ShouldFail()
    {
        // Arrange
        var longContent = new string('A', 501);
        var command = new GenerateBarcodeCommand(
            Content: longContent,
            BarcodeType: "Code128",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Treść nie może przekraczać 500 znaków");
    }

    [Test]
    public void Validate_WithEmptyBarcodeType_ShouldFail()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Typ kodu kreskowego jest wymagany");
    }

    [TestCase(49)]
    [TestCase(2001)]
    public void Validate_WithInvalidWidth_ShouldFail(int width)
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "Code128",
            Width: width,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Szerokość musi być między 50 a 2000 px");
    }

    [TestCase(49)]
    [TestCase(2001)]
    public void Validate_WithInvalidHeight_ShouldFail(int height)
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "Code128",
            Width: 300,
            Height: height,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().ContainSingle()
            .Which.ErrorMessage.Should().Be("Wysokość musi być między 50 a 2000 px");
    }

    [TestCase(50)]
    [TestCase(300)]
    [TestCase(2000)]
    public void Validate_WithValidWidthBoundaries_ShouldPass(int width)
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "Code128",
            Width: width,
            Height: 300,
            ShowText: false
        );

        // Act
        var result = _validator.Validate(command);

        // Assert
        result.IsValid.Should().BeTrue();
    }
}

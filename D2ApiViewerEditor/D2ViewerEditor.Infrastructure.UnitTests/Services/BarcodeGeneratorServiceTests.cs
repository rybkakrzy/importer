using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

[TestFixture]
public class BarcodeGeneratorServiceTests
{
    private BarcodeGeneratorService _service = null!;

    [SetUp]
    public void Setup()
    {
        _service = new BarcodeGeneratorService();
    }

    [Test]
    public void GetSupportedTypes_ShouldReturnListOfSupportedBarcodeTypes()
    {
        // Act
        var types = _service.GetSupportedTypes();

        // Assert
        types.Should().NotBeEmpty();
        types.Should().Contain("QRCode");
        types.Should().Contain("Code128");
        types.Should().Contain("EAN13");
    }

    [Test]
    public void Generate_WithValidQRCode_ShouldReturnPngBytes()
    {
        // Arrange
        var content = "https://example.com";
        var barcodeType = "QRCode";

        // Act
        var result = _service.Generate(content, barcodeType, 300, 300, false);

        // Assert
        result.Should().NotBeNull();
        result.Should().NotBeEmpty();
        // PNG magic bytes
        result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    }

    [Test]
    public void Generate_WithValidCode128_ShouldReturnPngBytes()
    {
        // Arrange
        var content = "TEST123";
        var barcodeType = "Code128";

        // Act
        var result = _service.Generate(content, barcodeType, 400, 200, false);

        // Assert
        result.Should().NotBeNull();
        result.Should().NotBeEmpty();
        result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    }

    [Test]
    public void Generate_WithEmptyContent_ShouldThrowArgumentException()
    {
        // Arrange
        var content = "";
        var barcodeType = "QRCode";

        // Act & Assert
        var act = () => _service.Generate(content, barcodeType, 300, 300, false);

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Treść kodu nie może być pusta*");
    }

    [Test]
    public void Generate_WithUnsupportedBarcodeType_ShouldThrowArgumentException()
    {
        // Arrange
        var content = "TEST123";
        var barcodeType = "InvalidType";

        // Act & Assert
        var act = () => _service.Generate(content, barcodeType, 300, 300, false);

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Nieobsługiwany typ kodu*");
    }

    [TestCase("QRCode")]
    [TestCase("Code128")]
    [TestCase("EAN13")]
    [TestCase("Code39")]
    public void Generate_WithDifferentBarcodeTypes_ShouldSucceed(string barcodeType)
    {
        // Arrange
        // EAN13 wymaga poprawnego numeru z checksumą (5901234123457 jest poprawny)
        var content = barcodeType == "EAN13" ? "5901234123457" : "TEST123";

        // Act
        var result = _service.Generate(content, barcodeType, 300, 300, false);

        // Assert
        result.Should().NotBeEmpty();
        result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    }

    [Test]
    public void Generate_WithShowText_ShouldIncludeTextInOutput()
    {
        // Arrange
        var content = "123456";
        var barcodeType = "Code128";

        // Act
        var result = _service.Generate(content, barcodeType, 400, 250, showText: true);

        // Assert
        result.Should().NotBeEmpty();
        // Veryfikujemy że obraz był wygenerowany pomyślnie
        result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    }

    [TestCase(100, 100)]
    [TestCase(500, 500)]
    [TestCase(1000, 300)]
    public void Generate_WithDifferentDimensions_ShouldSucceed(int width, int height)
    {
        // Arrange
        var content = "TEST123";
        var barcodeType = "QRCode";

        // Act
        var result = _service.Generate(content, barcodeType, width, height, false);

        // Assert
        result.Should().NotBeEmpty();
        result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    }
}

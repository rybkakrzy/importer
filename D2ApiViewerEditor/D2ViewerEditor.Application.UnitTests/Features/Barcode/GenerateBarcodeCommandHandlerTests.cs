using D2ViewerEditor.Application.Features.Barcode.Commands.GenerateBarcode;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NSubstitute;
using NSubstitute.ExceptionExtensions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Barcode;

[TestFixture]
public class GenerateBarcodeCommandHandlerTests
{
    private IBarcodeGenerator _barcodeGenerator = null!;
    private GenerateBarcodeCommandHandler _handler = null!;

    [SetUp]
    public void Setup()
    {
        _barcodeGenerator = Substitute.For<IBarcodeGenerator>();
        _handler = new GenerateBarcodeCommandHandler(_barcodeGenerator);
    }

    [Test]
    public async Task Handle_WithValidCommand_ShouldReturnSuccessResult()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "Code128",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        var fakePngBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47 }; // PNG header
        _barcodeGenerator
            .Generate(command.Content, command.BarcodeType, 300, 300, command.ShowText)
            .Returns(fakePngBytes);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().NotBeNull();
        result.Value!.BarcodeType.Should().Be("Code128");
        result.Value.ContentType.Should().Be("image/png");
        result.Value.Base64Image.Should().StartWith("data:image/png;base64,");
    }

    [Test]
    public async Task Handle_ShouldClampWidthAndHeightTo1000()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "QRCode",
            Width: 5000, // Powyżej limitu
            Height: 5000, // Powyżej limitu
            ShowText: false
        );

        var fakePngBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        _barcodeGenerator
            .Generate(Arg.Any<string>(), Arg.Any<string>(), Arg.Any<int>(), Arg.Any<int>(), Arg.Any<bool>())
            .Returns(fakePngBytes);

        // Act
        await _handler.Handle(command, CancellationToken.None);

        // Assert
        _barcodeGenerator.Received(1).Generate(
            "TEST123",
            "QRCode",
            1000, // Zclampowane
            1000, // Zclampowane
            false
        );
    }

    [Test]
    public async Task Handle_ShouldClampWidthAndHeightToMinimum50()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "TEST123",
            BarcodeType: "QRCode",
            Width: 10, // Poniżej limitu
            Height: 10, // Poniżej limitu
            ShowText: false
        );

        var fakePngBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        _barcodeGenerator
            .Generate(Arg.Any<string>(), Arg.Any<string>(), Arg.Any<int>(), Arg.Any<int>(), Arg.Any<bool>())
            .Returns(fakePngBytes);

        // Act
        await _handler.Handle(command, CancellationToken.None);

        // Assert
        _barcodeGenerator.Received(1).Generate(
            "TEST123",
            "QRCode",
            50, // Zclampowane
            50, // Zclampowane
            false
        );
    }

    [Test]
    public async Task Handle_WhenGeneratorThrowsArgumentException_ShouldReturnFailureResult()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "INVALID",
            BarcodeType: "InvalidType",
            Width: 300,
            Height: 300,
            ShowText: false
        );

        var errorMessage = "Nieobsługiwany typ kodu kreskowego";
        _barcodeGenerator
            .Generate(Arg.Any<string>(), Arg.Any<string>(), Arg.Any<int>(), Arg.Any<int>(), Arg.Any<bool>())
            .Throws(new ArgumentException(errorMessage));

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Be(errorMessage);
    }

    [Test]
    public async Task Handle_WithShowTextTrue_ShouldPassShowTextToGenerator()
    {
        // Arrange
        var command = new GenerateBarcodeCommand(
            Content: "123456789",
            BarcodeType: "EAN13",
            Width: 400,
            Height: 200,
            ShowText: true
        );

        var fakePngBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        _barcodeGenerator
            .Generate(Arg.Any<string>(), Arg.Any<string>(), Arg.Any<int>(), Arg.Any<int>(), Arg.Any<bool>())
            .Returns(fakePngBytes);

        // Act
        await _handler.Handle(command, CancellationToken.None);

        // Assert
        _barcodeGenerator.Received(1).Generate(
            command.Content,
            command.BarcodeType,
            Arg.Any<int>(),
            Arg.Any<int>(),
            true
        );
    }
}

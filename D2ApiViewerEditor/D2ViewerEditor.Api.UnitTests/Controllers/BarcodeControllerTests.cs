using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Barcode.Commands.GenerateBarcode;
using D2ViewerEditor.Application.Features.Barcode.Queries.GetSupportedTypes;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using MediatR;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Controllers;

[TestFixture]
public class BarcodeControllerTests
{
    private IMediator _mediator = null!;
    private BarcodeController _controller = null!;
    private ServiceCollection _services = null!;
    private ServiceProvider _serviceProvider = null!;

    [SetUp]
    public void Setup()
    {
        _mediator = Substitute.For<IMediator>();
        
        // Setupujemy DI container
        _services = new ServiceCollection();
        _services.AddSingleton(_mediator);
        _serviceProvider = _services.BuildServiceProvider();
        
        // Tworzymy kontroler
        _controller = new BarcodeController();
        
        // Mock HttpContext z DI
        var httpContext = new DefaultHttpContext
        {
            RequestServices = _serviceProvider
        };
        
        _controller.ControllerContext = new ControllerContext
        {
            HttpContext = httpContext
        };
    }

    [TearDown]
    public void TearDown()
    {
        _serviceProvider?.Dispose();
    }

    [Test]
    public async Task GetSupportedTypes_ShouldReturnOkWithListOfTypes()
    {
        // Arrange
        var supportedTypes = new List<string> { "QRCode", "Code128", "EAN13" }.AsReadOnly();
        _mediator.Send(Arg.Any<GetSupportedTypesQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<IReadOnlyList<string>>.Success(supportedTypes));

        // Act
        var result = await _controller.GetSupportedTypes();

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var okResult = result as OkObjectResult;
        var response = okResult!.Value as IReadOnlyList<string>;
        response.Should().Contain("QRCode");
        response.Should().Contain("Code128");
    }

    [Test]
    public async Task GenerateBarcode_WithValidRequest_ShouldReturnOkWithBarcodeResponse()
    {
        // Arrange
        var request = new BarcodeRequest
        {
            Content = "TEST123",
            BarcodeType = "Code128",
            Width = 300,
            Height = 300,
            ShowText = false
        };

        var barcodeResponse = new BarcodeResponse
        {
            Base64Image = "data:image/png;base64,iVBORw0KGgo...",
            ContentType = "image/png",
            BarcodeType = "Code128"
        };

        _mediator.Send(Arg.Any<GenerateBarcodeCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<BarcodeResponse>.Success(barcodeResponse));

        // Act
        var result = await _controller.GenerateBarcode(request);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var okResult = result as OkObjectResult;
        var response = okResult!.Value as BarcodeResponse;
        response.Should().NotBeNull();
        response!.BarcodeType.Should().Be("Code128");
        response.Base64Image.Should().StartWith("data:image/png;base64,");
    }

    [Test]
    public async Task GenerateBarcode_WhenCommandFails_ShouldReturnBadRequest()
    {
        // Arrange
        var request = new BarcodeRequest
        {
            Content = "INVALID",
            BarcodeType = "InvalidType",
            Width = 300,
            Height = 300
        };

        _mediator.Send(Arg.Any<GenerateBarcodeCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<BarcodeResponse>.Failure("Nieobsługiwany typ kodu"));

        // Act
        var result = await _controller.GenerateBarcode(request);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task GenerateBarcodeImage_WithValidRequest_ShouldReturnFileResult()
    {
        // Arrange
        var request = new BarcodeRequest
        {
            Content = "123456",
            BarcodeType = "QRCode",
            Width = 300,
            Height = 300
        };

        var base64Image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        var barcodeResponse = new BarcodeResponse
        {
            Base64Image = base64Image,
            ContentType = "image/png",
            BarcodeType = "QRCode"
        };

        _mediator.Send(Arg.Any<GenerateBarcodeCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<BarcodeResponse>.Success(barcodeResponse));

        // Act
        var result = await _controller.GenerateBarcodeImage(request);

        // Assert
        result.Should().BeOfType<FileContentResult>();
        var fileResult = result as FileContentResult;
        fileResult!.ContentType.Should().Be("image/png");
        fileResult.FileDownloadName.Should().Be("QRCode.png");
    }

    [Test]
    public async Task GenerateBarcodeImage_WhenCommandFails_ShouldReturnBadRequest()
    {
        // Arrange
        var request = new BarcodeRequest
        {
            Content = "",
            BarcodeType = "QRCode",
            Width = 300,
            Height = 300
        };

        _mediator.Send(Arg.Any<GenerateBarcodeCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<BarcodeResponse>.Failure("Treść nie może być pusta"));

        // Act
        var result = await _controller.GenerateBarcodeImage(request);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }
}

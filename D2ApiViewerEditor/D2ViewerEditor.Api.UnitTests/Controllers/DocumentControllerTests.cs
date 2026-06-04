using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UploadImage;
using D2ViewerEditor.Application.Features.Documents.Queries.GetNewDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetTemplate;
using D2ViewerEditor.Application.Features.Documents.Queries.GetTemplates;
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
public class DocumentControllerTests
{
    private IMediator _mediator;
    private DocumentController _controller;
    private ServiceProvider _serviceProvider;

    [SetUp]
    public void Setup()
    {
        _mediator = Substitute.For<IMediator>();

        var services = new ServiceCollection();
        services.AddSingleton(_mediator);
        _serviceProvider = services.BuildServiceProvider();

        _controller = new DocumentController();
        _controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext { RequestServices = _serviceProvider }
        };
    }

    [TearDown]
    public void TearDown()
    {
        _serviceProvider?.Dispose();
    }

    [Test]
    public async Task NewDocument_ShouldReturnOkWithDocumentContent()
    {
        // Arrange
        var content = new DocumentContent
        {
            Html = "<p>&nbsp;</p>",
            Metadata = new DocumentMetadata { Title = "Nowy dokument" }
        };
        _mediator.Send(Arg.Any<GetNewDocumentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentContent>.Success(content));

        // Act
        var result = await _controller.NewDocument();

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var ok = (OkObjectResult)result;
        ok.Value.Should().Be(content);
    }

    [Test]
    public async Task GetTemplates_ShouldReturnOkWithList()
    {
        // Arrange
        var templates = new List<TemplateDto>
        {
            new("blank", "Pusty", "Opis"),
            new("letter", "List", "Opis")
        };
        _mediator.Send(Arg.Any<GetTemplatesQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<List<TemplateDto>>.Success(templates));

        // Act
        var result = await _controller.GetTemplates();

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var ok = (OkObjectResult)result;
        (ok.Value as List<TemplateDto>)!.Should().HaveCount(2);
    }

    [Test]
    public async Task GetTemplate_WithValidId_ShouldReturnOkWithContent()
    {
        // Arrange
        var content = new DocumentContent { Html = "<p>List</p>", Metadata = new DocumentMetadata() };
        _mediator.Send(Arg.Any<GetTemplateQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentContent>.Success(content));

        // Act
        var result = await _controller.GetTemplate("letter");

        // Assert
        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task GetTemplate_WhenFails_ShouldReturnBadRequest()
    {
        // Arrange
        _mediator.Send(Arg.Any<GetTemplateQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentContent>.Failure("Nieznany szablon"));

        // Act
        var result = await _controller.GetTemplate("nonexistent");

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task SaveDocument_WithValidRequest_ShouldReturnFileResult()
    {
        // Arrange
        var docxBytes = new byte[] { 1, 2, 3 };
        _mediator.Send(Arg.Any<SaveDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<SaveDocumentResult>.Success(new SaveDocumentResult(docxBytes, "test.docx")));

        var request = new SaveDocumentRequest
        {
            Html = "<p>Test</p>",
            OriginalFileName = "test.docx"
        };

        // Act
        var result = await _controller.SaveDocument(request);

        // Assert
        result.Should().BeOfType<FileContentResult>();
        var fileResult = (FileContentResult)result;
        fileResult.ContentType.Should().Be("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        fileResult.FileDownloadName.Should().Be("test.docx");
    }

    [Test]
    public async Task SaveDocument_WhenFails_ShouldReturnBadRequest()
    {
        // Arrange
        _mediator.Send(Arg.Any<SaveDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<SaveDocumentResult>.Failure("Błąd konwersji"));

        var request = new SaveDocumentRequest { Html = "" };

        // Act
        var result = await _controller.SaveDocument(request);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task OpenDocument_WithNullFile_ShouldReturnBadRequest()
    {
        // Act
        var result = await _controller.OpenDocument(null!);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task OpenDocument_WithNonDocxFile_ShouldReturnBadRequest()
    {
        // Arrange
        var fileMock = Substitute.For<IFormFile>();
        fileMock.FileName.Returns("file.pdf");
        fileMock.Length.Returns(100);

        // Act
        var result = await _controller.OpenDocument(fileMock);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task OpenDocument_WithLegacyDocFile_ShouldReturnBadRequest_WithConversionGuidance()
    {
        // .doc is the legacy binary Word format — must be rejected with an actionable message,
        // not handed to the DOCX parser (which would fail deep with an opaque error).
        var fileMock = Substitute.For<IFormFile>();
        fileMock.FileName.Returns("stary.doc");
        fileMock.Length.Returns(100);

        var result = await _controller.OpenDocument(fileMock);

        var bad = result.Should().BeOfType<BadRequestObjectResult>().Subject;
        bad.Value!.ToString().Should().Contain(".docx");
    }

    [Test]
    public async Task UploadImage_WithNullFile_ShouldReturnBadRequest()
    {
        // Act
        var result = await _controller.UploadImage(null!);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task UploadImage_WithValidFile_ShouldReturnOk()
    {
        // Arrange
        var fileContent = new byte[] { 1, 2, 3 };
        var stream = new MemoryStream(fileContent);
        var fileMock = Substitute.For<IFormFile>();
        fileMock.FileName.Returns("photo.jpg");
        fileMock.ContentType.Returns("image/jpeg");
        fileMock.Length.Returns(fileContent.Length);
        fileMock.OpenReadStream().Returns(stream);

        var response = new ImageUploadResponse("data:image/jpeg;base64,AAAA", "photo.jpg", fileContent.Length);
        _mediator.Send(Arg.Any<UploadImageCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<ImageUploadResponse>.Success(response));

        // Act
        var result = await _controller.UploadImage(fileMock);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public void ExportToPdf_ShouldReturn501()
    {
        // Arrange
        var request = new SaveDocumentRequest { Html = "<p>Test</p>" };

        // Act
        var result = _controller.ExportToPdf(request);

        // Assert
        result.Should().BeOfType<ObjectResult>()
            .Which.StatusCode.Should().Be(501);
    }
}

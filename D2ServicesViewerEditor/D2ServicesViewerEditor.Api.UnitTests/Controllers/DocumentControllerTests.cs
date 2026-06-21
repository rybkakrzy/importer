using System.Text;
using D2ServicesViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using FluentAssertions;
using MediatR;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using NSubstitute;
using NUnit.Framework;

namespace D2ServicesViewerEditor.Api.UnitTests.Controllers;

[TestFixture]
public class DocumentControllerTests
{
    private IMediator _mediator = null!;
    private DocumentController _controller = null!;

    [SetUp]
    public void SetUp()
    {
        _mediator = Substitute.For<IMediator>();
        _controller = new DocumentController(NullLogger<DocumentController>.Instance, _mediator);
        _controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
    }

    [Test]
    public async Task CreateDocument_WhenFileMissing_ReturnsBadRequest()
    {
        var result = await _controller.CreateDocument(new CreateDocumentRequest(), CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
        await _mediator.DidNotReceiveWithAnyArgs().Send(default(IRequest<Result<IngestExternalDocumentResult>>)!, default);
    }

    [Test]
    public async Task CreateDocument_WhenClassificationInvalid_ReturnsBadRequest()
    {
        var request = new CreateDocumentRequest
        {
            File = BuildFormFile("doc.pdf", "application/pdf", "x"),
            Classification = "BAD"
        };

        var result = await _controller.CreateDocument(request, CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
        await _mediator.DidNotReceiveWithAnyArgs().Send(default(IRequest<Result<IngestExternalDocumentResult>>)!, default);
    }

    [Test]
    public async Task CreateDocument_WhenDocxWithoutReturnUrl_ReturnsBadRequest()
    {
        var request = new CreateDocumentRequest
        {
            File = BuildFormFile("doc.docx", IngestExternalDocumentCommandHandler.DocxMimeType, "abc")
        };

        var result = await _controller.CreateDocument(request, CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task CreateDocument_WhenPdfAndNoContentType_ResolvesFromExtension_AndReturnsCreated()
    {
        var masterId = Guid.NewGuid();
        var request = new CreateDocumentRequest
        {
            File = BuildFormFile("doc.pdf", string.Empty, "pdf-content"),
            Classification = "C2"
        };

        _mediator.Send(Arg.Any<IngestExternalDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<IngestExternalDocumentResult>.Success(new IngestExternalDocumentResult(masterId, null, "doc.pdf", DateTime.UtcNow)));

        var result = await _controller.CreateDocument(request, CancellationToken.None);

        var objectResult = result.Should().BeOfType<ObjectResult>().Subject;
        objectResult.StatusCode.Should().Be(StatusCodes.Status201Created);
        objectResult.Value.Should().BeOfType<CreateDocumentResponse>();

        await _mediator.Received(1).Send(
            Arg.Is<IngestExternalDocumentCommand>(c => c.MimeType == IngestExternalDocumentCommandHandler.PdfMimeType && c.FileName == "doc.pdf"),
            Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task GetDocumentStatus_WhenNotFound_ReturnsNotFound()
    {
        _mediator.Send(Arg.Any<GetDocumentStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentStatusDto>.NotFound());

        var result = await _controller.GetDocumentStatus(Guid.NewGuid(), CancellationToken.None);

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task GetDocumentStatus_WhenSuccess_ReturnsEnumStatus()
    {
        var masterId = Guid.NewGuid();
        _mediator.Send(Arg.Any<GetDocumentStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentStatusDto>.Success(
                new DocumentStatusDto(masterId, DocumentStatus.Saved.ToString(), false, false, false, null, null, null, null)));

        var result = await _controller.GetDocumentStatus(masterId, CancellationToken.None);

        var ok = result.Should().BeOfType<OkObjectResult>().Subject;
        var response = ok.Value.Should().BeOfType<DocumentStatusResponse>().Subject;
        response.MasterId.Should().Be(masterId);
        response.Status.Should().Be(DocumentStatus.Saved);
    }

    private static IFormFile BuildFormFile(string fileName, string contentType, string content)
    {
        var bytes = Encoding.UTF8.GetBytes(content);
        var stream = new MemoryStream(bytes);
        return new FormFile(stream, 0, bytes.Length, "file", fileName)
        {
            Headers = new HeaderDictionary(),
            ContentType = contentType
        };
    }
}

using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.AbortSend;
using D2ViewerEditor.Application.Features.Documents.Commands.CancelDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.ContinueDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.DownloadEditedDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.RequeueDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveriesByStatus;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveryStatus;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentBaseContent;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocuments;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersionContent;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;
using D2ViewerEditor.Domain.Common;
using FluentAssertions;
using MediatR;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Controllers;

[TestFixture]
public class DocumentStorageControllerTests
{
    private IMediator _mediator;
    private DocumentStorageController _controller;
    private ServiceProvider _serviceProvider;

    [SetUp]
    public void Setup()
    {
        _mediator = Substitute.For<IMediator>();

        var services = new ServiceCollection();
        services.AddSingleton(_mediator);
        _serviceProvider = services.BuildServiceProvider();

        _controller = new DocumentStorageController();
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
    public async Task UploadDocument_WithValidRequest_ShouldReturnOk()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versionId = Guid.NewGuid();
        var uploadResult = new UploadDocumentResult(masterId, versionId, "test.pdf", DateTime.UtcNow);

        _mediator.Send(Arg.Any<UploadDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UploadDocumentResult>.Success(uploadResult));

        var request = new UploadDocumentRequest("test.pdf", "application/pdf", new byte[] { 1, 2, 3 }, "User");

        // Act
        var result = await _controller.UploadDocument(request);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var ok = (OkObjectResult)result;
        ok.Value.Should().Be(uploadResult);
    }

    [Test]
    public async Task UploadDocument_WhenCommandFails_ShouldReturnBadRequest()
    {
        // Arrange
        _mediator.Send(Arg.Any<UploadDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UploadDocumentResult>.Failure("Błąd uploadu"));

        var request = new UploadDocumentRequest("test.pdf", "application/pdf", new byte[] { 1 }, "User");

        // Act
        var result = await _controller.UploadDocument(request);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task GetDocument_WithExistingId_ShouldReturnOk()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var docDto = new DocumentDto(masterId, "test.pdf", "application/pdf",
            DateTime.UtcNow, Guid.NewGuid(), new byte[] { 1 }, 1);

        _mediator.Send(Arg.Any<GetDocumentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentDto>.Success(docDto));

        // Act
        var result = await _controller.GetDocument(masterId);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var ok = (OkObjectResult)result;
        ok.Value.Should().Be(docDto);
    }

    [Test]
    public async Task GetDocument_WithNonExistingId_ShouldReturnNotFound()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _mediator.Send(Arg.Any<GetDocumentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentDto>.Failure("Dokument nie znaleziony"));

        // Act
        var result = await _controller.GetDocument(masterId);

        // Assert
        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task GetDocumentVersions_WithExistingId_ShouldReturnOk()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versions = new List<DocumentVersionDto>
        {
            new(masterId, Guid.NewGuid(), 1, DateTime.UtcNow, "User", true, 1024),
            new(masterId, Guid.NewGuid(), 2, DateTime.UtcNow, "User", false, 2048)
        };

        _mediator.Send(Arg.Any<GetDocumentVersionsQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<List<DocumentVersionDto>>.Success(versions));

        // Act
        var result = await _controller.GetDocumentVersions(masterId);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var ok = (OkObjectResult)result;
        (ok.Value as List<DocumentVersionDto>)!.Should().HaveCount(2);
    }

    [Test]
    public async Task GetDocumentVersions_WithNonExistingId_ShouldReturnNotFound()
    {
        // Arrange
        _mediator.Send(Arg.Any<GetDocumentVersionsQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<List<DocumentVersionDto>>.Failure("Dokument nie istnieje"));

        // Act
        var result = await _controller.GetDocumentVersions(Guid.NewGuid());

        // Assert
        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task SaveDocumentVersion_WithValidCommand_ShouldReturnOk()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versionResult = new SaveDocumentVersionResult(Guid.NewGuid(), 2, DateTime.UtcNow);

        _mediator.Send(Arg.Any<SaveDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<SaveDocumentVersionResult>.Success(versionResult));

        var request = new SaveDocumentVersionRequest(new byte[] { 1, 2, 3 }, "User");

        // Act
        var result = await _controller.SaveDocumentVersion(masterId, request);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task SaveDocumentVersion_WhenCommandFails_ShouldReturnBadRequest()
    {
        // Arrange
        _mediator.Send(Arg.Any<SaveDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<SaveDocumentVersionResult>.Failure("Dokument nie istnieje"));

        var request = new SaveDocumentVersionRequest(new byte[] { 1 }, "User");

        // Act
        var result = await _controller.SaveDocumentVersion(Guid.NewGuid(), request);

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task RestoreDocumentVersion_WithValidIds_ShouldReturnOk()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versionId = Guid.NewGuid();
        var restoreResult = new RestoreDocumentVersionResult(versionId, 1, "Przywrócono");

        _mediator.Send(Arg.Any<RestoreDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<RestoreDocumentVersionResult>.Success(restoreResult));

        // Act
        var result = await _controller.RestoreDocumentVersion(masterId, versionId);

        // Assert
        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task RestoreDocumentVersion_WhenCommandFails_ShouldReturnBadRequest()
    {
        // Arrange
        _mediator.Send(Arg.Any<RestoreDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<RestoreDocumentVersionResult>.Failure("Wersja nie istnieje"));

        // Act
        var result = await _controller.RestoreDocumentVersion(Guid.NewGuid(), Guid.NewGuid());

        // Assert
        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task DownloadBaseDocument_WithExistingId_ShouldReturnFileWithCorrectMimeType()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var content = new byte[] { 37, 80, 68, 70 };
        var dto = new DocumentBaseContentDto("contract_v1.pdf", "application/pdf", content);

        _mediator.Send(Arg.Any<GetDocumentBaseContentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentBaseContentDto>.Success(dto));

        // Act
        var result = await _controller.DownloadBaseDocument(masterId);

        // Assert
        result.Should().BeOfType<FileContentResult>();
        var file = (FileContentResult)result;
        file.FileContents.Should().BeEquivalentTo(content);
        file.ContentType.Should().Be("application/pdf");
        file.FileDownloadName.Should().Be("contract_v1.pdf");
    }

    [Test]
    public async Task DownloadBaseDocument_WithNonExistingId_ShouldReturnNotFound()
    {
        // Arrange
        _mediator.Send(Arg.Any<GetDocumentBaseContentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentBaseContentDto>.NotFound());

        // Act
        var result = await _controller.DownloadBaseDocument(Guid.NewGuid());

        // Assert
        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task DownloadDocumentVersion_WithValidIds_ShouldReturnFileWithCorrectMimeType()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versionId = Guid.NewGuid();
        var content = new byte[] { 1, 2, 3, 4 };
        var dto = new DocumentVersionContentDto("report_v2.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", content);

        _mediator.Send(Arg.Any<GetDocumentVersionContentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentVersionContentDto>.Success(dto));

        // Act
        var result = await _controller.DownloadDocumentVersion(masterId, versionId);

        // Assert
        result.Should().BeOfType<FileContentResult>();
        var file = (FileContentResult)result;
        file.FileContents.Should().BeEquivalentTo(content);
        file.ContentType.Should().Be("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        file.FileDownloadName.Should().Be("report_v2.docx");
    }

    [Test]
    public async Task DownloadDocumentVersion_WithNonExistingVersion_ShouldReturnNotFound()
    {
        // Arrange
        _mediator.Send(Arg.Any<GetDocumentVersionContentQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentVersionContentDto>.NotFound());

        // Act
        var result = await _controller.DownloadDocumentVersion(Guid.NewGuid(), Guid.NewGuid());

        // Assert
        result.Should().BeOfType<NotFoundObjectResult>();
    }

    // ── GetDocuments ─────────────────────────────────────────────────────────────

    [Test]
    public async Task GetDocuments_WhenSuccess_ShouldReturnOk()
    {
        var docs = new List<DocumentListItemDto>
        {
            new(Guid.NewGuid(), "a.docx", "mime", DateTime.UtcNow, Guid.NewGuid(), 1, "Saved", "ACME-42")
        };
        _mediator.Send(Arg.Any<GetDocumentsQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<List<DocumentListItemDto>>.Success(docs));

        var result = await _controller.GetDocuments();

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task GetDocuments_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<GetDocumentsQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<List<DocumentListItemDto>>.Failure("boom"));

        var result = await _controller.GetDocuments();

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── UpdateDocumentVersion (auto-save) ────────────────────────────────────────

    [Test]
    public async Task UpdateDocumentVersion_WhenSuccess_ShouldReturnOk()
    {
        var versionId = Guid.NewGuid();
        _mediator.Send(Arg.Any<UpdateDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDocumentVersionResult>.Success(
                new UpdateDocumentVersionResult(versionId, 2, 123, DateTime.UtcNow)));

        var result = await _controller.UpdateDocumentVersion(
            Guid.NewGuid(), versionId, new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task UpdateDocumentVersion_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<UpdateDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDocumentVersionResult>.NotFound());

        var result = await _controller.UpdateDocumentVersion(
            Guid.NewGuid(), Guid.NewGuid(), new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task UpdateDocumentVersion_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<UpdateDocumentVersionCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDocumentVersionResult>.Failure("v1 immutable"));

        var result = await _controller.UpdateDocumentVersion(
            Guid.NewGuid(), Guid.NewGuid(), new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── GetDocumentMetadata ──────────────────────────────────────────────────────

    [Test]
    public async Task GetDocumentMetadata_WhenSuccess_ShouldReturnOk()
    {
        var dto = new DocumentMetadataDto(Guid.NewGuid(), "mime", "https://cb", "C2", true, true);
        _mediator.Send(Arg.Any<GetDocumentMetadataQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentMetadataDto>.Success(dto));

        var result = await _controller.GetDocumentMetadata(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task GetDocumentMetadata_WhenForbidden_ShouldReturn403()
    {
        _mediator.Send(Arg.Any<GetDocumentMetadataQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentMetadataDto>.Forbidden("brak uprawnień"));

        var result = await _controller.GetDocumentMetadata(Guid.NewGuid());

        result.Should().BeOfType<ObjectResult>()
            .Which.StatusCode.Should().Be(StatusCodes.Status403Forbidden);
    }

    [Test]
    public async Task GetDocumentMetadata_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<GetDocumentMetadataQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DocumentMetadataDto>.NotFound());

        var result = await _controller.GetDocumentMetadata(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    // ── DownloadEditedDocument (user-download) ────────────────────────────────────

    [Test]
    public async Task DownloadEditedDocument_WhenSuccess_ShouldReturnFile()
    {
        _mediator.Send(Arg.Any<DownloadEditedDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<DownloadEditedDocumentResult>.Success(
                new DownloadEditedDocumentResult(new byte[] { 1, 2 }, "edited.docx")));

        var result = await _controller.DownloadEditedDocument(
            Guid.NewGuid(), new Domain.Models.SaveDocumentRequest { Html = "<p></p>" }, CancellationToken.None);

        result.Should().BeOfType<FileContentResult>()
            .Which.FileDownloadName.Should().Be("edited.docx");
    }

    [Test]
    public async Task DownloadEditedDocument_WhenForbiddenPrefix_ShouldReturn403()
    {
        _mediator.Send(Arg.Any<DownloadEditedDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<DownloadEditedDocumentResult>.Failure(
                DownloadEditedDocumentCommandHandler.ForbiddenErrorPrefix + " nie wolno"));

        var result = await _controller.DownloadEditedDocument(
            Guid.NewGuid(), new Domain.Models.SaveDocumentRequest { Html = "<p></p>" }, CancellationToken.None);

        result.Should().BeOfType<ObjectResult>()
            .Which.StatusCode.Should().Be(StatusCodes.Status403Forbidden);
    }

    [Test]
    public async Task DownloadEditedDocument_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<DownloadEditedDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<DownloadEditedDocumentResult>.NotFound());

        var result = await _controller.DownloadEditedDocument(
            Guid.NewGuid(), new Domain.Models.SaveDocumentRequest { Html = "<p></p>" }, CancellationToken.None);

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task DownloadEditedDocument_WhenOtherFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<DownloadEditedDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<DownloadEditedDocumentResult>.Failure("konwersja padła"));

        var result = await _controller.DownloadEditedDocument(
            Guid.NewGuid(), new Domain.Models.SaveDocumentRequest { Html = "<p></p>" }, CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── FinishAndSend ────────────────────────────────────────────────────────────

    [Test]
    public async Task FinishAndSend_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<FinishAndSendDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<FinishAndSendResult>.Success(
                new FinishAndSendResult(Guid.NewGuid(), "Sent", "Sent", true)));

        var result = await _controller.FinishAndSend(
            Guid.NewGuid(), Guid.NewGuid(), new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task FinishAndSend_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<FinishAndSendDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<FinishAndSendResult>.NotFound());

        var result = await _controller.FinishAndSend(
            Guid.NewGuid(), Guid.NewGuid(), new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task FinishAndSend_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<FinishAndSendDocumentCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<FinishAndSendResult>.Failure("brak returnUrl"));

        var result = await _controller.FinishAndSend(
            Guid.NewGuid(), Guid.NewGuid(), new SaveDocumentVersionRequest(new byte[] { 1 }, "User"));

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── AbortSend ────────────────────────────────────────────────────────────────

    [Test]
    public async Task AbortSend_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<AbortSendCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<AbortSendResult>.Success(
                new AbortSendResult(Guid.NewGuid(), "SendAborted", "Cancelled")));

        var result = await _controller.AbortSend(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task AbortSend_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<AbortSendCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<AbortSendResult>.NotFound());

        var result = await _controller.AbortSend(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task AbortSend_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<AbortSendCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<AbortSendResult>.Failure("w toku"));

        var result = await _controller.AbortSend(Guid.NewGuid());

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── ContinueDelivery ─────────────────────────────────────────────────────────

    [Test]
    public async Task ContinueDelivery_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<ContinueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<ContinueDeliveryResult>.Success(
                new ContinueDeliveryResult(Guid.NewGuid(), "Queued", Guid.NewGuid(), "Pending")));

        var result = await _controller.ContinueDelivery(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task ContinueDelivery_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<ContinueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<ContinueDeliveryResult>.NotFound());

        var result = await _controller.ContinueDelivery(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task ContinueDelivery_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<ContinueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<ContinueDeliveryResult>.Failure("brak zadania"));

        var result = await _controller.ContinueDelivery(Guid.NewGuid());

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    // ── Delivery status / list / retry / cancel / recipient-url ───────────────────

    [Test]
    public async Task GetDeliveryStatus_WhenSuccess_ShouldReturnOk()
    {
        var dto = new DeliveryStatusDto(Guid.NewGuid(), Guid.NewGuid(), "Sent", 1,
            DateTime.UtcNow, null, null, DateTime.UtcNow);
        _mediator.Send(Arg.Any<GetDeliveryStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DeliveryStatusDto>.Success(dto));

        var result = await _controller.GetDeliveryStatus(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task GetDeliveryStatus_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<GetDeliveryStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<DeliveryStatusDto>.NotFound());

        var result = await _controller.GetDeliveryStatus(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task GetDeliveries_WhenSuccess_ShouldReturnOk()
    {
        IReadOnlyList<DeliveryListItemDto> list = new List<DeliveryListItemDto>();
        _mediator.Send(Arg.Any<GetDeliveriesByStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<IReadOnlyList<DeliveryListItemDto>>.Success(list));

        var result = await _controller.GetDeliveries();

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task GetDeliveries_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<GetDeliveriesByStatusQuery>(), Arg.Any<CancellationToken>())
            .Returns(Result<IReadOnlyList<DeliveryListItemDto>>.Failure("zły status"));

        var result = await _controller.GetDeliveries("nope");

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task RetryDelivery_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<RequeueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<RequeueDeliveryResult>.Success(new RequeueDeliveryResult(Guid.NewGuid(), "Pending")));

        var result = await _controller.RetryDelivery(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task RetryDelivery_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<RequeueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<RequeueDeliveryResult>.NotFound());

        var result = await _controller.RetryDelivery(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task RetryDelivery_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<RequeueDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<RequeueDeliveryResult>.Failure("nie można"));

        var result = await _controller.RetryDelivery(Guid.NewGuid());

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task CancelDelivery_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<CancelDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<CancelDeliveryResult>.Success(new CancelDeliveryResult(Guid.NewGuid(), "Cancelled")));

        var result = await _controller.CancelDelivery(Guid.NewGuid());

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task CancelDelivery_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<CancelDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<CancelDeliveryResult>.NotFound());

        var result = await _controller.CancelDelivery(Guid.NewGuid());

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task CancelDelivery_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<CancelDeliveryCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<CancelDeliveryResult>.Failure("nie można"));

        var result = await _controller.CancelDelivery(Guid.NewGuid());

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task UpdateDeliveryRecipientUrl_WhenSuccess_ShouldReturnOk()
    {
        _mediator.Send(Arg.Any<UpdateDeliveryRecipientUrlCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDeliveryRecipientUrlResult>.Success(
                new UpdateDeliveryRecipientUrlResult(Guid.NewGuid(), "https://new", "RetryScheduled")));

        var result = await _controller.UpdateDeliveryRecipientUrl(
            Guid.NewGuid(), new UpdateDeliveryRecipientUrlRequest("https://new"));

        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public async Task UpdateDeliveryRecipientUrl_WhenNotFound_ShouldReturnNotFound()
    {
        _mediator.Send(Arg.Any<UpdateDeliveryRecipientUrlCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDeliveryRecipientUrlResult>.NotFound());

        var result = await _controller.UpdateDeliveryRecipientUrl(
            Guid.NewGuid(), new UpdateDeliveryRecipientUrlRequest("https://new"));

        result.Should().BeOfType<NotFoundObjectResult>();
    }

    [Test]
    public async Task UpdateDeliveryRecipientUrl_WhenFailure_ShouldReturnBadRequest()
    {
        _mediator.Send(Arg.Any<UpdateDeliveryRecipientUrlCommand>(), Arg.Any<CancellationToken>())
            .Returns(Result<UpdateDeliveryRecipientUrlResult>.Failure("zły url"));

        var result = await _controller.UpdateDeliveryRecipientUrl(
            Guid.NewGuid(), new UpdateDeliveryRecipientUrlRequest("bad"));

        result.Should().BeOfType<BadRequestObjectResult>();
    }
}

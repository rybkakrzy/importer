using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;
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
            new(Guid.NewGuid(), 1, DateTime.UtcNow, "User", true, 1024),
            new(Guid.NewGuid(), 2, DateTime.UtcNow, "User", false, 2048)
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
}

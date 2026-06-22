using D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UpdateCallbackUrlCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IReturnUrlValidator> _returnUrlValidator = null!;
    private UpdateCallbackUrlCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _returnUrlValidator = new Mock<IReturnUrlValidator>();
        _returnUrlValidator.Setup(v => v.Validate(It.IsAny<string>()))
            .Returns((string url) => ReturnUrlValidationResult.Success(url.Trim()));
        _handler = new UpdateCallbackUrlCommandHandler(_documentRepo.Object, _returnUrlValidator.Object);
    }

    private static Document NewDoc(string? metadata) =>
        new(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata);

    [Test]
    public async Task Handle_HappyPath_PersistsUrlAndPreservesClassification()
    {
        var doc = NewDoc("{\"returnUrl\":\"https://old.example.com/cb\",\"classification\":\"C2\"}");
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UpdateCallbackUrlCommand(doc.Id, "https://new.example.com/cb"),
            CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        doc.Metadata.Should().Contain("https://new.example.com/cb");
        doc.Metadata.Should().Contain("C2", "classification must be preserved when updating only the URL");
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_DocumentMissing_ReturnsNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(
            new UpdateCallbackUrlCommand(Guid.NewGuid(), "https://example.com/cb"),
            CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [TestCase("", TestName = "empty URL")]
    [TestCase("   ", TestName = "whitespace-only URL")]
    [TestCase("not-a-url", TestName = "non-absolute URL")]
    [TestCase("ftp://example.com/file", TestName = "non-http(s) scheme")]
    [TestCase("javascript:alert(1)", TestName = "javascript: scheme")]
    public async Task Handle_InvalidUrl_ReturnsFailureWithoutHittingRepo(string url)
    {
        _returnUrlValidator.Setup(v => v.Validate(url))
            .Returns(ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.InvalidAbsoluteUri,
                "Callback URL musi być absolutnym adresem http(s)."));

        var result = await _handler.Handle(
            new UpdateCallbackUrlCommand(Guid.NewGuid(), url),
            CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        _documentRepo.Verify(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()), Times.Never);
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [TestCase(DocumentStatus.Sending)]
    [TestCase(DocumentStatus.Sent)]
    [TestCase(DocumentStatus.DeliveryFailed)]
    public async Task Handle_DeliveryAlreadyInFlightOrDone_ReturnsConflictAndDoesNotSave(DocumentStatus status)
    {
        var doc = NewDoc("{\"returnUrl\":\"https://old.example.com/cb\",\"classification\":\"C2\"}");
        ForceStatus(doc, status);
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UpdateCallbackUrlCommand(doc.Id, "https://new.example.com/cb"),
            CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.Error.Should().Contain("Nie można zaktualizować callback URL");
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_RepeatCall_IsIdempotent_AndKeepsClassification()
    {
        var doc = NewDoc("{\"returnUrl\":\"https://example.com/cb\",\"classification\":\"C3\"}");
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var first = await _handler.Handle(
            new UpdateCallbackUrlCommand(doc.Id, "https://example.com/cb"),
            CancellationToken.None);
        var second = await _handler.Handle(
            new UpdateCallbackUrlCommand(doc.Id, "https://example.com/cb"),
            CancellationToken.None);

        first.IsSuccess.Should().BeTrue();
        second.IsSuccess.Should().BeTrue();
        doc.Metadata.Should().Contain("https://example.com/cb");
        doc.Metadata.Should().Contain("C3");
    }

    [Test]
    public async Task Handle_TrimsWhitespaceAroundUrl()
    {
        var doc = NewDoc(null);
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UpdateCallbackUrlCommand(doc.Id, "  https://example.com/cb  "),
            CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        doc.Metadata.Should().Contain("https://example.com/cb");
        doc.Metadata.Should().NotContain("  https");
    }

    private static void ForceStatus(Document doc, DocumentStatus status)
    {
        switch (status)
        {
            case DocumentStatus.Editing: doc.MarkEditing(); break;
            case DocumentStatus.Sending: doc.MarkSending(); break;
            case DocumentStatus.Sent: doc.MarkSent(); break;
            case DocumentStatus.DeliveryFailed: doc.MarkDeliveryFailed(); break;
            case DocumentStatus.Saved: doc.MarkSaved(); break;
        }
    }
}

using D2ViewerEditor.Application.Features.Documents.Commands.UnlockDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UnlockDocumentCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private UnlockDocumentCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _handler = new UnlockDocumentCommandHandler(
            _documentRepo.Object,
            NullLogger<UnlockDocumentCommandHandler>.Instance);
    }

    private static Document NewDoc() => new(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata: null);

    [Test]
    public async Task Handle_EditingDocument_TransitionsToSaved_ChangedTrue()
    {
        var doc = NewDoc();
        doc.MarkEditing();
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UnlockDocumentCommand(doc.Id, "manual"),
            CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Changed.Should().BeTrue();
        doc.Status.Should().Be(DocumentStatus.Saved);
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_AlreadySavedDocument_IsIdempotent_ChangedFalse_NoSave()
    {
        var doc = NewDoc(); // ctor sets Saved
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UnlockDocumentCommand(doc.Id, null),
            CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Changed.Should().BeFalse();
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_DocumentMissing_ReturnsNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(
            new UnlockDocumentCommand(Guid.NewGuid(), null),
            CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [TestCase(DocumentStatus.Sending)]
    [TestCase(DocumentStatus.Sent)]
    [TestCase(DocumentStatus.DeliveryFailed)]
    public async Task Handle_DeliveryStates_AreConflict(DocumentStatus status)
    {
        var doc = NewDoc();
        ForceStatus(doc, status);
        _documentRepo.Setup(r => r.GetByIdAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(
            new UnlockDocumentCommand(doc.Id, null),
            CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        result.Error.Should().Contain("Nie można odblokować");
        doc.Status.Should().Be(status, "delivery-owned states must not be silently overwritten");
        _documentRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
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

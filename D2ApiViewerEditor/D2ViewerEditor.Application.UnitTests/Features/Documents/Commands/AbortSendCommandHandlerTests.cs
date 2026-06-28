using D2ViewerEditor.Application.Features.Documents.Commands.AbortSend;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class AbortSendCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private AbortSendCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _handler = new AbortSendCommandHandler(_documentRepo.Object, _deliveryRepo.Object);
    }

    private static DocumentDelivery BuildHeldDelivery(Guid documentId)
    {
        var delivery = DocumentDelivery.Create(Guid.NewGuid(), documentId, Guid.NewGuid(),
            "deliveries/x", 10, "h", "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));
        delivery.BeginInlineAttempt();
        delivery.HoldAfterFailedInlineAttempt("boom"); // → RetryScheduled (held)
        return delivery;
    }

    [Test]
    public async Task Handle_ShouldCancelDeliveryAndMarkDocumentAborted()
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User");
        document.MarkDeliveryFailed();
        var delivery = BuildHeldDelivery(document.Id);

        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(new AbortSendCommand(document.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DocumentStatus.Should().Be(nameof(DocumentStatus.SendAborted));
        result.Value.DeliveryStatus.Should().Be(nameof(DeliveryStatus.Cancelled));
        document.Status.Should().Be(DocumentStatus.SendAborted);
        delivery.Status.Should().Be(DeliveryStatus.Cancelled);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_WhenInlineAttemptInProgress_ShouldCancelAndMarkAborted()
    {
        // „Przerwij wysyłkę" w trakcie trwającej próby inline (Sending bez lease) — użytkownik
        // może przerwać także wtedy, gdy wysyłka już trwa. CancelByUser → Cancelled (nie błąd).
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User");
        document.MarkSending();
        var delivery = DocumentDelivery.Create(Guid.NewGuid(), document.Id, Guid.NewGuid(),
            "deliveries/x", 10, "h", "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));
        delivery.BeginInlineAttempt(); // → Sending (inline, bez lease)

        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(new AbortSendCommand(document.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DeliveryStatus.Should().Be(nameof(DeliveryStatus.Cancelled));
        document.Status.Should().Be(DocumentStatus.SendAborted);
        delivery.Status.Should().Be(DeliveryStatus.Cancelled);
    }

    [Test]
    public async Task Handle_WithoutActiveDelivery_ShouldStillMarkDocumentAborted()
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User");
        document.MarkDeliveryFailed();

        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new AbortSendCommand(document.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DeliveryStatus.Should().BeNull();
        document.Status.Should().Be(DocumentStatus.SendAborted);
    }

    [Test]
    public async Task Handle_WhenDocumentMissing_ShouldReturnNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(new AbortSendCommand(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }
}

using D2ViewerEditor.Application.Features.Documents.Commands.ContinueDelivery;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class ContinueDeliveryCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private ContinueDeliveryCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _handler = new ContinueDeliveryCommandHandler(_documentRepo.Object, _deliveryRepo.Object);
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
    public async Task Handle_ShouldRequeueDeliveryAndMarkDocumentQueued()
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

        var result = await _handler.Handle(new ContinueDeliveryCommand(document.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DocumentStatus.Should().Be(nameof(DocumentStatus.Queued));
        result.Value.DeliveryStatus.Should().Be(nameof(DeliveryStatus.Pending));
        document.Status.Should().Be(DocumentStatus.Queued);
        delivery.Status.Should().Be(DeliveryStatus.Pending); // back in the worker queue, due now
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_WithoutActiveDelivery_ShouldFail()
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new ContinueDeliveryCommand(document.Id), CancellationToken.None);

        result.IsFailure.Should().BeTrue();
    }

    [Test]
    public async Task Handle_WhenDocumentMissing_ShouldReturnNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(new ContinueDeliveryCommand(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }
}

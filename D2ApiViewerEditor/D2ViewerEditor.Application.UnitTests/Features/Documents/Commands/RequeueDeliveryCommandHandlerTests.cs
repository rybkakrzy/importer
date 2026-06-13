using D2ViewerEditor.Application.Features.Documents.Commands.RequeueDelivery;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class RequeueDeliveryCommandHandlerTests
{
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private Mock<IDocumentRepository> _documentRepo = null!;
    private RequeueDeliveryCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _documentRepo = new Mock<IDocumentRepository>();
        _handler = new RequeueDeliveryCommandHandler(_deliveryRepo.Object, _documentRepo.Object);
    }

    private static DocumentDelivery DeadLetteredDelivery()
    {
        var delivery = DocumentDelivery.Create(
            Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), "deliveries/x", 10, "hash",
            "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromMinutes(1));
        delivery.ScheduleRetryOrDeadLetter("err", new FixedBackoff(TimeSpan.FromHours(2))); // → DeadLettered
        return delivery;
    }

    [Test]
    public async Task Handle_DeadLettered_ShouldRequeueAndMarkDocumentSending()
    {
        var delivery = DeadLetteredDelivery();
        var document = new Document(delivery.DocumentId, "doc.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "User", null);
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>())).ReturnsAsync(delivery);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(delivery.DocumentId, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);

        var result = await _handler.Handle(new RequeueDeliveryCommand(delivery.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Status.Should().Be(nameof(DeliveryStatus.Pending));
        delivery.Status.Should().Be(DeliveryStatus.Pending);
        document.Status.Should().Be(DocumentStatus.Sending);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_UnknownDelivery_ShouldReturnNotFound()
    {
        _deliveryRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new RequeueDeliveryCommand(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_AlreadySentDelivery_ShouldFail()
    {
        var delivery = DeadLetteredDelivery();
        delivery.Requeue(TimeSpan.FromHours(24));
        delivery.MarkSent(); // Sent → wznowienie niedozwolone
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>())).ReturnsAsync(delivery);

        var result = await _handler.Handle(new RequeueDeliveryCommand(delivery.Id), CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    private sealed class FixedBackoff : IBackoffStrategy
    {
        private readonly TimeSpan _delay;
        public FixedBackoff(TimeSpan delay) => _delay = delay;
        public DateTime NextAttempt(int attemptCount, DateTime now) => now.Add(_delay);
    }
}

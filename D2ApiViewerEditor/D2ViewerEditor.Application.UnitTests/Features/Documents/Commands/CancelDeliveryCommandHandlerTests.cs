using D2ViewerEditor.Application.Features.Documents.Commands.CancelDelivery;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class CancelDeliveryCommandHandlerTests
{
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private Mock<IDocumentRepository> _documentRepo = null!;
    private CancelDeliveryCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _documentRepo = new Mock<IDocumentRepository>();
        _handler = new CancelDeliveryCommandHandler(_deliveryRepo.Object, _documentRepo.Object);
    }

    private static DocumentDelivery PendingDelivery() =>
        DocumentDelivery.Create(
            Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), "deliveries/x", 10, "hash",
            "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));

    [Test]
    public async Task Handle_PendingDelivery_ShouldCancelAndPersist()
    {
        var delivery = PendingDelivery();
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(new CancelDeliveryCommand(delivery.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Status.Should().Be(nameof(DeliveryStatus.Cancelled));
        delivery.Status.Should().Be(DeliveryStatus.Cancelled);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_UnknownDelivery_ShouldReturnNotFound()
    {
        _deliveryRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new CancelDeliveryCommand(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_AlreadySentDelivery_ShouldFail()
    {
        var delivery = PendingDelivery();
        delivery.MarkSent(); // stan końcowy — anulować nie wolno
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(new CancelDeliveryCommand(delivery.Id), CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }
}

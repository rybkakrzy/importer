using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UpdateDeliveryRecipientUrlCommandHandlerTests
{
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private UpdateDeliveryRecipientUrlCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _handler = new UpdateDeliveryRecipientUrlCommandHandler(_deliveryRepo.Object);
    }

    private static DocumentDelivery PendingDelivery() =>
        DocumentDelivery.Create(
            Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), "deliveries/x", 10, "hash",
            "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));

    [Test]
    public async Task Handle_ValidUrl_ShouldUpdateAndPersist()
    {
        var delivery = PendingDelivery();
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(
            new UpdateDeliveryRecipientUrlCommand(delivery.Id, "  https://nowy.example.com/cb  "),
            CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.RecipientUrl.Should().Be("https://nowy.example.com/cb"); // przycięte
        delivery.RecipientUrl.Should().Be("https://nowy.example.com/cb");
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_InvalidUrl_ShouldFailWithoutPersist()
    {
        var delivery = PendingDelivery();
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(
            new UpdateDeliveryRecipientUrlCommand(delivery.Id, "not-a-url"), CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_UnknownDelivery_ShouldReturnNotFound()
    {
        _deliveryRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(
            new UpdateDeliveryRecipientUrlCommand(Guid.NewGuid(), "https://x.example.com"), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }
}

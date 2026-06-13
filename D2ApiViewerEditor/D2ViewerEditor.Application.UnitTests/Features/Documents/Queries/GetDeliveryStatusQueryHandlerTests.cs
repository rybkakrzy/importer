using D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveryStatus;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDeliveryStatusQueryHandlerTests
{
    private IDocumentDeliveryRepository _repo = null!;
    private GetDeliveryStatusQueryHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _repo = Substitute.For<IDocumentDeliveryRepository>();
        _handler = new GetDeliveryStatusQueryHandler(_repo);
    }

    private static DocumentDelivery Delivery() =>
        DocumentDelivery.Create(
            Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), "deliveries/x", 10, "hash",
            "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));

    [Test]
    public async Task Handle_ExistingDelivery_ReturnsMappedDto()
    {
        var delivery = Delivery();
        _repo.GetByIdAsync(delivery.Id, Arg.Any<CancellationToken>()).Returns(delivery);

        var result = await _handler.Handle(new GetDeliveryStatusQuery(delivery.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DeliveryId.Should().Be(delivery.Id);
        result.Value.DocumentId.Should().Be(delivery.DocumentId);
        result.Value.Status.Should().Be(delivery.Status.ToString());
        result.Value.AttemptCount.Should().Be(delivery.AttemptCount);
    }

    [Test]
    public async Task Handle_UnknownDelivery_ReturnsNotFound()
    {
        _repo.GetByIdAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>())
            .Returns((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDeliveryStatusQuery(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }
}

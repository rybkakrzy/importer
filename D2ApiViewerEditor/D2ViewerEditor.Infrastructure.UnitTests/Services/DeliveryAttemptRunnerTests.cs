using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Infrastructure.Services.Delivery;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Pojedyncza próba wysyłki zaclaimowanego zadania (status Sending): mapowanie wyniku sendera
/// na stan zadania (Sent / FailedPermanently / RetryScheduled) i status dokumentu.
/// </summary>
[TestFixture]
public class DeliveryAttemptRunnerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IDocumentStorageService> _storage = null!;
    private Mock<IDeliverySender> _sender = null!;
    private Mock<IBackoffStrategy> _backoff = null!;
    private DeliveryAttemptRunner _runner = null!;

    [SetUp]
    public void SetUp()
    {
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _documentRepo = new Mock<IDocumentRepository>();
        _storage = new Mock<IDocumentStorageService>();
        _sender = new Mock<IDeliverySender>();
        _backoff = new Mock<IBackoffStrategy>();
        _runner = new DeliveryAttemptRunner(
            _deliveryRepo.Object, _documentRepo.Object, _storage.Object,
            _sender.Object, _backoff.Object, NullLogger<DeliveryAttemptRunner>.Instance);
    }

    private static DocumentDelivery SendingDelivery(Guid documentId)
    {
        var delivery = DocumentDelivery.Create(Guid.NewGuid(), documentId, Guid.NewGuid(),
            "deliveries/x", 10, "h", "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));
        delivery.BeginInlineAttempt(); // → Sending (warunek wejścia runnera)
        return delivery;
    }

    private Document ArrangeDocument()
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User");
        document.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        return document;
    }

    private void ArrangeDelivery(DocumentDelivery delivery)
    {
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);
        _storage.Setup(s => s.DownloadAsync(delivery.SnapshotObjectName, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new byte[] { 1, 2, 3 });
    }

    [Test]
    public async Task RunAsync_WhenSenderSucceeds_ShouldMarkDeliveryAndDocumentSent()
    {
        var document = ArrangeDocument();
        var delivery = SendingDelivery(document.Id);
        ArrangeDelivery(delivery);
        _sender.Setup(s => s.SendAsync(It.IsAny<DeliveryDispatch>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(DeliveryResult.Succeeded());

        await _runner.RunAsync(delivery.Id, CancellationToken.None);

        delivery.Status.Should().Be(DeliveryStatus.Sent);
        document.Status.Should().Be(DocumentStatus.Sent);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task RunAsync_WhenSenderPermanentError_ShouldMarkDeliveryAndDocumentFailed()
    {
        var document = ArrangeDocument();
        var delivery = SendingDelivery(document.Id);
        ArrangeDelivery(delivery);
        _sender.Setup(s => s.SendAsync(It.IsAny<DeliveryDispatch>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(DeliveryResult.Permanent("bad recipient"));

        await _runner.RunAsync(delivery.Id, CancellationToken.None);

        delivery.Status.Should().Be(DeliveryStatus.FailedPermanently);
        document.Status.Should().Be(DocumentStatus.DeliveryFailed);
    }

    [Test]
    public async Task RunAsync_WhenRetryableWithinDeadline_ShouldScheduleRetryAndLeaveDocumentUntouched()
    {
        var document = ArrangeDocument();
        var delivery = SendingDelivery(document.Id);
        ArrangeDelivery(delivery);
        _sender.Setup(s => s.SendAsync(It.IsAny<DeliveryDispatch>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(DeliveryResult.Retryable("temporary"));
        _backoff.Setup(b => b.NextAttempt(It.IsAny<int>(), It.IsAny<DateTime>()))
            .Returns(DateTime.UtcNow.AddMinutes(1)); // przed deadline → retry, nie dead-letter

        await _runner.RunAsync(delivery.Id, CancellationToken.None);

        delivery.Status.Should().Be(DeliveryStatus.RetryScheduled);
        document.Status.Should().Be(DocumentStatus.Saved); // status dokumentu nietknięty przy retry
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task RunAsync_WhenDeliveryNotSending_ShouldNoOp()
    {
        var pending = DocumentDelivery.Create(Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(),
            "deliveries/x", 10, "h", "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));
        _deliveryRepo.Setup(r => r.GetByIdAsync(pending.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(pending); // status Pending — runner powinien wyjść

        await _runner.RunAsync(pending.Id, CancellationToken.None);

        _sender.Verify(s => s.SendAsync(It.IsAny<DeliveryDispatch>(), It.IsAny<CancellationToken>()), Times.Never);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task RunAsync_WhenDeliveryMissing_ShouldNoOp()
    {
        _deliveryRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        await _runner.RunAsync(Guid.NewGuid(), CancellationToken.None);

        _sender.Verify(s => s.SendAsync(It.IsAny<DeliveryDispatch>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task RunAsync_WhenStorageThrows_ShouldRescheduleAndSave()
    {
        var document = ArrangeDocument();
        var delivery = SendingDelivery(document.Id);
        _deliveryRepo.Setup(r => r.GetByIdAsync(delivery.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);
        _storage.Setup(s => s.DownloadAsync(delivery.SnapshotObjectName, It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("snapshot gone"));
        _backoff.Setup(b => b.NextAttempt(It.IsAny<int>(), It.IsAny<DateTime>()))
            .Returns(DateTime.UtcNow.AddMinutes(1));

        await _runner.RunAsync(delivery.Id, CancellationToken.None);

        delivery.Status.Should().Be(DeliveryStatus.RetryScheduled);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }
}

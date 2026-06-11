using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Entities;

[TestFixture]
public class DocumentDeliveryTests
{
    private const string ValidUrl = "https://example.com/return";

    private static DocumentDelivery CreateValid(TimeSpan? window = null) =>
        DocumentDelivery.Create(
            id: Guid.NewGuid(),
            documentId: Guid.NewGuid(),
            sourceVersionId: Guid.NewGuid(),
            snapshotObjectName: "deliveries/abc",
            snapshotSizeBytes: 1234,
            snapshotSha256: "ABCDEF",
            recipientUrl: ValidUrl,
            createdBy: "User",
            correlationId: Guid.NewGuid(),
            retentionWindow: window ?? TimeSpan.FromHours(24));

    [Test]
    public void Create_WithValidArgs_ShouldStartPendingWithDeadline()
    {
        var delivery = CreateValid();

        delivery.Status.Should().Be(DeliveryStatus.Pending);
        delivery.AttemptCount.Should().Be(0);
        delivery.DeadlineAt.Should().BeCloseTo(DateTime.UtcNow.AddHours(24), TimeSpan.FromSeconds(5));
        delivery.IsTerminal.Should().BeFalse();
    }

    [TestCase("")]
    [TestCase("ftp://example.com")]
    [TestCase("not-a-url")]
    [TestCase("/relative/path")]
    public void Create_WithInvalidRecipientUrl_ShouldThrow(string url)
    {
        var act = () => DocumentDelivery.Create(
            Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), "deliveries/x", 1, "h", url,
            "User", Guid.NewGuid(), TimeSpan.FromHours(24));

        act.Should().Throw<ArgumentException>();
    }

    [Test]
    public void MarkSent_ShouldClearLeaseAndError()
    {
        var delivery = CreateValid();

        delivery.MarkSent();

        delivery.Status.Should().Be(DeliveryStatus.Sent);
        delivery.LockedUntil.Should().BeNull();
        delivery.LastError.Should().BeNull();
        delivery.IsTerminal.Should().BeTrue();
    }

    [Test]
    public void MarkPermanentFailure_ShouldStoreError()
    {
        var delivery = CreateValid();

        delivery.MarkPermanentFailure("HTTP 422");

        delivery.Status.Should().Be(DeliveryStatus.FailedPermanently);
        delivery.LastError.Should().Be("HTTP 422");
        delivery.IsTerminal.Should().BeTrue();
    }

    [Test]
    public void ScheduleRetryOrDeadLetter_WithinDeadline_ShouldScheduleRetry()
    {
        var delivery = CreateValid();
        var backoff = new FixedBackoff(TimeSpan.FromMinutes(5));

        var retrying = delivery.ScheduleRetryOrDeadLetter("HTTP 503", backoff);

        retrying.Should().BeTrue();
        delivery.Status.Should().Be(DeliveryStatus.RetryScheduled);
        delivery.NextAttemptAt.Should().BeAfter(DateTime.UtcNow);
        delivery.LastError.Should().Be("HTTP 503");
    }

    [Test]
    public void ScheduleRetryOrDeadLetter_BeyondDeadline_ShouldDeadLetter()
    {
        var delivery = CreateValid(window: TimeSpan.FromMinutes(1));
        var backoff = new FixedBackoff(TimeSpan.FromHours(2)); // wykracza poza deadline

        var retrying = delivery.ScheduleRetryOrDeadLetter("HTTP 503", backoff);

        retrying.Should().BeFalse();
        delivery.Status.Should().Be(DeliveryStatus.DeadLettered);
        delivery.IsTerminal.Should().BeTrue();
    }

    [Test]
    public void Requeue_FromDeadLettered_ShouldResetToPending()
    {
        var delivery = CreateValid(window: TimeSpan.FromMinutes(1));
        delivery.ScheduleRetryOrDeadLetter("err", new FixedBackoff(TimeSpan.FromHours(2)));

        delivery.Requeue(TimeSpan.FromHours(24));

        delivery.Status.Should().Be(DeliveryStatus.Pending);
        delivery.LastError.Should().BeNull();
        delivery.DeadlineAt.Should().BeCloseTo(DateTime.UtcNow.AddHours(24), TimeSpan.FromSeconds(5));
    }

    [Test]
    public void Requeue_FromSent_ShouldThrow()
    {
        var delivery = CreateValid();
        delivery.MarkSent();

        var act = () => delivery.Requeue(TimeSpan.FromHours(24));

        act.Should().Throw<InvalidOperationException>();
    }

    [Test]
    public void Requeue_FromRetryScheduled_ShouldSendNow()
    {
        var delivery = CreateValid();
        delivery.ScheduleRetryOrDeadLetter("HTTP 503", new FixedBackoff(TimeSpan.FromMinutes(30)));

        delivery.Requeue(TimeSpan.FromHours(24));

        delivery.Status.Should().Be(DeliveryStatus.Pending);
        delivery.NextAttemptAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5)); // „teraz", nie za 30 min
        delivery.LastError.Should().BeNull();
    }

    [Test]
    public void Requeue_FromCancelled_ShouldResetToPending()
    {
        var delivery = CreateValid();
        delivery.Cancel();

        delivery.Requeue(TimeSpan.FromHours(24));

        delivery.Status.Should().Be(DeliveryStatus.Pending);
        delivery.IsTerminal.Should().BeFalse();
    }

    [Test]
    public void Cancel_FromPending_ShouldBecomeCancelledTerminal()
    {
        var delivery = CreateValid();

        delivery.Cancel();

        delivery.Status.Should().Be(DeliveryStatus.Cancelled);
        delivery.IsTerminal.Should().BeTrue();
        delivery.LockedUntil.Should().BeNull();
    }

    [Test]
    public void Cancel_FromRetryScheduled_ShouldBecomeCancelled()
    {
        var delivery = CreateValid();
        delivery.ScheduleRetryOrDeadLetter("err", new FixedBackoff(TimeSpan.FromMinutes(5)));

        delivery.Cancel();

        delivery.Status.Should().Be(DeliveryStatus.Cancelled);
    }

    [Test]
    public void Cancel_FromSent_ShouldThrow()
    {
        var delivery = CreateValid();
        delivery.MarkSent();

        var act = () => delivery.Cancel();

        act.Should().Throw<InvalidOperationException>();
    }

    [Test]
    public void UpdateRecipientUrl_WithValidUrl_ShouldChangeAddress()
    {
        var delivery = CreateValid();

        delivery.UpdateRecipientUrl("https://nowy.example.com/cb");

        delivery.RecipientUrl.Should().Be("https://nowy.example.com/cb");
    }

    [TestCase("")]
    [TestCase("ftp://example.com")]
    [TestCase("not-a-url")]
    public void UpdateRecipientUrl_WithInvalidUrl_ShouldThrow(string url)
    {
        var delivery = CreateValid();

        var act = () => delivery.UpdateRecipientUrl(url);

        act.Should().Throw<ArgumentException>();
    }

    [Test]
    public void UpdateRecipientUrl_WhenSent_ShouldThrow()
    {
        var delivery = CreateValid();
        delivery.MarkSent();

        var act = () => delivery.UpdateRecipientUrl("https://nowy.example.com/cb");

        act.Should().Throw<InvalidOperationException>();
    }

    private sealed class FixedBackoff : IBackoffStrategy
    {
        private readonly TimeSpan _delay;
        public FixedBackoff(TimeSpan delay) => _delay = delay;
        public DateTime NextAttempt(int attemptCount, DateTime now) => now.Add(_delay);
    }
}

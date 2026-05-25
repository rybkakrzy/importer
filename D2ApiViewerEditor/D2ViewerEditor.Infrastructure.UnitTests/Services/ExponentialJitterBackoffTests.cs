using D2ViewerEditor.Infrastructure.Services.Delivery;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

[TestFixture]
public class ExponentialJitterBackoffTests
{
    private ExponentialJitterBackoff _backoff = null!;

    [SetUp]
    public void SetUp() => _backoff = new ExponentialJitterBackoff();

    [Test]
    public void NextAttempt_ShouldAlwaysBeInFuture()
    {
        var now = DateTime.UtcNow;

        for (var attempt = 1; attempt <= 20; attempt++)
        {
            var next = _backoff.NextAttempt(attempt, now);
            next.Should().BeOnOrAfter(now);
        }
    }

    [Test]
    public void NextAttempt_ShouldNeverExceedCapOf15Minutes()
    {
        var now = DateTime.UtcNow;

        for (var attempt = 1; attempt <= 50; attempt++)
        {
            var next = _backoff.NextAttempt(attempt, now);
            (next - now).Should().BeLessThanOrEqualTo(TimeSpan.FromMinutes(15));
        }
    }
}

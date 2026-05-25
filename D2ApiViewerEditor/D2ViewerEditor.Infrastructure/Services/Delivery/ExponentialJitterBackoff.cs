using D2ViewerEditor.Domain.Interfaces;

namespace D2ViewerEditor.Infrastructure.Services.Delivery;

/// <summary>
/// Exponential backoff z full jitter: delay = random(0, min(cap, base * 2^(n-1))).
/// Pełny jitter rozprasza ponowienia wielu instancji (brak thundering herd),
/// a twardy limit 24 h egzekwuje warstwa domeny (DeadlineAt), nie ta strategia.
/// </summary>
public sealed class ExponentialJitterBackoff : IBackoffStrategy
{
    private static readonly TimeSpan Base = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan Cap = TimeSpan.FromMinutes(15);

    public DateTime NextAttempt(int attemptCount, DateTime now)
    {
        var exponent = Math.Max(0, attemptCount - 1);
        var ceilingSeconds = Math.Min(Cap.TotalSeconds, Base.TotalSeconds * Math.Pow(2, exponent));
        var delaySeconds = Random.Shared.NextDouble() * ceilingSeconds;
        return now.AddSeconds(delaySeconds);
    }
}

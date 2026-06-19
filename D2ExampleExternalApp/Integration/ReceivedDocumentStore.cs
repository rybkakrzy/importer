using System.Collections.Concurrent;

namespace D2ExampleExternalApp.Integration;

/// <summary>
/// Snapshot pojedynczego dokumentu odebranego w callbacku (returnUrl).
/// </summary>
public sealed record ReceivedDocument(
    string? MasterId,
    string? VersionId,
    string? CorporateKey,
    string FileName,
    long SizeBytes,
    string? IdempotencyKey,
    string? ExpectedSha256,
    string ComputedSha256,
    bool HashMatches,
    DateTime ReceivedAtUtc,
    string SavedPath);

/// <summary>
/// Pamięciowy rejestr odebranych callbacków — do podglądu w demie (GET /api/integration/received).
/// At-least-once: ten sam Idempotency-Key może przyjść wielokrotnie, więc deduplikujemy po nim.
/// </summary>
public static class ReceivedDocumentStore
{
    private static readonly ConcurrentDictionary<string, ReceivedDocument> ByIdempotencyKey = new();
    private static readonly ConcurrentQueue<ReceivedDocument> WithoutKey = new();

    /// <summary>Dodaje rekord; zwraca false, jeśli to duplikat (znany Idempotency-Key).</summary>
    public static bool Add(ReceivedDocument record)
    {
        if (string.IsNullOrEmpty(record.IdempotencyKey))
        {
            WithoutKey.Enqueue(record);
            return true;
        }

        return ByIdempotencyKey.TryAdd(record.IdempotencyKey, record);
    }

    public static IReadOnlyCollection<ReceivedDocument> Snapshot() =>
        ByIdempotencyKey.Values.Concat(WithoutKey)
            .OrderByDescending(r => r.ReceivedAtUtc)
            .ToList();
}

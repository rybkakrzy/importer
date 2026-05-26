using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Npgsql;

namespace D2ViewerEditor.Infrastructure.IntegrationTests;

/// <summary>
/// Testy integracyjne mechanizmu claimu kolejki wysyłek na REALNYM PostgreSQL.
/// Pokrywają to, czego testy jednostkowe nie są w stanie sprawdzić: semantykę
/// <c>FOR UPDATE SKIP LOCKED</c>, lease (locked_until/locked_by), reclaim po crashu,
/// pomijanie zadań niegotowych oraz unikalny indeks częściowy (idempotencja per dokument).
///
/// Uruchom: RUN_DB_INTEGRATION_TESTS=1 (+ działający PostgreSQL z migracjami 001..007).
/// </summary>
[TestFixture]
[Category("Integration")]
public class DocumentDeliveryRepositoryClaimTests
{
    private const string Worker = "worker-test";
    private static readonly TimeSpan Lease = TimeSpan.FromMinutes(5);

    private readonly List<Guid> _createdDocuments = new();

    [OneTimeSetUp]
    public void OneTimeSetUp() => TestDatabase.EnsureAvailableOrIgnore();

    [SetUp]
    public async Task SetUp()
    {
        await TestDatabase.ResetDeliveriesAsync();
        _createdDocuments.Clear();
    }

    [TearDown]
    public async Task TearDown()
    {
        await TestDatabase.ResetDeliveriesAsync();
        foreach (var docId in _createdDocuments)
            await TestDatabase.DeleteDocumentAsync(docId);
    }

    private async Task<Guid> NewDocumentAsync()
    {
        var id = await TestDatabase.InsertDocumentAsync();
        _createdDocuments.Add(id);
        return id;
    }

    private static DocumentDeliveryRepository NewRepository(out IDisposable scope)
    {
        var ctx = TestDatabase.CreateContext();
        scope = ctx;
        return new DocumentDeliveryRepository(ctx);
    }

    [Test]
    public async Task ClaimDueBatch_DuePendingTask_TransitionsToSendingWithLeaseAndAttemptIncrement()
    {
        var docId = await NewDocumentAsync();
        var delId = await TestDatabase.InsertDeliveryAsync(
            docId, status: "Pending", nextAttemptSql: "now() - interval '1 minute'");

        var repo = NewRepository(out var scope);
        using (scope)
        {
            var claimed = await repo.ClaimDueBatchAsync(batchSize: 100, lease: Lease, workerId: Worker);
            claimed.Should().ContainSingle(d => d.Id == delId);
        }

        var row = await TestDatabase.ReadDeliveryAsync(delId);
        row.Status.Should().Be(nameof(DeliveryStatus.Sending));
        row.AttemptCount.Should().Be(1, "claim inkrementuje attempt_count");
        row.LockedBy.Should().Be(Worker);
        row.LockedUntil.Should().NotBeNull();
        row.LockedUntil!.Value.Should().BeAfter(DateTime.UtcNow, "lease ustawiony w przyszłość");
    }

    [Test]
    public async Task ClaimDueBatch_RowLockedByAnotherTransaction_IsSkipped_ThenClaimableAfterRelease()
    {
        var docId = await NewDocumentAsync();
        var delId = await TestDatabase.InsertDeliveryAsync(
            docId, status: "Pending", nextAttemptSql: "now() - interval '1 minute'");

        // Konkurencyjny worker trzyma blokadę wiersza w otwartej transakcji (FOR UPDATE).
        await using var lockConn = TestDatabase.OpenConnection();
        await using var tx = await lockConn.BeginTransactionAsync();
        await using (var lockCmd = lockConn.CreateCommand())
        {
            lockCmd.Transaction = (NpgsqlTransaction)tx;
            lockCmd.CommandText = "SELECT id FROM document_deliveries WHERE id = @id FOR UPDATE;";
            lockCmd.Parameters.AddWithValue("id", delId);
            await lockCmd.ExecuteScalarAsync();
        }

        // Drugi worker: SKIP LOCKED → pomija zablokowany wiersz, nic nie claimuje.
        var repo1 = NewRepository(out var scope1);
        using (scope1)
        {
            var claimedWhileLocked = await repo1.ClaimDueBatchAsync(100, Lease, "worker-B");
            claimedWhileLocked.Should().NotContain(d => d.Id == delId, "wiersz jest zablokowany przez inną transakcję");
        }

        (await TestDatabase.ReadDeliveryAsync(delId)).Status
            .Should().Be("Pending", "pominięty wiersz pozostaje nietknięty");

        // Zwolnienie blokady → wiersz znów do wzięcia.
        await tx.RollbackAsync();

        var repo2 = NewRepository(out var scope2);
        using (scope2)
        {
            var claimedAfterRelease = await repo2.ClaimDueBatchAsync(100, Lease, "worker-B");
            claimedAfterRelease.Should().ContainSingle(d => d.Id == delId);
        }
    }

    [Test]
    public async Task ClaimDueBatch_StuckSendingWithExpiredLease_IsReclaimed()
    {
        var docId = await NewDocumentAsync();
        // Symulacja crashu workera: zadanie utknęło w Sending z wygasłym lease.
        var delId = await TestDatabase.InsertDeliveryAsync(
            docId,
            status: "Sending",
            nextAttemptSql: "now() - interval '10 minutes'",
            lockedUntilSql: "now() - interval '1 minute'",
            lockedBy: "dead-worker",
            attemptCount: 1);

        var repo = NewRepository(out var scope);
        using (scope)
        {
            var claimed = await repo.ClaimDueBatchAsync(100, Lease, Worker);
            claimed.Should().ContainSingle(d => d.Id == delId, "zadanie z wygasłym lease ma zostać przejęte");
        }

        var row = await TestDatabase.ReadDeliveryAsync(delId);
        row.Status.Should().Be(nameof(DeliveryStatus.Sending));
        row.AttemptCount.Should().Be(2, "reclaim również inkrementuje attempt_count");
        row.LockedBy.Should().Be(Worker, "lease przejęty przez nowego workera");
        row.LockedUntil!.Value.Should().BeAfter(DateTime.UtcNow);
    }

    [Test]
    public async Task ClaimDueBatch_NotDueOrActivelyLeased_AreNotClaimed()
    {
        var docFuture = await NewDocumentAsync();
        var notDueId = await TestDatabase.InsertDeliveryAsync(
            docFuture, status: "RetryScheduled", nextAttemptSql: "now() + interval '10 minutes'");

        var docLeased = await NewDocumentAsync();
        var activelyLeasedId = await TestDatabase.InsertDeliveryAsync(
            docLeased,
            status: "Sending",
            nextAttemptSql: "now() - interval '10 minutes'",
            lockedUntilSql: "now() + interval '5 minutes'",
            lockedBy: "busy-worker",
            attemptCount: 1);

        var repo = NewRepository(out var scope);
        using (scope)
        {
            var claimed = await repo.ClaimDueBatchAsync(100, Lease, Worker);
            claimed.Should().BeEmpty("ani zadanie niegotowe, ani aktywnie dzierżawione nie podlega claimowi");
        }

        (await TestDatabase.ReadDeliveryAsync(notDueId)).Status.Should().Be("RetryScheduled");
        var leased = await TestDatabase.ReadDeliveryAsync(activelyLeasedId);
        leased.LockedBy.Should().Be("busy-worker", "aktywny lease nie został naruszony");
    }

    [Test]
    public async Task ClaimDueBatch_RespectsBatchSize()
    {
        // 3 różne dokumenty (unikalny indeks pozwala tylko na jedno aktywne zadanie per dokument).
        for (var i = 0; i < 3; i++)
        {
            var docId = await NewDocumentAsync();
            await TestDatabase.InsertDeliveryAsync(docId, "Pending", "now() - interval '1 minute'");
        }

        var repo = NewRepository(out var scope);
        using (scope)
        {
            var claimed = await repo.ClaimDueBatchAsync(batchSize: 2, lease: Lease, workerId: Worker);
            claimed.Should().HaveCount(2, "batch ograniczony do 2");
        }

        // Trzeci pozostaje gotowy do następnej tury.
        var repo2 = NewRepository(out var scope2);
        using (scope2)
        {
            var remaining = await repo2.ClaimDueBatchAsync(batchSize: 10, lease: Lease, workerId: Worker);
            remaining.Should().HaveCount(1);
        }
    }

    [Test]
    public async Task UniquePartialIndex_PreventsSecondActiveDeliveryPerDocument()
    {
        var docId = await NewDocumentAsync();
        await TestDatabase.InsertDeliveryAsync(docId, "Pending", "now() - interval '1 minute'");

        var act = async () =>
            await TestDatabase.InsertDeliveryAsync(docId, "Pending", "now() - interval '1 minute'");

        (await act.Should().ThrowAsync<PostgresException>())
            .Which.SqlState.Should().Be(PostgresErrorCodes.UniqueViolation,
                "unikalny indeks częściowy gwarantuje jedno aktywne zadanie na dokument (idempotencja)");
    }
}

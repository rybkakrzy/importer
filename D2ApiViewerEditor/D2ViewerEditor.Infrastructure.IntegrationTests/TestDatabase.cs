using D2ViewerEditor.Infrastructure.Persistence;
using Microsoft.EntityFrameworkCore;
using Npgsql;

namespace D2ViewerEditor.Infrastructure.IntegrationTests;

/// <summary>
/// Pomocnik do testów integracyjnych na REALNYM PostgreSQL (lokalny Podman / fake-gcs niepotrzebny).
///
/// Bezpieczeństwo i izolacja:
/// - Testy uruchamiają się TYLKO gdy ustawiono <c>RUN_DB_INTEGRATION_TESTS=1</c> — inaczej <see cref="EnsureAvailableOrIgnore"/>
///   wywołuje <c>Assert.Ignore</c>. Dzięki temu zwykły <c>dotnet test</c> / CI bez bazy nie pada i nie czyści danych.
/// - Connection string z env <c>D2_TEST_DB</c>, domyślnie lokalna baza dev na Podman.
/// - Każdy test czyści tabelę <c>document_deliveries</c> (patrz <see cref="ResetDeliveriesAsync"/>), więc operuje na
///   deterministycznym, pustym zbiorze — claim (globalny <c>FOR UPDATE SKIP LOCKED</c>) jest przez to powtarzalny.
/// </summary>
internal static class TestDatabase
{
    public const string DefaultConnectionString =
        "Host=localhost;Port=5432;Database=d2viewereditor;Username=postgres;Password=postgres;Include Error Detail=true";

    public static string ConnectionString =>
        Environment.GetEnvironmentVariable("D2_TEST_DB") is { Length: > 0 } cs ? cs : DefaultConnectionString;

    private static bool OptedIn =>
        Environment.GetEnvironmentVariable("RUN_DB_INTEGRATION_TESTS") is "1" or "true" or "TRUE";

    /// <summary>
    /// Sprawdza warunki uruchomienia; jeśli niespełnione — pomija fixture przez Assert.Ignore (nie failuje).
    /// </summary>
    public static void EnsureAvailableOrIgnore()
    {
        if (!OptedIn)
            Assert.Ignore("Testy integracyjne DB pominięte. Ustaw RUN_DB_INTEGRATION_TESTS=1 i uruchom PostgreSQL (Podman) aby je włączyć.");

        try
        {
            using var conn = new NpgsqlConnection(ConnectionString);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT to_regclass('public.document_deliveries');";
            var exists = cmd.ExecuteScalar();
            if (exists is null or DBNull)
                Assert.Ignore("Tabela document_deliveries nie istnieje — uruchom migracje infra/sql/007.");
        }
        catch (NpgsqlException ex)
        {
            Assert.Ignore($"PostgreSQL niedostępny pod {ConnectionString}: {ex.Message}");
        }
    }

    public static DocumentDbContext CreateContext()
    {
        var options = new DbContextOptionsBuilder<DocumentDbContext>()
            .UseNpgsql(ConnectionString)
            .Options;
        return new DocumentDbContext(options);
    }

    public static NpgsqlConnection OpenConnection()
    {
        var conn = new NpgsqlConnection(ConnectionString);
        conn.Open();
        return conn;
    }

    /// <summary>Czyści kolejkę wysyłek (test gated env-varem → operujemy na bazie dev/test).</summary>
    public static async Task ResetDeliveriesAsync()
    {
        await using var conn = new NpgsqlConnection(ConnectionString);
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM document_deliveries;";
        await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>Wstawia rekord <c>documents</c> (FK dla wysyłki) i zwraca jego id.</summary>
    public static async Task<Guid> InsertDocumentAsync()
    {
        var id = Guid.NewGuid();
        await using var conn = new NpgsqlConnection(ConnectionString);
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO documents (id, name, mime_type, created_at, created_by)
            VALUES (@id, 'integration-test', 'text/plain', now(), 'integration-test');";
        cmd.Parameters.AddWithValue("id", id);
        await cmd.ExecuteNonQueryAsync();
        return id;
    }

    public static async Task DeleteDocumentAsync(Guid documentId)
    {
        await using var conn = new NpgsqlConnection(ConnectionString);
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        // ON DELETE CASCADE usunie powiązane wysyłki.
        cmd.CommandText = "DELETE FROM documents WHERE id = @id;";
        cmd.Parameters.AddWithValue("id", documentId);
        await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>
    /// Wstawia zadanie wysyłki z pełną kontrolą stanu. Kolumny czasowe przyjmują wyrażenia SQL
    /// (np. <c>now() - interval '1 minute'</c>), bo są to wartości testowe — nie dane użytkownika.
    /// </summary>
    public static async Task<Guid> InsertDeliveryAsync(
        Guid documentId,
        string status,
        string nextAttemptSql,
        string deadlineSql = "now() + interval '24 hours'",
        string? lockedUntilSql = null,
        string? lockedBy = null,
        int attemptCount = 0)
    {
        var id = Guid.NewGuid();
        await using var conn = new NpgsqlConnection(ConnectionString);
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $@"
            INSERT INTO document_deliveries (
                id, document_id, source_version_id, snapshot_object_name, snapshot_size_bytes,
                snapshot_sha256, recipient_url, status, attempt_count,
                created_at, updated_at, next_attempt_at, deadline_at,
                locked_until, locked_by, correlation_id, created_by)
            VALUES (
                @id, @documentId, @sourceVersionId, @snapshotName, 10,
                'abc123', 'https://example.test/recipient', @status, @attemptCount,
                now(), now(), {nextAttemptSql}, {deadlineSql},
                {(lockedUntilSql ?? "NULL")}, @lockedBy, @correlationId, 'integration-test');";
        cmd.Parameters.AddWithValue("id", id);
        cmd.Parameters.AddWithValue("documentId", documentId);
        cmd.Parameters.AddWithValue("sourceVersionId", Guid.NewGuid());
        cmd.Parameters.AddWithValue("snapshotName", $"deliveries/{id}");
        cmd.Parameters.AddWithValue("status", status);
        cmd.Parameters.AddWithValue("attemptCount", attemptCount);
        cmd.Parameters.AddWithValue("lockedBy", (object?)lockedBy ?? DBNull.Value);
        cmd.Parameters.AddWithValue("correlationId", Guid.NewGuid());
        await cmd.ExecuteNonQueryAsync();
        return id;
    }

    public static async Task<(string Status, int AttemptCount, DateTime? LockedUntil, string? LockedBy)> ReadDeliveryAsync(Guid id)
    {
        await using var conn = new NpgsqlConnection(ConnectionString);
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT status, attempt_count, locked_until, locked_by FROM document_deliveries WHERE id = @id;";
        cmd.Parameters.AddWithValue("id", id);
        await using var reader = await cmd.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
            throw new InvalidOperationException($"Delivery {id} nie istnieje.");

        var status = reader.GetString(0);
        var attempt = reader.GetInt32(1);
        var lockedUntil = reader.IsDBNull(2) ? (DateTime?)null : reader.GetDateTime(2);
        var lockedBy = reader.IsDBNull(3) ? null : reader.GetString(3);
        return (status, attempt, lockedUntil, lockedBy);
    }
}

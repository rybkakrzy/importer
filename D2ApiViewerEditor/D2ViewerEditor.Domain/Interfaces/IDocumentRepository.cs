using D2ViewerEditor.Domain.Entities;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Repository dla zarządzania dokumentami
/// </summary>
public interface IDocumentRepository
{
    /// <summary>
    /// Pobiera dokument po ID (guid_master)
    /// </summary>
    Task<Document?> GetByIdAsync(Guid id, CancellationToken cancellationToken = default);

    /// <summary>
    /// Pobiera dokument po ID z załadowanymi wersjami
    /// </summary>
    Task<Document?> GetByIdWithVersionsAsync(Guid id, CancellationToken cancellationToken = default);

    /// <summary>
    /// Pobiera wszystkie dokumenty (z paginacją)
    /// </summary>
    Task<IReadOnlyList<Document>> GetAllAsync(int skip = 0, int take = 50, CancellationToken cancellationToken = default);

    /// <summary>
    /// Dodaje nowy dokument
    /// </summary>
    Task AddAsync(Document document, CancellationToken cancellationToken = default);

    /// <summary>
    /// Aktualizuje dokument
    /// </summary>
    Task UpdateAsync(Document document, CancellationToken cancellationToken = default);

    /// <summary>
    /// Zapisuje zmiany w bazie danych
    /// </summary>
    Task<int> SaveChangesAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Sprawdza czy dokument istnieje
    /// </summary>
    Task<bool> ExistsAsync(Guid id, CancellationToken cancellationToken = default);
}

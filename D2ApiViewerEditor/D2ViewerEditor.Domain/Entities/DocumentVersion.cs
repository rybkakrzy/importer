namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Wersja dokumentu - Entity
/// Reprezentuje konkretną wersję z referencją do pliku w storage (GCS)
/// </summary>
public class DocumentVersion
{
    private DocumentVersion() { } // EF Core

    public DocumentVersion(Guid id, Guid documentId, string storagePath, long sizeInBytes, int versionNumber, string createdBy)
    {
        if (string.IsNullOrWhiteSpace(storagePath))
            throw new ArgumentException("Ścieżka storage nie może być pusta", nameof(storagePath));

        Id = id;
        DocumentId = documentId;
        StoragePath = storagePath;
        SizeInBytes = sizeInBytes;
        VersionNumber = versionNumber;
        CreatedAt = DateTime.UtcNow;
        CreatedBy = createdBy;
        IsActive = true; // Nowa wersja zawsze jest aktywna
    }

    /// <summary>
    /// GUID wersji (guid_wersji)
    /// </summary>
    public Guid Id { get; private set; }

    /// <summary>
    /// Referencja do dokumentu głównego
    /// </summary>
    public Guid DocumentId { get; private set; }

    /// <summary>
    /// Ścieżka do pliku w GCS (np. "documents/{versionId}")
    /// </summary>
    public string StoragePath { get; private set; } = string.Empty;

    /// <summary>
    /// Rozmiar zawartości w bajtach
    /// </summary>
    public long SizeInBytes { get; private set; }

    /// <summary>
    /// Numer wersji (1, 2, 3, ...)
    /// </summary>
    public int VersionNumber { get; private set; }

    /// <summary>
    /// Data utworzenia wersji
    /// </summary>
    public DateTime CreatedAt { get; private set; }

    /// <summary>
    /// Kto stworzył wersję
    /// </summary>
    public string CreatedBy { get; private set; } = string.Empty;

    /// <summary>
    /// Czy ta wersja jest aktywna (tylko jedna wersja może być aktywna)
    /// </summary>
    public bool IsActive { get; private set; }

    /// <summary>
    /// Navigation property do dokumentu
    /// </summary>
    public Document? Document { get; private set; }

    /// <summary>
    /// Aktywuje tę wersję
    /// </summary>
    internal void Activate()
    {
        IsActive = true;
    }

    /// <summary>
    /// Dezaktywuje tę wersję
    /// </summary>
    internal void Deactivate()
    {
        IsActive = false;
    }
}

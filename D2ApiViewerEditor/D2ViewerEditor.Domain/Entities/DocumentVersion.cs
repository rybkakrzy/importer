namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Wersja dokumentu - Entity
/// Reprezentuje konkretną wersję z zawartością binarną
/// </summary>
public class DocumentVersion
{
    private DocumentVersion() { } // EF Core

    public DocumentVersion(Guid id, Guid documentId, byte[] content, int versionNumber, string createdBy)
    {
        if (content == null || content.Length == 0)
            throw new ArgumentException("Zawartość nie może być pusta", nameof(content));

        Id = id;
        DocumentId = documentId;
        Content = content;
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
    /// Zawartość binarna dokumentu (DOCX, PDF)
    /// </summary>
    public byte[] Content { get; private set; } = Array.Empty<byte>();

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
    /// Rozmiar zawartości w bajtach
    /// </summary>
    public long SizeInBytes => Content?.Length ?? 0;

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

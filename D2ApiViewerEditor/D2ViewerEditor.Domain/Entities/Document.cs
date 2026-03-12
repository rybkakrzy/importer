namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Dokument - Aggregate Root
/// Reprezentuje główny dokument z metadanymi i wersjami
/// </summary>
public class Document
{
    private readonly List<DocumentVersion> _versions = new();

    private Document() { } // EF Core

    public Document(Guid id, string name, string mimeType, string createdBy)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Nazwa dokumentu nie może być pusta", nameof(name));

        if (string.IsNullOrWhiteSpace(mimeType))
            throw new ArgumentException("MimeType nie może być pusty", nameof(mimeType));

        Id = id;
        Name = name;
        MimeType = mimeType;
        CreatedAt = DateTime.UtcNow;
        CreatedBy = createdBy;
        IsDeleted = false;
    }

    /// <summary>
    /// GUID główny dokumentu (guid_master)
    /// </summary>
    public Guid Id { get; private set; }

    /// <summary>
    /// Nazwa dokumentu
    /// </summary>
    public string Name { get; private set; } = string.Empty;

    /// <summary>
    /// Typ MIME (application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/pdf)
    /// </summary>
    public string MimeType { get; private set; } = string.Empty;

    /// <summary>
    /// Data utworzenia dokumentu
    /// </summary>
    public DateTime CreatedAt { get; private set; }

    /// <summary>
    /// Kto stworzył dokument
    /// </summary>
    public string CreatedBy { get; private set; } = string.Empty;

    /// <summary>
    /// Soft delete flag
    /// </summary>
    public bool IsDeleted { get; private set; }

    /// <summary>
    /// Wersje dokumentu
    /// </summary>
    public IReadOnlyCollection<DocumentVersion> Versions => _versions.AsReadOnly();

    /// <summary>
    /// Dodaje nową wersję dokumentu
    /// </summary>
    public DocumentVersion AddVersion(byte[] content, string createdBy)
    {
        if (content == null || content.Length == 0)
            throw new ArgumentException("Zawartość dokumentu nie może być pusta", nameof(content));

        var versionNumber = _versions.Count + 1;
        var version = new DocumentVersion(
            id: Guid.NewGuid(),
            documentId: Id,
            content: content,
            versionNumber: versionNumber,
            createdBy: createdBy
        );

        // Dezaktywuj poprzednie wersje
        foreach (var v in _versions)
        {
            v.Deactivate();
        }

        _versions.Add(version);
        return version;
    }

    /// <summary>
    /// Przywraca wybraną wersję jako aktywną
    /// </summary>
    public void RestoreVersion(Guid versionId)
    {
        var versionToRestore = _versions.FirstOrDefault(v => v.Id == versionId);
        if (versionToRestore == null)
            throw new InvalidOperationException($"Wersja {versionId} nie istnieje");

        foreach (var v in _versions)
        {
            v.Deactivate();
        }

        versionToRestore.Activate();
    }

    /// <summary>
    /// Pobiera aktywną wersję dokumentu
    /// </summary>
    public DocumentVersion? GetActiveVersion()
    {
        return _versions.FirstOrDefault(v => v.IsActive);
    }

    /// <summary>
    /// Zmienia nazwę dokumentu
    /// </summary>
    public void ChangeName(string newName)
    {
        if (string.IsNullOrWhiteSpace(newName))
            throw new ArgumentException("Nazwa nie może być pusta", nameof(newName));

        Name = newName;
    }

    /// <summary>
    /// Soft delete dokumentu
    /// </summary>
    public void Delete()
    {
        IsDeleted = true;
    }
}

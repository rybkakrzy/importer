namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Abstrakcja przechowywania plików binarnych dokumentów (GCS / fake-gcs-server)
/// </summary>
public interface IDocumentStorageService
{
    /// <summary>
    /// Zapisuje plik i zwraca ścieżkę w storage (object name)
    /// </summary>
    Task<string> UploadAsync(Guid versionId, byte[] content, string mimeType, CancellationToken cancellationToken = default);

    /// <summary>
    /// Zapisuje plik pod jawnie podaną nazwą obiektu (np. niezmienny snapshot "deliveries/{id}") i zwraca tę nazwę.
    /// Idempotentny: ponowny zapis tej samej nazwy nadpisuje obiekt.
    /// </summary>
    Task<string> UploadRawAsync(string objectName, byte[] content, string mimeType, CancellationToken cancellationToken = default);

    /// <summary>
    /// Pobiera plik binarny ze storage
    /// </summary>
    Task<byte[]> DownloadAsync(string storagePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Usuwa plik ze storage
    /// </summary>
    Task DeleteAsync(string storagePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Sprawdza czy plik istnieje w storage
    /// </summary>
    Task<bool> ExistsAsync(string storagePath, CancellationToken cancellationToken = default);
}

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Konfiguracja Google Cloud Storage
/// </summary>
public class GcsStorageOptions
{
    public const string SectionName = "GoogleCloudStorage";

    /// <summary>
    /// Nazwa bucketu GCS
    /// </summary>
    public string BucketName { get; set; } = string.Empty;

    /// <summary>
    /// Endpoint API (pusty = produkcyjny GCS, ustawiony = fake-gcs-server dla DEV)
    /// </summary>
    public string? ApiEndpoint { get; set; }

    /// <summary>
    /// Ścieżka do pliku service account JSON (opcjonalnie, produkcja może używać Workload Identity)
    /// </summary>
    public string? CredentialPath { get; set; }
}

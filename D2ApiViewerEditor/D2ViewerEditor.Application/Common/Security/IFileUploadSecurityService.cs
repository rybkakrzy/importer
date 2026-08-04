namespace D2ViewerEditor.Application.Common.Security;

public enum UploadRejectionCode
{
    None = 0,
    Empty,
    UnsupportedExtension,
    UnsupportedMimeType,
    ExtensionMimeMismatch,
    SignatureMismatch,
    DocxInvalidArchive,
    DocxZipSlipRisk,
    DocxTooManyEntries,
    DocxTooLargeUncompressed,
    DocxCompressionRatioExceeded,
    DocxRequiredPartMissing,
    DocxSuspiciousEntry,
    MalwareDetected,
    MalwareScanUnavailable
}

public sealed record UploadValidationResult(
    bool IsValid,
    UploadRejectionCode Code,
    string? Error,
    string? NormalizedMimeType)
{
    public static UploadValidationResult Success(string normalizedMimeType) =>
        new(true, UploadRejectionCode.None, null, normalizedMimeType);

    public static UploadValidationResult Failure(UploadRejectionCode code, string error) =>
        new(false, code, error, null);
}

public interface IFileUploadSecurityService
{
    Task<UploadValidationResult> ValidateDocumentAsync(
        byte[] content,
        string fileName,
        string mimeType,
        CancellationToken cancellationToken);

    Task<UploadValidationResult> ValidateImageAsync(
        byte[] content,
        string fileName,
        string contentType,
        CancellationToken cancellationToken);

    /// <summary>
    /// Sama walidacja STRUKTURY DOCX (podpis ZIP + poprawność archiwum + wymagane części OOXML),
    /// bez MIME/rozszerzenia i bez skanera AV. Dla ścieżek, które mają już bajty po normalizacji
    /// (np. POST /open — bug 13625398: uszkodzony plik otwierał się „po cichu" jako pusty dokument,
    /// podczas gdy upload ze strony startowej go odrzucał).
    /// </summary>
    UploadValidationResult ValidateDocxStructure(byte[] content);
}

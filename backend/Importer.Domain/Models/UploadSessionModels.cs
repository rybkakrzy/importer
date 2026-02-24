namespace Importer.Domain.Models;

public record UploadSession(
    Guid Id,
    string FileName,
    string Status,
    DateTime StartedAt,
    DateTime? CompletedAt,
    int? FilesCount,
    long? TotalSize,
    string? ErrorMessage
);

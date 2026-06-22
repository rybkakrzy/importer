namespace D2ViewerEditor.Application.Common.Security;

public enum FileScanStatus
{
    Clean = 0,
    Infected,
    Unavailable,
    Error
}

public sealed record FileScanRequest(string FileName, string MimeType, byte[] Content);

public sealed record FileScanResult(FileScanStatus Status, string? Signature = null, string? Details = null)
{
    public static FileScanResult Clean() => new(FileScanStatus.Clean);
}

public interface IFileScanner
{
    Task<FileScanResult> ScanAsync(FileScanRequest request, CancellationToken cancellationToken);
}

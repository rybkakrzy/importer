namespace D2ViewerEditor.Application.Common.Security;

public sealed class NoOpFileScanner : IFileScanner
{
    public Task<FileScanResult> ScanAsync(FileScanRequest request, CancellationToken cancellationToken)
    {
        return Task.FromResult(FileScanResult.Clean());
    }
}

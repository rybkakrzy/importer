namespace D2ViewerEditor.Application.Common.Security;

public sealed class UploadSecurityOptions
{
    public const string SectionName = "Security:Upload";

    public int MaxDocxEntryCount { get; set; } = 4096;
    public long MaxDocxUncompressedBytes { get; set; } = 120L * 1024 * 1024;
    public double MaxCompressionRatio { get; set; } = 100;
    public bool RejectDocxMacros { get; set; } = true;
    public bool FailClosedWhenScannerFails { get; set; } = true;
}

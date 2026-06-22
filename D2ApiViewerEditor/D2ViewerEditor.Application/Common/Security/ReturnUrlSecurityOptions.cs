namespace D2ViewerEditor.Application.Common.Security;

public sealed class ReturnUrlSecurityOptions
{
    public const string SectionName = "Security:ReturnUrl";

    public int MaxLength { get; set; } = 2048;
    public bool RequireHttps { get; set; } = true;
    public bool AllowLoopback { get; set; }
    public bool AllowPrivateNetworkIp { get; set; }
    public string[] AllowedHosts { get; set; } = [];
}

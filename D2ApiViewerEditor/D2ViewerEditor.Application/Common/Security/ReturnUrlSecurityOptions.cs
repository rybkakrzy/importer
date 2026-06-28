namespace D2ViewerEditor.Application.Common.Security;

public sealed class ReturnUrlSecurityOptions
{
    public const string SectionName = "Security:ReturnUrl";

    /// <summary>Maksymalna długość adresu callbacku (returnUrl). Powyżej → odrzucenie.</summary>
    public int MaxLength { get; set; } = 2048;

    /// <summary>Gdy true — wymusza https i odrzuca http. Domyślnie false: http i https są równoprawne
    /// (sam http nie jest powodem odrzucenia). Każde środowisko może to nadpisać w appsettings.</summary>
    public bool RequireHttps { get; set; } = false;

    /// <summary>Czy dopuścić host pętli zwrotnej (localhost / 127.0.0.1).</summary>
    public bool AllowLoopback { get; set; }

    /// <summary>Czy dopuścić prywatne / specjalne adresy IP (10.x, 192.168.x, 172.16-31.x, 169.254.x).</summary>
    public bool AllowPrivateNetworkIp { get; set; }

    /// <summary>Allowlista hostów odbiorców (pusta = dowolny host przechodzi pozostałe reguły).</summary>
    public string[] AllowedHosts { get; set; } = [];
}

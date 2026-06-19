namespace D2ExampleExternalApp.Configuration;

/// <summary>
/// Konfiguracja połączenia do D2ServicesViewerEditor (API integracji zewnętrznej).
/// </summary>
public sealed class D2ServicesOptions
{
    public const string SectionName = "D2Services";
    public const string HttpClientName = "D2Services";

    /// <summary>Bazowy URL D2Services, np. http://localhost:15112.</summary>
    public string BaseUrl { get; set; } = "http://localhost:15112";

    /// <summary>Opcjonalny klucz API; jeśli ustawiony, wysyłany jako nagłówek X-Api-Key.</summary>
    public string? ApiKey { get; set; }
}

/// <summary>
/// Konfiguracja samej aplikacji przykładowej (jak D2 ma do nas oddzwonić).
/// </summary>
public sealed class ExternalAppOptions
{
    public const string SectionName = "ExternalApp";

    /// <summary>
    /// Publiczny URL TEJ aplikacji — z niego budujemy returnUrl wysyłany do D2.
    /// Musi być osiągalny z poziomu D2Services (lokalnie zwykle ten sam host/port co aplikacja).
    /// </summary>
    public string PublicBaseUrl { get; set; } = "http://localhost:15120";

    /// <summary>Katalog, do którego zapisujemy pliki odebrane w callbacku.</summary>
    public string ReceivedFilesPath { get; set; } = "received";
}

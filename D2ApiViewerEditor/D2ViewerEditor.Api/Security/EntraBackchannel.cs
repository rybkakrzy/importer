using System.Net;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Builds the enterprise forward proxy used to reach Microsoft Entra ID from behind a corporate
/// proxy (JwtBearer metadata / JWKS fetch and downstream HTTP). Configured only when
/// <c>AzureAd:Proxy:Url</c> is set; otherwise null → direct connection (local / no proxy).
/// Mirrors the D2WebCore pattern. Google Cloud Storage (<c>storage.googleapis.com</c>) is always
/// bypassed so document storage traffic stays direct, not routed through the corporate proxy.
/// </summary>
public static class EntraBackchannel
{
    /// <summary>The proxy from <c>AzureAd:Proxy:Url</c>, or null when none is configured.</summary>
    public static WebProxy? CreateProxy(AzureAdOptions options)
    {
        var url = options.Proxy?.Url;
        if (string.IsNullOrWhiteSpace(url))
        {
            return null;
        }

        return new WebProxy
        {
            Address = new Uri(url),
            BypassList = ["storage.googleapis.com"],
            UseDefaultCredentials = true,
        };
    }
}

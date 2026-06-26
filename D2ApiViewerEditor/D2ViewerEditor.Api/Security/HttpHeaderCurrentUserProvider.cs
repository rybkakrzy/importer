using D2ViewerEditor.Application.Common.Security;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// DEV-only identity provider: reads CorporateKey / admin flag from request headers.
/// Kept for local testing without an identity provider. The production setup registers
/// <c>ClaimsCurrentUserProvider</c> (Entra ID access token) instead — see Program.cs.
/// </summary>
public sealed class HttpHeaderCurrentUserProvider : ICurrentUserProvider
{
    public const string HeaderName = "X-Corporate-Key";
    public const string AdminHeaderName = "X-App-Admin";

    /// <summary>Fallback CorporateKey for local dev when no <see cref="HeaderName"/> is supplied, so the
    /// editor identity (last-modified-by / delivery sender) is never null in the dev bypass flow.</summary>
    public const string DefaultDevCorporateKey = "DEV-LOCAL";

    private readonly IHttpContextAccessor _httpContextAccessor;

    public HttpHeaderCurrentUserProvider(IHttpContextAccessor httpContextAccessor)
    {
        _httpContextAccessor = httpContextAccessor;
    }

    public string? CorporateKey
    {
        get
        {
            var value = _httpContextAccessor.HttpContext?.Request.Headers[HeaderName].ToString();
            return string.IsNullOrWhiteSpace(value) ? DefaultDevCorporateKey : value.Trim();
        }
    }

    public bool IsAdmin =>
        string.Equals(
            _httpContextAccessor.HttpContext?.Request.Headers[AdminHeaderName].ToString()?.Trim(),
            "true",
            StringComparison.OrdinalIgnoreCase);
}

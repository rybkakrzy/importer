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
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }

    public bool IsAdmin =>
        string.Equals(
            _httpContextAccessor.HttpContext?.Request.Headers[AdminHeaderName].ToString()?.Trim(),
            "true",
            StringComparison.OrdinalIgnoreCase);
}

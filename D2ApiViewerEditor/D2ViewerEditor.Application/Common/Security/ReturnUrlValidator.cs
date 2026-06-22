using System.Net;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Application.Common.Security;

public sealed class ReturnUrlValidator : IReturnUrlValidator
{
    private readonly ReturnUrlSecurityOptions _options;

    public ReturnUrlValidator(IOptions<ReturnUrlSecurityOptions> options)
    {
        _options = options.Value;
    }

    public ReturnUrlValidationResult Validate(string? rawUrl)
    {
        if (string.IsNullOrWhiteSpace(rawUrl))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.Empty, "Callback URL jest wymagany.");

        var candidate = rawUrl.Trim();
        if (candidate.Length > _options.MaxLength)
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.TooLong,
                $"Callback URL nie może być dłuższy niż {_options.MaxLength} znaków.");

        if (ContainsControlCharacters(candidate))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.ContainsControlCharacters,
                "Callback URL zawiera niedozwolone znaki kontrolne.");

        var decoded = DecodeForInspection(candidate);
        if (ContainsDangerousCharacters(decoded))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.DangerousCharacters,
                "Callback URL zawiera niedozwolone znaki.");

        if (decoded.StartsWith("//", StringComparison.Ordinal))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.ProtocolRelative,
                "Callback URL nie może być adresem protocol-relative.");

        if (!Uri.TryCreate(candidate, UriKind.Absolute, out var uri))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.InvalidAbsoluteUri,
                "Callback URL musi być absolutnym adresem http(s).");

        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.InvalidScheme,
                "Callback URL musi być absolutnym adresem http(s).");

        if (_options.RequireHttps && uri.Scheme != Uri.UriSchemeHttps)
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.InsecureScheme,
                "Callback URL musi używać protokołu https.");

        if (!string.IsNullOrEmpty(uri.UserInfo))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.UserInfoNotAllowed,
                "Callback URL nie może zawierać danych uwierzytelniających.");

        if (!_options.AllowLoopback && (uri.IsLoopback || IsLoopbackHost(uri.Host)))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.LoopbackNotAllowed,
                "Callback URL nie może wskazywać na host lokalny.");

        if (!_options.AllowPrivateNetworkIp && IPAddress.TryParse(uri.Host, out var ip) && IsPrivateOrSpecialIp(ip))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.PrivateNetworkIpNotAllowed,
                "Callback URL nie może wskazywać prywatnego adresu IP.");

        if (_options.AllowedHosts.Length > 0 && !IsAllowedHost(uri.Host, _options.AllowedHosts))
            return ReturnUrlValidationResult.Failure(ReturnUrlRejectionCode.HostNotAllowListed,
                "Callback URL wskazuje host spoza listy dozwolonych odbiorców.");

        var normalized = uri.GetComponents(UriComponents.HttpRequestUrl, UriFormat.UriEscaped);
        return ReturnUrlValidationResult.Success(normalized);
    }

    private static bool ContainsControlCharacters(string value)
    {
        for (var i = 0; i < value.Length; i++)
        {
            if (char.IsControl(value[i])) return true;
        }

        return false;
    }

    private static string DecodeForInspection(string value)
    {
        var current = value;
        for (var i = 0; i < 3; i++)
        {
            var decoded = Uri.UnescapeDataString(current);
            if (string.Equals(decoded, current, StringComparison.Ordinal))
                break;

            current = decoded;
        }

        return current;
    }

    private static bool ContainsDangerousCharacters(string value) =>
        value.Contains('<') || value.Contains('>') || value.Contains('"') || value.Contains('\'') || value.Contains('\\');

    private static bool IsAllowedHost(string host, IEnumerable<string> allowList)
    {
        foreach (var allowed in allowList)
        {
            var normalizedAllowed = allowed.Trim().TrimStart('.');
            if (normalizedAllowed.Length == 0) continue;

            if (host.Equals(normalizedAllowed, StringComparison.OrdinalIgnoreCase)
                || host.EndsWith($".{normalizedAllowed}", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsLoopbackHost(string host) =>
        host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
        || host.Equals("localhost.localdomain", StringComparison.OrdinalIgnoreCase);

    private static bool IsPrivateOrSpecialIp(IPAddress ip)
    {
        if (IPAddress.IsLoopback(ip)) return true;

        if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
        {
            return ip.IsIPv6LinkLocal || ip.IsIPv6SiteLocal || ip.IsIPv6Multicast || ip.IsIPv6Teredo;
        }

        var b = ip.GetAddressBytes();
        if (b.Length != 4) return false;

        return b[0] == 10
               || (b[0] == 172 && b[1] >= 16 && b[1] <= 31)
               || (b[0] == 192 && b[1] == 168)
               || (b[0] == 169 && b[1] == 254)
               || b[0] == 127;
    }
}

using System.Text.Json;
using D2ViewerEditor.Domain.Security;

namespace D2ViewerEditor.Application.Common.Security;

public sealed class DocumentAccessGuard : IDocumentAccessGuard
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private readonly ICurrentUserProvider _currentUser;

    public DocumentAccessGuard(ICurrentUserProvider currentUser)
    {
        _currentUser = currentUser;
    }

    public bool IsViewAllowed(string? metadataJson)
    {
        // APP_Admin has full document oversight — bypasses the per-document allow-list
        // (business decision: admin role grants content oversight). Audited via logs upstream.
        if (_currentUser.IsAdmin)
            return true;

        var allowedKeys = ParseAllowedKeys(metadataJson);
        return DocumentAccessPolicy.IsViewAllowed(allowedKeys, _currentUser.CorporateKey);
    }

    private static IReadOnlyCollection<string>? ParseAllowedKeys(string? metadataJson)
    {
        if (string.IsNullOrWhiteSpace(metadataJson))
            return null;

        try
        {
            var parsed = JsonSerializer.Deserialize<AccessMetadata>(metadataJson, JsonOptions);
            return parsed?.AllowedCorporateKeys;
        }
        catch (JsonException)
        {
            // Malformed metadata must neither crash the request nor silently restrict access;
            // treated as "no allow-list" → link-based public access (same as the empty-list rule).
            return null;
        }
    }

    private sealed record AccessMetadata(string[]? AllowedCorporateKeys);
}

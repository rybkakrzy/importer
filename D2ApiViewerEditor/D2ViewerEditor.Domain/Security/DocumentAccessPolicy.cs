namespace D2ViewerEditor.Domain.Security;

/// <summary>
/// Pure view-access rule for a document, driven by the optional allow-list of CorporateKeys
/// carried in document metadata. No dependency on identity provider — the caller supplies
/// the already-resolved user CorporateKey, which keeps this rule reusable for the future
/// Entra ID phase (token → CorporateKey mapping happens upstream).
/// </summary>
public static class DocumentAccessPolicy
{
    public static bool IsViewAllowed(IReadOnlyCollection<string>? allowedCorporateKeys, string? userCorporateKey)
    {
        var allowList = Normalize(allowedCorporateKeys);

        // No list / null / empty / blank-only → public for anyone holding the link.
        if (allowList.Count == 0)
            return true;

        var key = userCorporateKey?.Trim();
        if (string.IsNullOrEmpty(key))
            return false;

        // Case-insensitive Ordinal: CorporateKeys are corporate identifiers, not secrets;
        // casing/whitespace differences must not cause false denials.
        return allowList.Contains(key, StringComparer.OrdinalIgnoreCase);
    }

    private static List<string> Normalize(IReadOnlyCollection<string>? keys)
    {
        if (keys is null)
            return [];

        return keys
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Select(k => k.Trim())
            .ToList();
    }
}

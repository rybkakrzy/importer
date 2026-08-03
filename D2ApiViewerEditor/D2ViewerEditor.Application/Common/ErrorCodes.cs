namespace D2ViewerEditor.Application.Common;

/// <summary>
/// Stabilne kody maszynowe błędów domenowych. Backend = źródło prawdy: GUI rozpoznaje przypadek
/// po <c>code</c> (nie po treści komunikatu), więc kody muszą być NIEZMIENNE i niezależne od
/// lokalizacji tekstu. Kod jest wystawiany zarówno w odpowiedziach <c>{ code, error }</c>
/// (np. POST /open), jak i w rozszerzeniach ProblemDetails walidacji (patrz middleware).
/// </summary>
public static class ErrorCodes
{
    /// <summary>Dokument nie zawiera treści (pusty plik / zawartość zerowej długości).</summary>
    public const string DocumentContentEmpty = "DOCUMENT_CONTENT_EMPTY";

    /// <summary>Dokument jest zaszyfrowany — GUI musi poprosić użytkownika o hasło i ponowić otwarcie.</summary>
    public const string DocumentProtected = "PASSWORD_REQUIRED";

    /// <summary>Hasło podane przy otwieraniu nie odszyfrowało dokumentu.</summary>
    public const string DocumentUnlockFailed = "WRONG_PASSWORD";

    /// <summary>Plik jest binarnym dokumentem .doc (starszy format Worda) — wymaga konwersji do .docx.</summary>
    public const string UnsupportedLegacyDoc = "UNSUPPORTED_LEGACY_DOC";

    private static readonly HashSet<string> Known = new(StringComparer.Ordinal)
    {
        DocumentContentEmpty,
        DocumentProtected,
        DocumentUnlockFailed,
        UnsupportedLegacyDoc
    };

    /// <summary>Czy podany kod jest znanym, stabilnym kodem domenowym (do wystawienia klientowi).</summary>
    public static bool IsKnown(string? code) => code is not null && Known.Contains(code);
}

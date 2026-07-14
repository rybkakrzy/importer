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

    private static readonly HashSet<string> Known = new(StringComparer.Ordinal)
    {
        DocumentContentEmpty
    };

    /// <summary>Czy podany kod jest znanym, stabilnym kodem domenowym (do wystawienia klientowi).</summary>
    public static bool IsKnown(string? code) => code is not null && Known.Contains(code);
}

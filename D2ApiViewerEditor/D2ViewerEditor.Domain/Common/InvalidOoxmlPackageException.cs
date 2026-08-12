namespace D2ViewerEditor.Domain.Common;

/// <summary>
/// Pakiet DOCX nie nadaje się do analizy struktury: uszkodzone archiwum, brak wymaganych części
/// OOXML, niepoprawny XML albo przekroczony limit bezpieczeństwa. Wejście jest niezaufane, więc
/// każdy z tych przypadków jest kontrolowanym błędem wejścia, a nie awarią serwera.
/// </summary>
public sealed class InvalidOoxmlPackageException : Exception
{
    public InvalidOoxmlPackageException(string message)
        : base(message)
    {
    }
}

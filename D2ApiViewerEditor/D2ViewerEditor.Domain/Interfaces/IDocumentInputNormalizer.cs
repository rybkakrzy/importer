namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Wynik normalizacji wejściowego pliku do zwykłego pakietu DOCX (OOXML/ZIP), gotowego dla parsera.
/// </summary>
public enum DocumentInputStatus
{
    /// <summary>Plik jest (lub został doprowadzony do) zwykłym DOCX — <see cref="DocumentInputResult.Docx"/> ustawione.</summary>
    Ok = 0,

    /// <summary>Plik jest zaszyfrowanym DOCX, a hasło nie zostało podane.</summary>
    PasswordRequired = 1,

    /// <summary>Plik jest zaszyfrowanym DOCX, ale podane hasło jest nieprawidłowe.</summary>
    WrongPassword = 2,

    /// <summary>Binarny .doc (starszy Word) — brak wbudowanego konwertera (wymaga zewnętrznej usługi).</summary>
    UnsupportedLegacyDoc = 3,

    /// <summary>Nie rozpoznano formatu (nie DOCX/CFB) lub plik uszkodzony.</summary>
    Invalid = 4
}

/// <summary>Wynik <see cref="IDocumentInputNormalizer.Normalize"/>.</summary>
public sealed record DocumentInputResult(DocumentInputStatus Status, byte[]? Docx)
{
    public static DocumentInputResult Success(byte[] docx) => new(DocumentInputStatus.Ok, docx);
    public static DocumentInputResult Failure(DocumentInputStatus status) => new(status, null);
}

/// <summary>
/// Doprowadza wejściowy plik do zwykłego DOCX zanim trafi do parsera DOCX→HTML:
/// rozpoznaje format po magic-bytes, rozszyfrowuje DOCX zabezpieczony hasłem i wykrywa binarny .doc.
/// Czysto zarządzane (Linux/GCP-safe) — bez LibreOffice/System.Drawing.
/// </summary>
public interface IDocumentInputNormalizer
{
    /// <param name="bytes">Surowe bajty przesłanego pliku.</param>
    /// <param name="password">Hasło do odszyfrowania (jeśli plik jest zaszyfrowany); null/puste = brak.</param>
    DocumentInputResult Normalize(byte[] bytes, string? password = null);
}

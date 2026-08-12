using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Diagnostyka struktury pakietu DOCX na potrzeby narzędzia „Walidator struktury".
///
/// Dwie niezależne warstwy:
///  - surowy XML wpisów ZIP (źródło prawdy — widać także konstrukcje, których nasz edytor
///    ani Open XML SDK nie interpretują),
///  - walidacja schematu przez <c>OpenXmlValidator</c> (warstwa interpretacji SDK).
///
/// Brak błędów walidatora NIE oznacza, że dokument jest poprawnie obsługiwany przez edytor —
/// dlatego równolegle działają własne reguły diagnostyczne i profil zgodności edytora.
/// </summary>
public interface IDocumentStructureInspector
{
    /// <summary>Analizuje pakiet. Rzuca <see cref="Common.InvalidOoxmlPackageException"/> dla wejścia nienadającego się do analizy.</summary>
    DocumentStructureAnalysis Analyze(byte[] documentBytes, string fileName, CancellationToken cancellationToken);

    /// <summary>Surowy XML pojedynczego elementu wraz z numerem linii w części źródłowej.</summary>
    OoxmlFragment? ReadElementXml(byte[] documentBytes, string partPath, IReadOnlyList<int> nodePath, CancellationToken cancellationToken);

    /// <summary>Numer linii, od której zaczyna się element w surowym XML części.</summary>
    int? FindElementLine(byte[] documentBytes, string partPath, IReadOnlyList<int> nodePath, CancellationToken cancellationToken);

    /// <summary>Surowy XML całej części pakietu, dokładnie taki jak we wpisie ZIP, albo <c>null</c> gdy części nie ma.</summary>
    string? ReadPartXml(byte[] documentBytes, string partPath, CancellationToken cancellationToken);

    /// <summary>Wyniki walidacji schematu dla wskazanego profilu Office (bez ponownego indeksowania drzewa).</summary>
    IReadOnlyList<SchemaValidationIssue> ValidateSchema(
        byte[] documentBytes,
        string targetVersion,
        IReadOnlyList<InspectedElement> elements,
        CancellationToken cancellationToken);

    /// <summary>Nazwy profili Open XML, które można wskazać przy walidacji schematu.</summary>
    IReadOnlyList<string> GetSupportedSchemaTargets();
}

/// <summary>Fragment surowego XML wraz z lokalizacją w części źródłowej.</summary>
public sealed record OoxmlFragment(string Xml, int? SourceLine);

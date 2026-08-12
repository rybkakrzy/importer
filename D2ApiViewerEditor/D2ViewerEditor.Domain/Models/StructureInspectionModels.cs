namespace D2ViewerEditor.Domain.Models;

/// <summary>
/// Waga problemu diagnostycznego wykrytego w strukturze OOXML.
/// </summary>
public enum StructureIssueSeverity
{
    Info = 0,
    Warning = 1,
    Error = 2
}

/// <summary>
/// Pojedynczy problem przypisany do elementu XML albo do całego pakietu. <c>Code</c> jest stabilnym
/// identyfikatorem maszynowym reguły — GUI rozpoznaje przypadek po kodzie, nie po treści opisu.
/// </summary>
public sealed record StructureIssue(
    string Code,
    StructureIssueSeverity Severity,
    string Title,
    string Description);

/// <summary>
/// Atrybut XML elementu. <c>InterpretedValue</c> to opcjonalne odczytanie wartości surowej
/// (np. <c>behindDoc="1"</c> → <c>true</c>); brak interpretacji nie ukrywa wartości surowej.
/// </summary>
public sealed record StructureAttribute(
    string Name,
    string LocalName,
    string NamespaceUri,
    string RawValue,
    string? InterpretedValue);

/// <summary>
/// Właściwość wyliczona przez warstwę semantyczną. <c>Source</c> mówi, SKĄD wzięła się wartość
/// (docDefaults / styl / formatowanie bezpośrednie / motyw), <c>SourceReference</c> uszczegóławia
/// źródło (np. identyfikator stylu), a <c>IsRedundant</c> oznacza wartość bezpośrednią identyczną
/// z odziedziczoną.
/// </summary>
public sealed record StructureProperty(
    string Name,
    string? Value,
    string Source,
    string? SourceReference = null,
    bool IsRedundant = false);

/// <summary>
/// Stan relationshipu: wskazuje istniejącą część pakietu, zasób zewnętrzny, cel nieistniejący,
/// albo nie jest w ogóle zadeklarowany w części źródłowej.
/// </summary>
public enum StructureRelationshipStatus
{
    Resolved = 0,
    External = 1,
    TargetMissing = 2,
    NotDeclared = 3
}

/// <summary>
/// Relationship (<c>rId</c>) rozwiązany względem części źródłowej. <c>rId</c> jest unikalny tylko
/// w obrębie części, dlatego tożsamość niesie para (część źródłowa, identyfikator).
/// </summary>
public sealed record StructureRelationship(
    string SourcePart,
    string RelationshipPartPath,
    string Id,
    string Type,
    string Target,
    string TargetMode,
    string? ResolvedTarget,
    StructureRelationshipStatus Status);

/// <summary>
/// Poziom obsługi cechy OOXML przez nasz edytor, wg konfigurowalnego profilu. Profil jest DANYMI
/// konfiguracji — walidator niczego nie zgaduje.
/// </summary>
public sealed record EditorCompatibilityInfo(
    string Feature,
    string Level,
    string? Notes);

/// <summary>
/// Błąd zwrócony przez <c>OpenXmlValidator</c> (warstwa schematu Open XML SDK), w miarę możliwości
/// zmapowany na konkretny element drzewa (<see cref="ElementId"/>).
/// </summary>
public sealed record SchemaValidationIssue(
    string Code,
    string Severity,
    string Description,
    string? PartPath,
    string? NodeName,
    string? Path,
    string? ElementId,
    string TargetVersion);

/// <summary>
/// Wpis pakietu (także binarny, np. obraz) — potrzebny do sprawdzania celów relationshipów
/// i osiągalności części.
/// </summary>
public sealed record InspectedPackageEntry(
    string Path,
    long UncompressedSize,
    long CompressedSize,
    string? ContentType,
    bool IsXml);

/// <summary>
/// Część XML pakietu uwzględniona w indeksowaniu.
/// </summary>
public sealed record InspectedPackagePart(
    string Path,
    string? ContentType,
    long UncompressedSize,
    long CompressedSize,
    int ElementCount,
    bool IsIndexed);

/// <summary>
/// Wiązanie nagłówka/stopki z sekcją. <c>Source</c> rozróżnia <c>Direct</c> (referencja w tej
/// sekcji), <c>Inherited</c> (z poprzedniej sekcji) i <c>Missing</c> (brak wiązania).
/// </summary>
public sealed record HeaderFooterBinding(
    string Kind,
    string Type,
    string Source,
    int? SourceSectionNumber,
    bool IsActive,
    string? ReferenceElementId,
    string? RelationshipId,
    string? RelationshipType,
    string? TargetMode,
    string? Target,
    string? PartPath,
    string? PartRootElementId,
    bool PartExists,
    IReadOnlyList<StructureIssue> Issues);

/// <summary>
/// Sekcja dokumentu wraz z efektywnymi wiązaniami nagłówków i stopek.
/// </summary>
public sealed record DocumentSectionInfo(
    int Number,
    string SectionPropertiesElementId,
    string DisplayPath,
    bool FirstPageDifferent,
    bool EvenAndOddHeaders,
    IReadOnlyList<HeaderFooterBinding> HeaderFooterBindings,
    IReadOnlyList<StructureIssue> Issues);

/// <summary>
/// Element XML pakietu wraz z diagnostyką. <see cref="NodePath"/> to indeksy dzieci od korzenia
/// części — stabilna, niezależna od prefiksów namespace lokalizacja pozwalająca odnaleźć ten sam
/// węzeł w surowym XML przy leniwym pobieraniu fragmentu.
/// </summary>
public sealed class InspectedElement
{
    public required string Id { get; init; }
    public string? ParentId { get; init; }
    public required string PartPath { get; init; }
    public required int Depth { get; init; }
    public required int Order { get; init; }
    public required IReadOnlyList<int> NodePath { get; init; }
    public required string DisplayPath { get; init; }
    public required string XmlName { get; init; }
    public required string LocalName { get; init; }
    public required string NamespaceUri { get; init; }
    public required string Category { get; init; }
    public required string DisplayName { get; init; }
    public string? Preview { get; init; }
    public required bool HasChildren { get; init; }
    public List<StructureAttribute> Attributes { get; } = [];
    public List<StructureProperty> Properties { get; } = [];
    public List<StructureRelationship> Relationships { get; } = [];
    public List<StructureIssue> Issues { get; } = [];
    public List<EditorCompatibilityInfo> EditorCompatibility { get; } = [];

    /// <summary>
    /// Gotowy indeks wyszukiwania (małymi literami), wypełniany PO zakończeniu wszystkich warstw
    /// analizy — obejmuje atrybuty, właściwości, relationshipy i diagnostykę, więc filtr GUI nie
    /// musi dociągać szczegółów elementów ani niczego normalizować per klawisz.
    /// </summary>
    public string SearchText { get; set; } = string.Empty;

    public StructureIssueSeverity? HighestSeverity =>
        Issues.Count == 0 ? null : Issues.Max(issue => issue.Severity);
}

/// <summary>
/// Wynik analizy struktury dokumentu: spłaszczone drzewo elementów wszystkich części XML, lista
/// części i wpisów pakietu, sekcje z wiązaniami nagłówków/stopek, diagnostyka OPC oraz wyniki
/// walidacji schematu Open XML.
/// </summary>
public sealed class DocumentStructureAnalysis
{
    public required string FileName { get; init; }
    public required long FileSizeInBytes { get; init; }

    /// <summary>Główna część dokumentu odnaleziona przez relationship, a nie po nazwie pliku.</summary>
    public required string MainDocumentPartPath { get; init; }

    public required IReadOnlyList<InspectedElement> Elements { get; init; }
    public required IReadOnlyDictionary<string, InspectedElement> ElementsById { get; init; }
    public required IReadOnlyList<InspectedPackagePart> Parts { get; init; }
    public required IReadOnlyList<InspectedPackageEntry> Entries { get; init; }
    public required IReadOnlyList<DocumentSectionInfo> Sections { get; init; }
    public required IReadOnlyList<StructureIssue> PackageIssues { get; init; }
    public required IReadOnlyList<SchemaValidationIssue> SchemaIssues { get; init; }

    /// <summary>Liczba błędów schematu przed przycięciem listy do limitu prezentacyjnego.</summary>
    public required int SchemaIssueCount { get; init; }

    /// <summary>Indeksowanie zatrzymało się na limicie elementów — drzewo jest niepełne.</summary>
    public required bool ElementsTruncated { get; init; }
}

/// <summary>
/// Zapamiętana analiza wraz z bajtami dokumentu — surowy XML elementu i części jest wyciągany
/// leniwie z tych bajtów, dzięki czemu ani odpowiedź na wczytanie, ani magazyn nie trzymają
/// rozpakowanego XML całego pakietu.
/// </summary>
public sealed record DocumentStructureInspection(
    Guid Id,
    DateTimeOffset CreatedAtUtc,
    DateTimeOffset ExpiresAtUtc,
    DocumentStructureAnalysis Analysis,
    byte[] DocumentBytes);

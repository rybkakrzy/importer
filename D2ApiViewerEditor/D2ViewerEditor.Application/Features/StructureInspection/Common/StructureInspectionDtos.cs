using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Application.Features.StructureInspection.Common;

/// <summary>
/// Podsumowanie wczytanego dokumentu. Nie niesie XML — surowy XML elementu i części jest pobierany
/// dopiero na żądanie, żeby wczytanie dużego dokumentu nie przesyłało megabajtów tekstu.
/// </summary>
public record StructureInspectionSummaryDto(
    Guid InspectionId,
    string FileName,
    long FileSizeInBytes,
    string MainDocumentPartPath,
    DateTimeOffset ExpiresAtUtc,
    int ElementCount,
    int ErrorCount,
    int WarningCount,
    int InfoCount,
    int SchemaIssueCount,
    int PackageIssueCount,
    int SectionCount,
    bool ElementsTruncated,
    IReadOnlyList<StructurePartDto> Parts,
    IReadOnlyList<string> Categories);

/// <summary>Część pakietu widoczna w filtrach GUI.</summary>
public record StructurePartDto(
    string Path,
    string? ContentType,
    long UncompressedSize,
    long CompressedSize,
    int ElementCount);

/// <summary>
/// Wiersz drzewa elementów — celowo wąski, bo lista może mieć dziesiątki tysięcy pozycji.
/// <c>SearchText</c> jest gotowym indeksem wyszukiwania, żeby GUI nie musiało pobierać szczegółów
/// wszystkich elementów do filtrowania.
/// </summary>
public record StructureElementDto(
    string Id,
    string? ParentId,
    int Depth,
    string PartPath,
    string XmlName,
    string Category,
    string DisplayName,
    string? Preview,
    string SearchText,
    string Severity,
    int IssueCount,
    bool HasChildren);

/// <summary>Pełne szczegóły jednego elementu, pobierane po kliknięciu.</summary>
public record StructureElementDetailsDto(
    string Id,
    string? ParentId,
    int Depth,
    string PartPath,
    string DisplayPath,
    string XmlName,
    string LocalName,
    string NamespaceUri,
    string Category,
    string DisplayName,
    string? Preview,
    IReadOnlyList<StructureAttribute> Attributes,
    IReadOnlyList<StructureProperty> Properties,
    IReadOnlyList<StructureRelationship> Relationships,
    IReadOnlyList<StructureIssue> Issues,
    IReadOnlyList<EditorCompatibilityInfo> EditorCompatibility);

/// <summary>Surowy XML wybranego elementu wraz z jego lokalizacją w pakiecie i numerem linii.</summary>
public record StructureElementXmlDto(
    string ElementId,
    string PartPath,
    string DisplayPath,
    string Xml,
    int? SourceLine);

/// <summary>Surowy XML całej części pakietu wraz z linią zaznaczonego elementu (jeśli wskazano).</summary>
public record StructurePartXmlDto(
    string PartPath,
    string Xml,
    string? HighlightElementId,
    int? HighlightLine);

/// <summary>Błąd walidatora schematu Open XML, w miarę możliwości powiązany z elementem drzewa.</summary>
public record SchemaIssueDto(
    string Code,
    string Severity,
    string Description,
    string? PartPath,
    string? NodeName,
    string? Path,
    string? ElementId,
    string TargetVersion);

/// <summary>Wyniki walidacji schematu dla wybranego profilu Office.</summary>
public record SchemaIssuesDto(
    string TargetVersion,
    int TotalCount,
    IReadOnlyList<SchemaIssueDto> Issues);

/// <summary>Diagnostyka pakietu OPC: główna część, problemy i lista profili walidacji.</summary>
public record PackageDiagnosticsDto(
    string MainDocumentPartPath,
    IReadOnlyList<StructureIssue> Issues,
    IReadOnlyList<StructurePackageEntryDto> Entries,
    IReadOnlyList<string> SupportedSchemaTargets);

/// <summary>Wpis pakietu (także binarny) wraz z rozstrzygniętym typem zawartości.</summary>
public record StructurePackageEntryDto(
    string Path,
    long UncompressedSize,
    long CompressedSize,
    string? ContentType,
    bool IsXml);

/// <summary>Sekcja dokumentu z efektywnymi wiązaniami nagłówków i stopek.</summary>
public record DocumentSectionDto(
    int Number,
    string SectionPropertiesElementId,
    string DisplayPath,
    bool FirstPageDifferent,
    bool EvenAndOddHeaders,
    IReadOnlyList<HeaderFooterBinding> HeaderFooterBindings,
    IReadOnlyList<StructureIssue> Issues);

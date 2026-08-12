namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>
/// Limity bezpieczeństwa i wydajności analizy struktury DOCX. Wejście jest niezaufane (ZIP + XML),
/// więc każdy wymiar pakietu ma twardy limit: rozmiar, liczba wpisów, rozmiar po dekompresji,
/// współczynnik kompresji, wielkość pojedynczego XML, liczba elementów i głębokość zagnieżdżenia.
/// </summary>
public sealed class StructureInspectionOptions
{
    public const string SectionName = "StructureInspection";

    public long MaxUploadBytes { get; set; } = 25L * 1024 * 1024;
    public int MaxZipEntries { get; set; } = 2_000;
    public long MaxSingleEntryBytes { get; set; } = 20L * 1024 * 1024;
    public long MaxTotalUncompressedBytes { get; set; } = 100L * 1024 * 1024;
    public double MaxCompressionRatio { get; set; } = 250d;
    public long MaxXmlCharacters { get; set; } = 30_000_000;
    public int MaxElements { get; set; } = 50_000;
    public int MaxXmlDepth { get; set; } = 160;

    /// <summary>Ile błędów walidatora schematu trafia do odpowiedzi (licznik całkowity zostaje pełny).</summary>
    public int MaxSchemaIssues { get; set; } = 500;

    /// <summary>Ile analiz naraz trzyma magazyn w pamięci (najstarsze są usuwane).</summary>
    public int MaxStoredInspections { get; set; } = 4;

    /// <summary>
    /// Sufit ŁĄCZNEGO rozmiaru bajtów pakietów w magazynie — uczciwszy niż sam licznik wpisów,
    /// bo cztery 25-megabajtowe analizy ważą co innego niż cztery 200-kilobajtowe.
    /// </summary>
    public long MaxStoredBytes { get; set; } = 120L * 1024 * 1024;

    public TimeSpan InspectionTimeToLive { get; set; } = TimeSpan.FromMinutes(30);

    /// <summary>Domyślny profil Open XML używany przy analizie; GUI może poprosić o inny.</summary>
    public string DefaultSchemaTarget { get; set; } = "Microsoft365";
}

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using D2ViewerEditor.Domain.Entities;

namespace D2ViewerEditor.Benchmarks.Benchmarks;

/// <summary>
/// Benchmarki logiki domenowej agregatu Document.
///
/// Mierzone obszary:
///  - AddVersion: O(N) pętla Deactivate() po wszystkich poprzednich wersjach
///  - RestoreVersion: O(N) pętla Deactivate() + O(1) wyszukanie FirstOrDefault
///  - GetActiveVersion: O(N) FirstOrDefault po liście wersji (zależy od pozycji aktywnej)
///
/// Parametr VersionCount symuluje dokument z różną historią edycji.
/// Warto obserwować jak rośnie czas przy rosnącej liczbie wersji —
/// idealnie liniowo, ale GC pressure z dużymi listami może zaburzyć wyniki.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class DocumentDomainBenchmarks
{
    private static readonly byte[] SmallContent = new byte[512];

    /// <summary>
    /// Liczba istniejących wersji dokumentu przed wykonaniem operacji.
    /// Imituje: 1 = nowy dokument, 10 = aktywny edytowany, 50 = długa historia.
    /// </summary>
    [Params(1, 10, 50)]
    public int VersionCount { get; set; }

    private Document _doc = null!;
    private Guid _firstVersionId;
    private Guid _lastVersionId;

    static DocumentDomainBenchmarks()
    {
        Random.Shared.NextBytes(SmallContent);
    }

    /// <summary>
    /// Przygotowuje dokument z VersionCount wersjami przed każdą iteracją.
    /// IterationSetup NIE jest wliczany do mierzonego czasu.
    /// </summary>
    [IterationSetup]
    public void IterationSetup()
    {
        _doc = new Document(Guid.NewGuid(), "benchmark.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "benchmark-user");

        for (int i = 0; i < VersionCount; i++)
        {
            var v = _doc.AddVersion(SmallContent, "benchmark-user");
            if (i == 0) _firstVersionId = v.Id;
            _lastVersionId = v.Id;
        }
    }

    /// <summary>
    /// Dodaje kolejną wersję do dokumentu z N istniejącymi wersjami.
    /// Koszt: O(N) wywołań Deactivate() na poprzednich wersjach.
    /// </summary>
    [Benchmark(Baseline = true)]
    public DocumentVersion AddVersion()
        => _doc.AddVersion(SmallContent, "benchmark-user");

    /// <summary>
    /// Przywraca pierwszą wersję (najdalszą w historii) z N wersji.
    /// Koszt: O(N) pętla Deactivate + O(N/2) FirstOrDefault (pierwsza wersja jest na początku listy).
    /// </summary>
    [Benchmark]
    public void RestoreFirstVersion()
        => _doc.RestoreVersion(_firstVersionId);

    /// <summary>
    /// Przywraca najnowszą wersję — jest już aktywna, ale pętla Deactivate nadal przebiega O(N).
    /// Dobry baseline porównawczy dla RestoreFirstVersion.
    /// </summary>
    [Benchmark]
    public void RestoreLastVersion()
        => _doc.RestoreVersion(_lastVersionId);

    /// <summary>
    /// Pobiera aktywną wersję dokumentu.
    /// Aktywna wersja to zawsze ostatnia — FirstOrDefault zatrzymuje się po N iteracjach (koszt O(N)).
    /// Uwaga: po IterationSetup aktywna jest ostatnia wersja, więc FirstOrDefault skanuje całą listę.
    /// </summary>
    [Benchmark]
    public DocumentVersion? GetActiveVersion()
        => _doc.GetActiveVersion();
}

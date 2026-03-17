using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks.Benchmarks;

/// <summary>
/// Benchmarki konwersji DOCX → HTML (DocxToHtmlConverter).
///
/// Mierzone obszary:
///  - Otwieranie paczki OPC (WordprocessingDocument.Open)
///  - Parsowanie stylów dokumentu (StyleDefinitionsPart)
///  - Parsowanie metadanych (CoreFilePropertiesPart, ExtendedFilePropertiesPart)
///  - Iteracja po akapitach ciała (Body paragraphs + runs)
///  - Ekstrakcja nagłówka/stopki (HeaderPart, FooterPart)
///  - Ekstrakcja i kodowanie obrazów do base64 (ImageParts → DocumentImage list)
///  - Parsowanie definicji numerowania (NumberingDefinitionsPart)
///
/// Pliki DOCX są generowane w GlobalSetup przez HtmlToDocxConverter,
/// więc benchmarki mierzą rzeczywistą strukturę produkowaną przez system.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class DocxToHtmlBenchmarks
{
    private byte[] _smallDocx = null!;
    private byte[] _mediumDocx = null!;
    private byte[] _largeDocx = null!;
    private byte[] _docxWithTable = null!;

    [GlobalSetup]
    public void Setup()
    {
        // Generujemy DOCX przez produkcyjny HtmlToDocxConverter,
        // żeby benchmarkować realny format wyjściowy systemu.
        var gen = new HtmlToDocxConverter();

        _smallDocx = gen.Convert(BuildParagraphHtml(2, withFormatting: false));
        _mediumDocx = gen.Convert(BuildParagraphHtml(15, withFormatting: true));
        _largeDocx = gen.Convert(BuildParagraphHtml(60, withFormatting: true));
        _docxWithTable = gen.Convert(BuildHtmlWithTable(rows: 20, cols: 5));
    }

    /// <summary>
    /// Mały dokument — 2 akapity.
    /// Baseline: mierzy stały koszt otwarcia paczki OPC + parsowanie stylów/metadanych.
    /// </summary>
    [Benchmark(Baseline = true)]
    public object SmallDocument()
    {
        using var stream = new MemoryStream(_smallDocx);
        return new DocxToHtmlConverter().Convert(stream);
    }

    /// <summary>
    /// Średni dokument — 15 akapitów z formatowaniem.
    /// Typowy przypadek użycia: otwarcie dokumentu roboczego z edytora.
    /// </summary>
    [Benchmark]
    public object MediumDocument()
    {
        using var stream = new MemoryStream(_mediumDocx);
        return new DocxToHtmlConverter().Convert(stream);
    }

    /// <summary>
    /// Duży dokument — 60 akapitów. Mierzy skalowalność parsowania ciała dokumentu.
    /// Oczekiwany wzrost: liniowy względem liczby akapitów.
    /// </summary>
    [Benchmark]
    public object LargeDocument()
    {
        using var stream = new MemoryStream(_largeDocx);
        return new DocxToHtmlConverter().Convert(stream);
    }

    /// <summary>
    /// Dokument z tabelą 20×5. Mierzy koszt parsowania Word Table → HTML table.
    /// Tabele to zagnieżdżone elementy XML — bardziej kosztowne w traversal.
    /// </summary>
    [Benchmark]
    public object DocumentWithTable()
    {
        using var stream = new MemoryStream(_docxWithTable);
        return new DocxToHtmlConverter().Convert(stream);
    }

    // --- Generatory identyczne jak w HtmlToDocxBenchmarks ---

    private static string BuildParagraphHtml(int paragraphCount, bool withFormatting)
    {
        var sb = new StringBuilder(paragraphCount * 120);
        for (int i = 1; i <= paragraphCount; i++)
        {
            if (withFormatting)
                sb.Append($"<p>Akapit {i}: Treść dokumentu zawierająca <b>pogrubiony tekst</b>, <i>kursywę</i> oraz <u>podkreślenie</u>. Numer referencyjny: REF-{i:D4}.</p>");
            else
                sb.Append($"<p>Akapit {i}: Zwykły tekst dokumentu. Numer referencyjny: REF-{i:D4}.</p>");
        }
        return sb.ToString();
    }

    private static string BuildHtmlWithTable(int rows, int cols)
    {
        var sb = new StringBuilder();
        sb.Append("<table border=\"1\"><thead><tr>");
        for (int c = 1; c <= cols; c++)
            sb.Append($"<th>Kolumna {c}</th>");
        sb.Append("</tr></thead><tbody>");
        for (int r = 1; r <= rows; r++)
        {
            sb.Append("<tr>");
            for (int c = 1; c <= cols; c++)
                sb.Append($"<td>Wiersz {r}, Kol {c}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }
}

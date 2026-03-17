using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks.Benchmarks;

/// <summary>
/// Benchmarki konwersji HTML → DOCX (HtmlToDocxConverter).
///
/// Mierzone obszary:
///  - Parsowanie HTML przez HtmlAgilityPack
///  - Budowanie struktury OpenXML (paragrafy, runs, tabele, nagłówki, stopki)
///  - Osadzanie obrazów (base64 data URI → ImagePart relationship)
///  - Serializacja do MemoryStream i zwrot byte[]
///
/// Każde wywołanie tworzy nową instancję konwertera, bo klasa zawiera
/// mutable state (_imageRelationships, _imageCounter, etc.) niezerowany między wywołaniami.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class HtmlToDocxBenchmarks
{
    private string _tinyHtml = null!;
    private string _mediumHtml = null!;
    private string _largeHtml = null!;
    private string _htmlWithTable = null!;
    private string _htmlWithHeaderFooter = null!;
    private DocumentMetadata _metadata = null!;
    private HeaderFooterContent _header = null!;
    private HeaderFooterContent _footer = null!;

    [GlobalSetup]
    public void Setup()
    {
        _tinyHtml = "<p>Krótki dokument benchmarkowy. Jeden akapit, zero formatowania.</p>";

        _mediumHtml = BuildParagraphHtml(paragraphCount: 15, withFormatting: true);

        _largeHtml = BuildParagraphHtml(paragraphCount: 60, withFormatting: true);

        _htmlWithTable = BuildHtmlWithTable(rows: 20, cols: 5);

        _htmlWithHeaderFooter = BuildParagraphHtml(paragraphCount: 20, withFormatting: false);

        _metadata = new DocumentMetadata
        {
            Title = "Dokument Benchmarkowy",
            Author = "Benchmark Suite",
            Subject = "Testy wydajnościowe",
            Company = "D2ViewerEditor"
        };

        _header = new HeaderFooterContent
        {
            Html = "<p><b>Nagłówek dokumentu</b> — strona {page} z {pages}</p>",
            Height = 1.5
        };

        _footer = new HeaderFooterContent
        {
            Html = "<p>Stopka — wygenerowano automatycznie przez benchmark</p>",
            Height = 1.0
        };
    }

    /// <summary>
    /// Minimalny dokument — jeden akapit, bez formatowania.
    /// Baseline: mierzy wyłącznie narzut infrastruktury OpenXML (tworzenie pakietu, styles, sekcja).
    /// </summary>
    [Benchmark(Baseline = true)]
    public byte[] TinyDocument()
        => new HtmlToDocxConverter().Convert(_tinyHtml);

    /// <summary>
    /// Średni dokument — 15 akapitów z pogrubieniem i kursywą.
    /// Typowa złożoność dokumentu roboczego.
    /// </summary>
    [Benchmark]
    public byte[] MediumDocument()
        => new HtmlToDocxConverter().Convert(_mediumHtml);

    /// <summary>
    /// Duży dokument — 60 akapitów z formatowaniem.
    /// Mierzy skalowalność parsowania i budowania struktury XML dla dużych dokumentów.
    /// </summary>
    [Benchmark]
    public byte[] LargeDocument()
        => new HtmlToDocxConverter().Convert(_largeHtml);

    /// <summary>
    /// Dokument z tabelą 20×5 — testuje ścieżkę konwersji HTML table → Word table.
    /// Tabele Word wymagają zagnieżdżonej struktury: Table > Row > Cell > Paragraph > Run.
    /// </summary>
    [Benchmark]
    public byte[] DocumentWithTable()
        => new HtmlToDocxConverter().Convert(_htmlWithTable);

    /// <summary>
    /// Dokument z metadanymi, nagłówkiem i stopką.
    /// Mierzy koszt dodatkowych części OpenXML: HeaderPart, FooterPart, CoreProperties, ExtendedProperties.
    /// </summary>
    [Benchmark]
    public byte[] MediumDocumentWithHeaderFooter()
        => new HtmlToDocxConverter().Convert(
            _htmlWithHeaderFooter,
            metadata: _metadata,
            header: _header,
            footer: _footer);

    // --- Generatory danych testowych ---

    private static string BuildParagraphHtml(int paragraphCount, bool withFormatting)
    {
        var sb = new StringBuilder(paragraphCount * 120);
        for (int i = 1; i <= paragraphCount; i++)
        {
            if (withFormatting)
                sb.Append($"<p>Akapit {i}: Treść dokumentu zawierająca <b>pogrubiony tekst</b>, <i>kursywę</i> oraz <u>podkreślenie</u>. Numer referencyjny: REF-{i:D4}.</p>");
            else
                sb.Append($"<p>Akapit {i}: Zwykły tekst dokumentu bez dodatkowego formatowania. Numer referencyjny: REF-{i:D4}.</p>");
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

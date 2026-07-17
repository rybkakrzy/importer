using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Spis treści (pole TOC) jak w MS Word.
///
/// Poprzednio: (1) ConvertHyperlinkToHtml renderował z runów hyperlinka wyłącznie Text —
/// tabulator i pole PAGEREF wewnątrz w:hyperlink wpisu TOC ginęły (tytuł sklejony z numerem,
/// bez kropek); (2) leader tab-stopu nie był nigdzie malowany; (3) instrukcja „PAGEREF"
/// zawiera „PAGE", więc numer wpisu dostawał żywy placeholder {page} (numer BIEŻĄCEJ strony
/// zamiast zbuforowanego numeru celu); (4) pole TOC, PAGEREF, zakładki _Toc i kotwice
/// hyperlinków ginęły przy zapisie — Word tracił możliwość aktualizacji spisu (F9).
///
/// Teraz: wpis TOC renderuje się jako linia flex ze span.docx-tab-leader (kropki maluje CSS
/// GUI), numer strony to wartość zbuforowana, a TOC/PAGEREF/zakładki/kotwice round-tripują
/// markerami (docx-fld-marker, docx-bookmark, data-anchor).
/// </summary>
[TestFixture]
public class TocFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    /// <summary>
    /// Dokument à la Word: nagłówek z zakładką _Toc1 oraz spis treści — akapit TOC1
    /// z prawym tab-stopem leader=dot, polem TOC (begin/instr/separate … end) i wpisem
    /// w hyperlinku kotwicowym: „Rozdział pierwszy ⇥ PAGEREF _Toc1 (cache: 2)".
    /// endInSeparateParagraph=true przenosi fldChar end pola TOC do DRUGIEGO akapitu
    /// (tak zapisuje Word w spisie wielowpisowym — pole przekracza granice akapitów).
    /// </summary>
    private static MemoryStream BuildDocxWithToc(bool endInSeparateParagraph = false)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var tocEntry = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "TOC1" },
                    new Tabs(new TabStop
                    {
                        Val = TabStopValues.Right,
                        Leader = TabStopLeaderCharValues.Dot,
                        Position = 9062
                    })),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" TOC \\o \"1-3\" \\h \\z \\u ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Hyperlink(
                    new Run(new Text("Rozdział pierwszy")),
                    new Run(new TabChar()),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" PAGEREF _Toc1 \\h ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("2")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }))
                { Anchor = "_Toc1", History = true });

            body.Append(tocEntry);

            var tocEnd = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
            if (endInSeparateParagraph)
            {
                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "TOC1" },
                        new Tabs(new TabStop
                        {
                            Val = TabStopValues.Right,
                            Leader = TabStopLeaderCharValues.Dot,
                            Position = 9062
                        })),
                    new Hyperlink(
                        new Run(new Text("Rozdział drugi")),
                        new Run(new TabChar()),
                        new Run(new Text("3")))
                    { Anchor = "_Toc2", History = true },
                    tocEnd));
            }
            else
            {
                tocEntry.Append(tocEnd);
            }

            // Nagłówek-cel z zakładką _Toc1 (tak Word znakuje cele wpisów spisu).
            body.Append(new Paragraph(
                new BookmarkStart { Name = "_Toc1", Id = "1" },
                new Run(new Text("Rozdział pierwszy")),
                new BookmarkEnd { Id = "1" }));

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_TocEntry_RendersFlexLeaderLineWithCachedPageNumber()
    {
        using var stream = BuildDocxWithToc();

        var content = _reader.Convert(stream);
        var html = content.Html;

        // Linia flex z wypełniaczem kropkowym; stop round-tripuje w data-tab-stops.
        html.Should().Contain("display:flex");
        html.Should().Contain("docx-tab-leader");
        html.Should().Contain("data-leader=\"dot\"");
        html.Should().Contain("data-tab-stops=\"9062:right:dot\"");

        // Tytuł wpisu i ZBUFOROWANY numer strony celu — nie żywy placeholder {page}.
        html.Should().Contain("Rozdział pierwszy");
        html.Should().Contain(">2<");
        html.Should().NotContain("{page}");

        // Hyperlink kotwicowy wpisu: href="#_Toc1", bez wymuszonego niebieskiego.
        html.Should().Contain("href=\"#_Toc1\"");
        html.Should().Contain("data-anchor=\"_Toc1\"");

        // Instrukcje pól TOC i PAGEREF round-tripują markerami (niewidoczne w treści).
        html.Should().Contain("data-fld=\"begin\"");
        html.Should().Contain("data-fld-instr=\"TOC");
        html.Should().Contain("data-fld-instr=\"PAGEREF _Toc1");
        html.Should().Contain("data-fld=\"end\"");
        html.Should().NotContain(">PAGEREF");

        // Zakładka-cel przy nagłówku + styl wpisu TOC1.
        html.Should().Contain("data-bm-name=\"_Toc1\"");
        html.Should().Contain("data-style-id=\"TOC1\"");
    }

    [Test]
    public void Read_TocFieldSpanningParagraphs_EmitsBalancedMarkers()
    {
        using var stream = BuildDocxWithToc(endInSeparateParagraph: true);

        var html = _reader.Convert(stream).Html;

        // begin: pole TOC + PAGEREF pierwszego wpisu; end: PAGEREF + TOC (w DRUGIM akapicie).
        CountOf(html, "data-fld=\"begin\"").Should().Be(2);
        CountOf(html, "data-fld=\"end\"").Should().Be(2);
        html.IndexOf("data-fld-instr=\"TOC", StringComparison.Ordinal)
            .Should().BeLessThan(html.IndexOf("Rozdział pierwszy", StringComparison.Ordinal));
        html.LastIndexOf("data-fld=\"end\"", StringComparison.Ordinal)
            .Should().BeGreaterThan(html.IndexOf("Rozdział drugi", StringComparison.Ordinal));
    }

    [Test]
    public void Write_TocEntryHtml_RecreatesFieldHyperlinkBookmarkAndStyle()
    {
        using var stream = BuildDocxWithToc();
        var html = _reader.Convert(stream).Html;

        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // Pola: TOC i PAGEREF odtworzone jako fldChar begin+instrText+separate … end (balans).
        var fieldChars = body.Descendants<FieldChar>().ToList();
        fieldChars.Count(f => f.FieldCharType?.Value == FieldCharValues.Begin).Should().Be(2);
        fieldChars.Count(f => f.FieldCharType?.Value == FieldCharValues.End).Should().Be(2);
        var instructions = body.Descendants<FieldCode>().Select(f => f.Text.Trim()).ToList();
        instructions.Should().Contain(i => i.StartsWith("TOC"));
        instructions.Should().Contain(i => i.StartsWith("PAGEREF _Toc1"));

        // Hyperlinki kotwicowe wpisu: reader rozcina <a> na tabulatorze (tytuł | numer),
        // więc wracają DWA w:hyperlink o tej samej kotwicy — funkcjonalnie jak w Wordzie,
        // a F9 i tak regeneruje wpisy. Tab żyje między nimi, w akapicie wpisu.
        var anchored = body.Descendants<Hyperlink>().Where(h => h.Anchor?.Value == "_Toc1").ToList();
        anchored.Should().NotBeEmpty();
        string.Concat(anchored.Select(h => h.InnerText)).Should().Contain("Rozdział pierwszy").And.Contain("2");

        // Zakładka-cel i styl wpisu.
        body.Descendants<BookmarkStart>().Should().Contain(b => b.Name!.Value == "_Toc1");
        body.Descendants<ParagraphStyleId>().Should().Contain(s => s.Val!.Value == "TOC1");

        // Tab-stop z leaderem odtworzony w pPr, a znak tabulatora w akapicie wpisu.
        var tocPara = body.Descendants<Paragraph>()
            .First(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC1");
        var stop = tocPara.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().Single();
        stop.Leader!.Value.Should().Be(TabStopLeaderCharValues.Dot);
        stop.Position!.Value.Should().Be(9062);
        tocPara.Descendants<TabChar>().Should().HaveCount(1, "wypełniacz wraca jako pojedynczy w:tab");
    }

    [Test]
    public void RoundTrip_TwoSaves_KeepsTocStableWithoutDuplication()
    {
        using var stream = BuildDocxWithToc();
        var html1 = _reader.Convert(stream).Html;
        var bytes1 = _writer.Convert(html1);

        var html2 = _reader.Convert(new MemoryStream(bytes1)).Html;
        var bytes2 = _writer.Convert(html2);

        using var doc1 = WordprocessingDocument.Open(new MemoryStream(bytes1), false);
        using var doc2 = WordprocessingDocument.Open(new MemoryStream(bytes2), false);
        var body1 = doc1.MainDocumentPart!.Document!.Body!;
        var body2 = doc2.MainDocumentPart!.Document!.Body!;

        // Stabilność między kolejnymi zapisami — nic nie przyrasta ani nie znika.
        body2.Descendants<FieldChar>().Count(f => f.FieldCharType?.Value == FieldCharValues.Begin).Should().Be(2);
        body2.Descendants<FieldChar>().Count(f => f.FieldCharType?.Value == FieldCharValues.End).Should().Be(2);
        body2.Descendants<BookmarkStart>().Count(b => b.Name?.Value == "_Toc1").Should().Be(1);
        body2.Descendants<Hyperlink>().Count(h => h.Anchor?.Value == "_Toc1")
            .Should().Be(body1.Descendants<Hyperlink>().Count(h => h.Anchor?.Value == "_Toc1"),
                "rozcięcie <a> na tabulatorze musi być stabilne, nie narastać z każdym zapisem");
        body2.Descendants<TabChar>().Count().Should().Be(body1.Descendants<TabChar>().Count());

        // Numer strony wpisu nie zdublował się i nie zmienił w placeholder.
        var toc = body2.Descendants<Paragraph>()
            .First(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC1");
        toc.InnerText.Should().Contain("Rozdział pierwszy");
        toc.InnerText.Should().Contain("2");
        toc.InnerText.Should().NotContain("{page}");
    }

    [Test]
    public void Write_OrphanEndMarker_IsSkipped()
    {
        var html = "<p><span class=\"docx-fld-marker\" data-fld=\"end\" style=\"display:none;\"></span>Tekst</p>";

        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<FieldChar>().Should().BeEmpty(
            "osierocony fldChar End uszkadza dokument w Wordzie");
    }

    [Test]
    public void Write_MissingEndMarker_IsClosedAtEndOfBody()
    {
        var html = "<p><span class=\"docx-fld-marker\" data-fld=\"begin\" data-fld-instr=\"TOC \\o &quot;1-3&quot;\" style=\"display:none;\"></span>Wpis</p><p>Dalej</p>";

        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var fieldChars = doc.MainDocumentPart!.Document!.Body!.Descendants<FieldChar>().ToList();
        fieldChars.Count(f => f.FieldCharType?.Value == FieldCharValues.Begin).Should().Be(1);
        fieldChars.Count(f => f.FieldCharType?.Value == FieldCharValues.End).Should().Be(1,
            "niedomknięte pole domykamy na końcu treści");
    }

    private static int CountOf(string haystack, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }
        return count;
    }
}

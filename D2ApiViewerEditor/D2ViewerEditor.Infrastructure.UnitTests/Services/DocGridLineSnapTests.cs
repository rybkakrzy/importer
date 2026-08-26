using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// ADR-0108 r.5: siatka dokumentu (w:docGrid type=lines). Pomiar glifów w PDF z Worda: przy
/// snapToGrid (domyślnie ON) Word zastępuje jednostkę „single" SKOKIEM siatki — interlinia
/// auto N = N × pitch (1.15× przy pitch 18 pt = 20.7 pt), snapToGrid=0 → 1.15 × single fontu,
/// exact → bez zmian. Siatka jest PER SEKCJA. Reader renderuje snap jako pt + marker
/// <c>--w-line-grid:1</c>, z którego writer odtwarza w:line/lineRule=auto (nie exact).
/// </summary>
[TestFixture]
public class DocGridLineSnapTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void SetUp()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static string TagOf(string html, string text)
    {
        var i = html.IndexOf(text, StringComparison.Ordinal);
        i.Should().BeGreaterThan(0, $"brak tekstu '{text}'");
        var start = html.LastIndexOf("<p", i, StringComparison.Ordinal);
        return html.Substring(start, html.IndexOf('>', start) - start + 1);
    }

    [Test]
    public void Grid_section_snaps_auto_line_spacing_to_pitch_multiple()
    {
        using var ms = Docx(gridInFirstSection: true);

        var html = _reader.Convert(ms).Html;

        // docDefaults line=276 auto (1.15×) × pitch 360 tw (18 pt) = 20.7 pt
        var snap = TagOf(html, "Snap");
        snap.Should().Contain("line-height:20.7pt").And.Contain("--w-line-grid:1").And.Contain("--w-line-tw:276");
        // własny mnożnik 1.5× (line=360) → 27 pt
        TagOf(html, "Poltora").Should().Contain("line-height:27pt").And.Contain("--w-line-tw:360");
    }

    [Test]
    public void SnapToGrid_off_and_exact_ignore_the_grid()
    {
        using var ms = Docx(gridInFirstSection: true);

        var html = _reader.Convert(ms).Html;

        TagOf(html, "Bez snapu").Should().NotContain("--w-line-grid").And.Contain("--w-snap-to-grid:0");
        var exact = TagOf(html, "Dokladnie");
        exact.Should().Contain("line-height:12pt").And.NotContain("--w-line-grid");
    }

    [Test]
    public void Grid_is_per_section_only_second_section_snaps()
    {
        // Sekcja 1 bez siatki, sekcja 2 z siatką: akapit przed przerwą NIE snapuje, po — tak.
        using var ms = Docx(gridInFirstSection: false);

        var html = _reader.Convert(ms).Html;

        TagOf(html, "Przed przerwa").Should().NotContain("--w-line-grid");
        TagOf(html, "Snap").Should().Contain("--w-line-grid:1");
    }

    [Test]
    public void List_items_in_grid_section_snap_too()
    {
        // Punkty list idą osobną ścieżką readera (ConvertConsecutiveListItems) — bez snapu
        // lista w sekcji z siatką renderowała się z interlinią fontu zamiast siatki.
        using var ms = Docx(gridInFirstSection: true);

        var html = _reader.Convert(ms).Html;
        var i = html.IndexOf("Punkt listy", StringComparison.Ordinal);
        i.Should().BeGreaterThan(0);
        var start = html.LastIndexOf("<li", i, StringComparison.Ordinal);
        var li = html.Substring(start, html.IndexOf('>', start) - start + 1);

        li.Should().Contain("line-height:20.7pt").And.Contain("--w-line-grid:1").And.Contain("--w-line-tw:276");
    }

    [Test]
    public void Writer_restores_auto_line_from_grid_marker_not_exact()
    {
        var html = "<p style=\"line-height:20.7pt;--w-line-tw:276;--w-line-grid:1;\">Snap</p>" +
                   "<p style=\"line-height:12pt;--w-line-rule:exact;\">Dokladnie</p>";

        using var ms = new MemoryStream(_writer.Convert(html));
        using var doc = WordprocessingDocument.Open(ms, false);
        var paras = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();

        var s0 = paras[0].ParagraphProperties!.GetFirstChild<SpacingBetweenLines>()!;
        s0.Line!.Value.Should().Be("276");
        s0.LineRule!.Value.Should().Be(LineSpacingRuleValues.Auto);

        var s1 = paras[1].ParagraphProperties!.GetFirstChild<SpacingBetweenLines>()!;
        s1.Line!.Value.Should().Be("240");
        s1.LineRule!.Value.Should().Be(LineSpacingRuleValues.Exact);
    }

    private static MemoryStream Docx(bool gridInFirstSection)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            SectionProperties Sect(bool grid)
            {
                var sp = new SectionProperties(
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 936, Bottom = 936, Left = 1080, Right = 1080, Header = 720, Footer = 720 });
                sp.Append(grid
                    ? new DocGrid { Type = DocGridValues.Lines, LinePitch = 360 }
                    : new DocGrid { LinePitch = 360 });
                return sp;
            }

            if (!gridInFirstSection)
            {
                body.Append(new Paragraph(new Run(new Text("Przed przerwa sekcji."))));
                // akapit z sectPr = koniec sekcji 1 (bez siatki), następna sekcja ma siatkę
                var brk = new Paragraph(new ParagraphProperties(Sect(grid: false)));
                body.Append(brk);
            }

            body.Append(new Paragraph(new Run(new Text("Snap do siatki (domyślnie)."))));
            body.Append(new Paragraph(
                new ParagraphProperties(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto }),
                new Run(new Text("Poltora na siatce."))));
            body.Append(new Paragraph(
                new ParagraphProperties(new SnapToGrid { Val = false }),
                new Run(new Text("Bez snapu."))));
            body.Append(new Paragraph(
                new ParagraphProperties(new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Exact }),
                new Run(new Text("Dokladnie 12 pt."))));
            body.Append(new Paragraph(
                new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                new Run(new Text("Punkt listy na siatce."))));
            body.Append(Sect(grid: true));
            mainPart.Document = new Document(body);

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(new Level(
                    new NumberingFormat { Val = NumberFormatValues.Bullet },
                    new LevelText { Val = "•" },
                    new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" }))
                { LevelIndex = 0 })
                { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            numberingPart.Numbering.Save();

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { After = "0", Line = "276", LineRule = LineSpacingRuleValues.Auto }))));
            styles.Append(new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }
}

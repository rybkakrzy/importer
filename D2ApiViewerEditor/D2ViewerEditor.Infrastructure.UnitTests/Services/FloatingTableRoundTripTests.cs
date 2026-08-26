using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// ADR-0108 r.6: tabele pływające (w:tblpPr). Dotąd reader ignorował tblpPr (tabela w przepływie
/// na pełną szerokość), a writer go nie zapisywał — pozycjonowanie GINĘŁO przy każdym zapisie.
/// Round-trip: wszystkie atrybuty w data-tblp-*; render: CSS float (pomiar PDF Worda:
/// FT07 tblpX=2200 tw → 110 pt od krawędzi strony = margin-left od marginesu 964 tw).
/// Osobno: w:tblW dxa mniejsze niż suma tblGrid przy autofit → kolumny podglądu skalowane
/// do tblW (Word układa w szerokości preferowanej; przeglądarka rozszerzała do sumy <col>).
/// </summary>
[TestFixture]
public class FloatingTableRoundTripTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void SetUp()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static string TableTag(string html)
    {
        var s = html.IndexOf("<table", StringComparison.Ordinal);
        s.Should().BeGreaterThan(0);
        return html.Substring(s, html.IndexOf('>', s) - s + 1);
    }

    [Test]
    public void Absolute_x_on_page_anchor_floats_left_with_offset_from_margin()
    {
        // tblpX=2200 tw od krawędzi strony, margines 964 tw → offset 1236 tw = 82 px
        using var ms = Docx(new TablePositionProperties
        {
            HorizontalAnchor = HorizontalAnchorValues.Page, VerticalAnchor = VerticalAnchorValues.Page,
            TablePositionX = 2200, TablePositionY = 3000,
            LeftFromText = 144, RightFromText = 144, TopFromText = 144, BottomFromText = 144
        });

        var tag = TableTag(_reader.Convert(ms).Html);

        tag.Should().Contain("data-tblp=\"1\"").And.Contain("data-tblp-horz-anchor=\"page\"")
            .And.Contain("data-tblp-x-tw=\"2200\"").And.Contain("data-tblp-y-tw=\"3000\"")
            .And.Contain("data-tblp-left-tw=\"144\"");
        tag.Should().Contain("float:left;");
        // margin: top right bottom left → 144 tw = 9.6 px (TwipsToPx obcina → 9); left = offset 82 px
        tag.Should().Contain("margin:9px 9px 9px 82px;");
    }

    [Test]
    public void Right_aligned_floating_table_floats_right()
    {
        using var ms = Docx(new TablePositionProperties
        {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Margin,
            TablePositionXAlignment = HorizontalAlignmentValues.Right, TablePositionY = 2200,
            LeftFromText = 144, RightFromText = 144
        });

        var tag = TableTag(_reader.Convert(ms).Html);

        tag.Should().Contain("float:right;").And.Contain("data-tblp-xspec=\"right\"");
    }

    [Test]
    public void Preferred_dxa_width_wins_over_wider_grid_for_autofit_table()
    {
        // tblW=3600 (240 px), tblGrid 5156+5156 (686 px) → kolumny skalowane do 240 px
        using var ms = Docx(null);

        var html = _reader.Convert(ms).Html;
        var tag = TableTag(html);
        var colgroup = html.Substring(html.IndexOf("<colgroup", StringComparison.Ordinal), 200);

        tag.Should().Contain("width:240px;");
        colgroup.Should().Contain("width:120px;").And.Contain("data-w-tw=\"5156\"");
    }

    [Test]
    public void Floating_position_survives_save_and_reopen()
    {
        using var ms = Docx(new TablePositionProperties
        {
            HorizontalAnchor = HorizontalAnchorValues.Page, VerticalAnchor = VerticalAnchorValues.Page,
            TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = 2800,
            LeftFromText = 360, RightFromText = 360, TopFromText = 360, BottomFromText = 360
        });
        var first = _reader.Convert(ms).Html;

        using var saved = new MemoryStream(_writer.Convert(first));
        using var doc = WordprocessingDocument.Open(saved, false);
        var tblp = doc.MainDocumentPart!.Document.Body!.Descendants<TablePositionProperties>().Single();

        tblp.HorizontalAnchor!.Value.Should().Be(HorizontalAnchorValues.Page);
        tblp.VerticalAnchor!.Value.Should().Be(VerticalAnchorValues.Page);
        tblp.TablePositionXAlignment!.Value.Should().Be(HorizontalAlignmentValues.Center);
        tblp.TablePositionY!.Value.Should().Be(2800);
        tblp.LeftFromText!.Value.Should().Be((short)360);
        tblp.BottomFromText!.Value.Should().Be((short)360);
        // tblpPr zaraz po tblStyle (kolejność CT_TblPr)
        var props = tblp.Parent as TableProperties;
        props!.Elements().First().Should().BeOfType<TableStyle>();
        props.Elements().ElementAt(1).Should().BeOfType<TablePositionProperties>();
    }

    private static MemoryStream Docx(TablePositionProperties? tblp)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var tblPr = new TableProperties(new TableStyle { Val = "TableGrid" });
            if (tblp != null) tblPr.Append(tblp);
            tblPr.Append(new TableWidth { Width = "3600", Type = TableWidthUnitValues.Dxa });
            tblPr.Append(new TableLayout { Type = TableLayoutValues.Autofit });

            var table = new Table(tblPr,
                new TableGrid(new GridColumn { Width = "5156" }, new GridColumn { Width = "5156" }),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("FT07")))),
                    new TableCell(new Paragraph(new Run(new Text("floating"))))));

            var body = new Body(
                new Paragraph(new Run(new Text("Anchor paragraph."))),
                table,
                new Paragraph(new Run(new Text("Surrounding flow text."))),
                new SectionProperties(
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 964, Bottom = 964, Left = 964, Right = 964, Header = 720, Footer = 720 }));
            mainPart.Document = new Document(body);

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(new StyleName { Val = "Table Grid" }) { Type = StyleValues.Table, StyleId = "TableGrid" });
            stylesPart.Styles.Save();
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }
}

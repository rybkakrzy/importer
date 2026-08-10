using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Szerokość tabeli tblW=auto z pełną siatką tblGrid.
///
/// Poprzednio reader emitował inline width:auto — przeglądarka robiła shrink-to-fit
/// i tabela była wyraźnie węższa niż w Wordzie (Word układa wg zapisanej siatki).
/// Teraz: szerokość = suma gridCol + table-layout:fixed (renderowo), a oryginalna
/// semantyka wraca przez markery data-tbl-w="auto" / data-tbl-layout="autofit".
/// </summary>
[TestFixture]
public class TableAutoWidthFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream BuildDocx(TableProperties tblPr, int cols = 2, int colWidthTw = 3000)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var table = new Table();
            table.Append(tblPr);
            var grid = new TableGrid();
            for (var c = 0; c < cols; c++)
                grid.Append(new GridColumn { Width = colWidthTw.ToString() });
            table.Append(grid);

            var row = new TableRow();
            for (var c = 0; c < cols; c++)
            {
                row.Append(new TableCell(
                    new TableCellProperties(new TableCellWidth { Width = colWidthTw.ToString(), Type = TableWidthUnitValues.Dxa }),
                    new Paragraph(new Run(new Text($"C{c}")))));
            }
            table.Append(row);

            body.Append(table);
            body.Append(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static string FirstTableTag(string html)
    {
        var start = html.IndexOf("<table", StringComparison.Ordinal);
        var end = html.IndexOf('>', start);
        return html.Substring(start, end - start + 1);
    }

    private static Table FirstTable(byte[] docx)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return (Table)doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().First().CloneNode(true);
    }

    [Test]
    public void Read_TblWAuto_WithFullGrid_RendersGridWidthWithSemanticsMarkers()
    {
        var tblPr = new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        var html = _reader.Convert(BuildDocx(tblPr)).Html;

        var tag = FirstTableTag(html);
        // 2 × 3000 tw = 2 × 200 px @96 DPI.
        tag.Should().Contain("width:400px");
        tag.Should().Contain("table-layout:fixed");
        tag.Should().Contain("data-tbl-w=\"auto\"");
        tag.Should().Contain("data-tbl-layout=\"autofit\"");
    }

    [Test]
    public void Read_TblWAuto_WithoutGridWidths_KeepsAutoWithoutMarkers()
    {
        var tblPr = new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        var html = _reader.Convert(BuildDocx(tblPr, colWidthTw: 0)).Html;

        var tag = FirstTableTag(html);
        tag.Should().Contain("width:auto");
        tag.Should().NotContain("data-tbl-w=");
        tag.Should().NotContain("data-tbl-layout=");
    }

    [Test]
    public void Read_TblWDxa_KeepsExplicitWidthWithoutAutoMarker()
    {
        var tblPr = new TableProperties(new TableWidth { Width = "6000", Type = TableWidthUnitValues.Dxa });
        var html = _reader.Convert(BuildDocx(tblPr)).Html;

        var tag = FirstTableTag(html);
        tag.Should().Contain("width:400px");
        tag.Should().NotContain("data-tbl-w=");
    }

    [Test]
    public void RoundTrip_TblWAuto_RestoresAutoAndAutofit_AndKeepsGrid()
    {
        var tblPr = new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        var html = _reader.Convert(BuildDocx(tblPr)).Html;

        var table = FirstTable(_writer.Convert(html));
        var outPr = table.GetFirstChild<TableProperties>()!;

        outPr.TableWidth!.Type!.Value.Should().Be(TableWidthUnitValues.Auto,
            "renderowe width:px nie może utrwalić się jako tblW dxa");
        outPr.GetFirstChild<TableLayout>()!.Type!.Value.Should().Be(TableLayoutValues.Autofit,
            "renderowe table-layout:fixed nie może utrwalić się jako tblLayout fixed");
        table.GetFirstChild<TableGrid>()!.Elements<GridColumn>()
            .Select(g => g.Width!.Value).Should().Equal("3000", "3000");
    }

    [Test]
    public void Write_TableWithoutMarkers_KeepsPixelWidthAsDxa()
    {
        // Bez markera (np. użytkownik ręcznie zmienił szerokość — GUI zdejmuje data-tbl-w)
        // px staje się jawnym tblW dxa, jak dotychczas.
        var html = "<table style=\"border-collapse:collapse;width:400px;table-layout:fixed;\">" +
                   "<tr><td>A</td></tr></table>";

        var outPr = FirstTable(_writer.Convert(html)).GetFirstChild<TableProperties>()!;
        outPr.TableWidth!.Type!.Value.Should().Be(TableWidthUnitValues.Dxa);
        outPr.TableWidth!.Width!.Value.Should().Be("6000");
        outPr.GetFirstChild<TableLayout>()!.Type!.Value.Should().Be(TableLayoutValues.Fixed);
    }
}

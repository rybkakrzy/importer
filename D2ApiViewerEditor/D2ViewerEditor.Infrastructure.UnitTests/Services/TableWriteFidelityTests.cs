using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Zapis tabel HTML→DOCX: geometria kolumn i scalenia pionowe.
///
/// Poprzednio writer (1) odtwarzał w:tblGrid z samej LICZBY komórek — bez szerokości,
/// więc układ kolumn Worda ginął przy każdym autosave; (2) zawsze wymuszał
/// TableLayout=Autofit, gubiąc fixed layout; (3) nie emitował komórek kontynuacji
/// w:vMerge dla rowspan — HTML pomija komórki przykryte scaleniem, a OOXML wymaga ich
/// jawnie, więc komórki kolejnych wierszy przesuwały się w lewo (tabela uszkodzona w Wordzie).
/// </summary>
[TestFixture]
public class TableWriteFidelityTests
{
    private HtmlToDocxConverter _writer = null!;
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup()
    {
        _writer = new HtmlToDocxConverter();
        _reader = new DocxToHtmlConverter();
    }

    private static Table FirstTable(byte[] docx)
    {
        var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().First();
    }

    [Test]
    public void Write_ColgroupWidths_LandInTblGrid()
    {
        // 200px i 400px → 3000 i 6000 twips (px * 15).
        var html =
            "<table style=\"border-collapse:collapse;width:600px;table-layout:fixed;\">" +
            "<colgroup><col style=\"width:200px;\" /><col style=\"width:400px;\" /></colgroup>" +
            "<tr><td>A</td><td>B</td></tr>" +
            "</table>";

        var table = FirstTable(_writer.Convert(html));

        var cols = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        cols.Should().HaveCount(2);
        int.Parse(cols[0].Width!.Value!).Should().BeCloseTo(3000, 20);
        int.Parse(cols[1].Width!.Value!).Should().BeCloseTo(6000, 20);
    }

    [Test]
    public void Write_TableLayoutFixed_IsPreserved()
    {
        var html =
            "<table style=\"table-layout:fixed;width:600px;\">" +
            "<colgroup><col style=\"width:300px;\" /><col style=\"width:300px;\" /></colgroup>" +
            "<tr><td>A</td><td>B</td></tr></table>";

        var table = FirstTable(_writer.Convert(html));

        table.GetFirstChild<TableProperties>()!.GetFirstChild<TableLayout>()!
            .Type!.Value.Should().Be(TableLayoutValues.Fixed);
    }

    [Test]
    public void Write_TableWithoutFixedLayout_KeepsAutofit()
    {
        var html = "<table><tr><td>A</td><td>B</td></tr></table>";

        var table = FirstTable(_writer.Convert(html));

        table.GetFirstChild<TableProperties>()!.GetFirstChild<TableLayout>()!
            .Type!.Value.Should().Be(TableLayoutValues.Autofit);
    }

    [Test]
    public void Write_Rowspan_EmitsVerticalMergeContinuationCell()
    {
        // Wiersz 2 ma w HTML tylko JEDNĄ komórkę (pierwszą przykrywa rowspan z wiersza 1).
        var html =
            "<table>" +
            "<tr><td rowspan=\"2\">scalona</td><td>r1c2</td></tr>" +
            "<tr><td>r2c2</td></tr>" +
            "</table>";

        var table = FirstTable(_writer.Convert(html));
        var rows = table.Elements<TableRow>().ToList();

        rows.Should().HaveCount(2);
        rows[0].Elements<TableCell>().Should().HaveCount(2);

        var row2Cells = rows[1].Elements<TableCell>().ToList();
        row2Cells.Should().HaveCount(2, "OOXML wymaga jawnej komórki kontynuacji pod komórką z rowspan");

        var continuation = row2Cells[0].TableCellProperties!.Elements<VerticalMerge>().Single();
        continuation.Val.Should().BeNull("vMerge bez w:val oznacza continue");

        // Komórka startowa scalenia ma vMerge=restart.
        rows[0].Elements<TableCell>().First().TableCellProperties!
            .Elements<VerticalMerge>().Single().Val!.Value.Should().Be(MergedCellValues.Restart);
    }

    [Test]
    public void Write_RowspanInMiddleColumn_ContinuationLandsAtCorrectGridPosition()
    {
        var html =
            "<table>" +
            "<tr><td>r1c1</td><td rowspan=\"2\">scalona</td><td>r1c3</td></tr>" +
            "<tr><td>r2c1</td><td>r2c3</td></tr>" +
            "</table>";

        var table = FirstTable(_writer.Convert(html));
        var row2Cells = table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().ToList();

        row2Cells.Should().HaveCount(3);
        // Kontynuacja w ŚRODKOWEJ kolumnie — nie na początku, nie na końcu.
        row2Cells[0].TableCellProperties?.Elements<VerticalMerge>().Should().BeNullOrEmpty();
        row2Cells[1].TableCellProperties!.Elements<VerticalMerge>().Single().Val.Should().BeNull();
        row2Cells[2].TableCellProperties?.Elements<VerticalMerge>().Should().BeNullOrEmpty();
    }

    [Test]
    public void RoundTrip_TableColumnWidths_SurviveDocxHtmlDocx()
    {
        // DOCX z tblGrid 3000/6000 tw → HTML (colgroup) → DOCX: szerokości nie mogą zginąć.
        var srcHtml =
            "<table style=\"border-collapse:collapse;width:600px;table-layout:fixed;\">" +
            "<colgroup><col style=\"width:200px;\" /><col style=\"width:400px;\" /></colgroup>" +
            "<tr><td>A</td><td>B</td></tr></table>";
        var firstDocx = _writer.Convert(srcHtml);

        var content = _reader.Convert(new MemoryStream(firstDocx));
        content.Html.Should().Contain("<colgroup>");
        content.Html.Should().Contain("table-layout:fixed");

        var secondDocx = _writer.Convert(content.Html);
        var cols = FirstTable(secondDocx).GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();

        cols.Should().HaveCount(2);
        int.Parse(cols[0].Width!.Value!).Should().BeCloseTo(3000, 40);
        int.Parse(cols[1].Width!.Value!).Should().BeCloseTo(6000, 40);
    }
}

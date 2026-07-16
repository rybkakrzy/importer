using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Wierność w:gridBefore/w:gridAfter (wiersz „wcięty" w siatce tabeli — puste sloty kolumn
/// przed pierwszą / za ostatnią komórką, bez realnych komórek).
///
/// Poprzednio reader całkowicie ignorował te właściwości: pierwsza komórka wciętego wiersza
/// była przez `table-layout:fixed`+colgroup przypinana do kolumny 1, a mechanizm „deficytu
/// kolumn" doklejał brakujące sloty do OSTATNIEJ komórki. Realny przypadek: tabela stopki Qutalo
/// (styl 02Qutalododatkowa, siatka [426, 9214, 991, 426] tw) — tytuł dokumentu o szerokości
/// kolumny 2 (9214 tw ≈ 614 px) lądował w kolumnie 28 px i łamał się słowo-po-słowie,
/// pompując pasmo stopki do ~170 px, co z kolei zaniżało wysokość kolumn paginacji
/// (availableFor/pageBodyHeights) na KAŻDEJ stronie dokumentu.
/// </summary>
[TestFixture]
public class TableGridBeforeAfterFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    // ---------- helpers ----------

    /// <summary>Tabela à la stopka Qutalo: siatka 4 kolumn, wiersz 2 wcięty gridBefore/gridAfter.</summary>
    private static MemoryStream BuildQutaloFooterLikeDocx()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var table = new Table();
            table.Append(new TableProperties(
                new TableWidth { Width = "11057", Type = TableWidthUnitValues.Dxa },
                new TableLayout { Type = TableLayoutValues.Fixed }));
            table.Append(new TableGrid(
                new GridColumn { Width = "426" },
                new GridColumn { Width = "9214" },
                new GridColumn { Width = "991" },
                new GridColumn { Width = "426" }));

            // Wiersz 1: jedna komórka na całą szerokość (gridSpan=4).
            table.Append(new TableRow(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "11057", Type = TableWidthUnitValues.Dxa },
                    new GridSpan { Val = 4 }),
                new Paragraph(new Run(new Text("Pełna szerokość"))))));

            // Wiersz 2: gridBefore=1 (pusty slot 426 tw), tytuł (9214), numer strony (991),
            // gridAfter=1 (pusty slot 426 tw).
            table.Append(new TableRow(
                new TableRowProperties(
                    new GridBefore { Val = 1 },
                    new GridAfter { Val = 1 },
                    new WidthBeforeTableRow { Width = "426", Type = TableWidthUnitValues.Dxa },
                    new WidthAfterTableRow { Width = "426", Type = TableWidthUnitValues.Dxa }),
                new TableCell(
                    new TableCellProperties(new TableCellWidth { Width = "9214", Type = TableWidthUnitValues.Dxa }),
                    new Paragraph(new Run(new Text("Umowa otwarcia i prowadzenia zamkniętego mieszkaniowego rachunku powierniczego")))),
                new TableCell(
                    new TableCellProperties(new TableCellWidth { Width = "991", Type = TableWidthUnitValues.Dxa }),
                    new Paragraph(new Run(new Text("3/18"))))));

            body.Append(table);
            body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Table FirstTable(byte[] docx)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return (Table)doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().First().CloneNode(true);
    }

    /// <summary>HTML drugiego wiersza pierwszej tabeli (fragment &lt;tr&gt;…&lt;/tr&gt;).</summary>
    private static string SecondRowHtml(string html)
    {
        var matches = System.Text.RegularExpressions.Regex.Matches(
            html, "<tr[^>]*>.*?</tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
        matches.Count.Should().BeGreaterThanOrEqualTo(2, "tabela musi mieć dwa wiersze");
        return matches[1].Value;
    }

    // ---------- reader ----------

    [Test]
    public void Read_GridBefore_EmitsLeadingSpacer_AndDoesNotPadLastCell()
    {
        var content = _reader.Convert(BuildQutaloFooterLikeDocx());
        var row2 = SecondRowHtml(content.Html);

        var tds = System.Text.RegularExpressions.Regex.Matches(row2, "<td[^>]*>")
            .Select(m => m.Value).ToList();
        tds.Should().HaveCount(4, "spacer przed + tytuł + numer strony + spacer po");

        // Spacer wiodący konsumuje slot gridBefore — bez niego table-layout:fixed
        // przypina komórkę tytułu do kolumny 426 tw (28 px) i tekst łamie się słowo-po-słowie.
        tds[0].Should().Contain("data-grid-spacer=\"before\"");
        tds[0].Should().NotContain("colspan", "gridBefore=1 = jeden slot");

        // Komórki merytoryczne bez sztucznych colspanów (stary „deficyt kolumn" doklejał
        // ostatniej komórce colspan=3 i przesuwał cały wiersz).
        tds[1].Should().NotContain("colspan");
        tds[1].Should().NotContain("data-grid-spacer");
        tds[2].Should().NotContain("colspan");

        tds[3].Should().Contain("data-grid-spacer=\"after\"");
    }

    [Test]
    public void Read_ShortRowWithoutGridBefore_StillPadsLastCell()
    {
        // Regresja: krótki wiersz BEZ gridBefore nadal dostaje deficyt na OSTATNIEJ komórce.
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var table = new Table(
                new TableGrid(
                    new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" },
                    new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("A")))),
                    new TableCell(new Paragraph(new Run(new Text("B"))))));
            mainPart.Document.Body!.Append(table);
            mainPart.Document.Save();
        }
        ms.Position = 0;

        var html = _reader.Convert(ms).Html;
        html.Should().Contain("colspan=\"3\"", "2 komórki w siatce 4 kolumn → ostatnia dostaje deficyt 2+1");
        html.Should().NotContain("data-grid-spacer");
    }

    // ---------- writer ----------

    [Test]
    public void Write_GridSpacers_RoundTripToGridBeforeAfter()
    {
        var html =
            "<table style=\"border-collapse:collapse;width:737px;table-layout:fixed;\">" +
            "<colgroup>" +
            "<col style=\"width:28px;\" data-w-tw=\"426\" /><col style=\"width:614px;\" data-w-tw=\"9214\" />" +
            "<col style=\"width:66px;\" data-w-tw=\"991\" /><col style=\"width:28px;\" data-w-tw=\"426\" />" +
            "</colgroup>" +
            "<tr><td colspan=\"4\">Pełna szerokość</td></tr>" +
            "<tr>" +
            "<td data-grid-spacer=\"before\" style=\"border:none;padding:0;\"></td>" +
            "<td>Umowa otwarcia i prowadzenia</td>" +
            "<td>3/18</td>" +
            "<td data-grid-spacer=\"after\" style=\"border:none;padding:0;\"></td>" +
            "</tr></table>";

        var rows = FirstTable(_writer.Convert(html)).Elements<TableRow>().ToList();
        rows.Should().HaveCount(2);

        var trPr = rows[1].TableRowProperties!;
        trPr.GetFirstChild<GridBefore>()!.Val!.Value.Should().Be(1);
        trPr.GetFirstChild<GridAfter>()!.Val!.Value.Should().Be(1);
        trPr.GetFirstChild<WidthBeforeTableRow>()!.Width!.Value.Should().Be("426");
        trPr.GetFirstChild<WidthAfterTableRow>()!.Width!.Value.Should().Be("426");

        // Spacery NIE są komórkami — wiersz ma dokładnie 2 realne w:tc.
        var cells = rows[1].Elements<TableCell>().ToList();
        cells.Should().HaveCount(2);

        // Kursor siatki uwzględnia slot gridBefore: tytuł leży nad kolumną 2 → tcW=9214.
        cells[0].TableCellProperties!.GetFirstChild<TableCellWidth>()!.Width!.Value.Should().Be("9214");
        cells[1].TableCellProperties!.GetFirstChild<TableCellWidth>()!.Width!.Value.Should().Be("991");
    }

    [Test]
    public void RoundTrip_QutaloFooterLikeTable_PreservesGridBeforeAfter()
    {
        var content = _reader.Convert(BuildQutaloFooterLikeDocx());
        var rows = FirstTable(_writer.Convert(content.Html)).Elements<TableRow>().ToList();

        var trPr = rows[1].TableRowProperties!;
        trPr.GetFirstChild<GridBefore>()!.Val!.Value.Should().Be(1);
        trPr.GetFirstChild<GridAfter>()!.Val!.Value.Should().Be(1);
        rows[1].Elements<TableCell>().Should().HaveCount(2, "spacery nie mogą stać się realnymi komórkami");
    }
}

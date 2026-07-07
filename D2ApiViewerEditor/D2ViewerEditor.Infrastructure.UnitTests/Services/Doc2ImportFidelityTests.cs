using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Regresja dla wady odwzorowania dokumentów Word zgłoszonych na dokumencie "Doc2"
/// (wyciąg bankowy ING): tabulatory lewe/prawe w treści, pionowe wyśrodkowanie komórek,
/// scalone kolumny w wierszach nieregularnych oraz bezpieczne SVG jako data-URI.
/// </summary>
[TestFixture]
public class Doc2ImportFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream Docx(params OpenXmlElement[] bodyChildren)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var c in bodyChildren) body.Append(c);
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    // ---- Issue 1: left / right / multiple tab stops in the body -----------------------

    [Test]
    public void BodyLeftTab_PlacesFollowingSegmentAtStopStart()
    {
        // left tab @1440 tw = 96 px: the text AFTER the tab starts at the stop (no translate).
        using var ms = Docx(new Paragraph(
            new ParagraphProperties(new Tabs(new TabStop { Val = TabStopValues.Left, Position = 1440 })),
            new Run(new Text("Before")), new Run(new TabChar()), new Run(new Text("After"))));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("data-tab-align=\"left\"");
        html.Should().Contain("left:96px");
        html.Should().NotContain("translateX");                 // left = brak przesunięcia
        html.Should().Contain("Before").And.Contain("After");   // treść przed i za tabem
    }

    [Test]
    public void BodyRightTab_AlignsFollowingSegmentEndToStop()
    {
        using var ms = Docx(new Paragraph(
            new ParagraphProperties(new Tabs(new TabStop { Val = TabStopValues.Right, Position = 9000 })),
            new Run(new Text("Podpis")), new Run(new TabChar()), new Run(new Text("Data"))));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("data-tab-align=\"right\"");
        html.Should().Contain("left:600px");            // 9000 tw @96 DPI
        html.Should().Contain("translateX(-100%)");     // prawy = koniec treści NA pozycji
    }

    [Test]
    public void BodyMultipleTabs_ProducesOneSegmentPerStop()
    {
        using var ms = Docx(new Paragraph(
            new ParagraphProperties(new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = 2000 },
                new TabStop { Val = TabStopValues.Center, Position = 5000 },
                new TabStop { Val = TabStopValues.Right, Position = 9000 })),
            new Run(new Text("A")), new Run(new TabChar()),
            new Run(new Text("B")), new Run(new TabChar()),
            new Run(new Text("C")), new Run(new TabChar()),
            new Run(new Text("D"))));

        var html = _reader.Convert(ms).Html;

        System.Text.RegularExpressions.Regex.Matches(html, "docx-tab-seg").Count.Should().Be(3);
        html.Should().Contain("data-tab-align=\"left\"")
            .And.Contain("data-tab-align=\"center\"")
            .And.Contain("data-tab-align=\"right\"");
        html.Should().Contain("A").And.Contain("B").And.Contain("C").And.Contain("D");
    }

    [Test]
    public void BodyTabs_RoundTripPreservesStopsAndTabCount()
    {
        var html = "<p data-tab-stops=\"1440:left;9000:right\">Before"
                   + "<span class=\"docx-tab-seg\" data-tab-align=\"left\">Mid</span>"
                   + "<span class=\"docx-tab-seg\" data-tab-align=\"right\">End</span></p>";

        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var para = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var tabs = para.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().ToList();
        tabs.Should().HaveCount(2);
        para.Descendants<TabChar>().Should().HaveCount(2);
        para.InnerText.Should().Contain("Before").And.Contain("Mid").And.Contain("End");
    }

    // ---- Issue 2: cell vertical alignment ---------------------------------------------

    [Test]
    public void CellVerticalCenter_EmitsSingleMiddle()
        => AssertVerticalAlignment(TableVerticalAlignmentValues.Center, "middle");

    [Test]
    public void CellVerticalBottom_EmitsSingleBottom()
        => AssertVerticalAlignment(TableVerticalAlignmentValues.Bottom, "bottom");

    private void AssertVerticalAlignment(TableVerticalAlignmentValues val, string css)
    {
        var table = SingleCellTable(new TableCellProperties(new TableCellVerticalAlignment { Val = val }));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain($"vertical-align:{css};");
        // Brak zduplikowanej deklaracji (wcześniej „top" szło zawsze, a wartość dopisywana drugi raz).
        System.Text.RegularExpressions.Regex.Matches(html, "vertical-align:").Count.Should().Be(1);
    }

    [Test]
    public void CellWithoutVerticalAlignment_DefaultsToTopOnce()
    {
        var table = SingleCellTable(new TableCellProperties());

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        System.Text.RegularExpressions.Regex.Matches(html, "vertical-align:").Count.Should().Be(1);
        html.Should().Contain("vertical-align:top;");
    }

    private static Table SingleCellTable(TableCellProperties cellProps) => new(
        new TableProperties(new TableStyle { Val = "TableGrid" }),
        new TableGrid(new GridColumn { Width = "3000" }),
        new TableRow(new TableCell(cellProps, new Paragraph(new Run(new Text("x"))))));

    // ---- Issue 3: merged columns / irregular rows -------------------------------------

    [Test]
    public void ShortSingleCellRow_SpansFullGrid()
    {
        // 3-column grid; a row with ONE cell and no gridSpan must span all three columns
        // (the ING "Umowa wieloproduktowa…" defect where content was pinned to column 1).
        var table = ThreeColTable(
            Row(Cell("H1"), Cell("H2"), Cell("H3")),
            Row(Cell("MERGED-ALL")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain("<td colspan=\"3\"");
        html.Should().Contain("MERGED-ALL");
    }

    [Test]
    public void ShortRow_LeadingCellThenContent_ExtendsLastCellOnly()
    {
        var table = ThreeColTable(
            Row(Cell("H1"), Cell("H2"), Cell("H3")),
            Row(Cell("Lp"), Cell("WIDE")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        // Pierwsza komórka normalna, druga (treść) pochłania brakujące 2 kolumny.
        html.Should().Contain("<td style=").And.Contain("Lp");
        html.Should().Contain("<td colspan=\"2\"");
        html.Should().Contain("WIDE");
        html.Should().NotContain("colspan=\"3\"");
    }

    [Test]
    public void ExplicitGridSpanRow_IsUnchanged_AndContentNotDuplicated()
    {
        var table = ThreeColTable(
            Row(Cell("H1"), Cell("H2"), Cell("H3")),
            Row(Cell("SPANNED", 3)));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain("<td colspan=\"3\"");
        System.Text.RegularExpressions.Regex.Matches(html, "SPANNED").Count.Should().Be(1);
    }

    [Test]
    public void FullRows_AreNotAlteredByShortRowLogic()
    {
        var table = ThreeColTable(
            Row(Cell("A"), Cell("B"), Cell("C")),
            Row(Cell("D"), Cell("E"), Cell("F")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().NotContain("colspan");
    }

    [Test]
    public void HMergeRow_FoldsContinueCellsIntoRestartColspan()
    {
        // Legacy w:hMerge (restart + continue continue) — the mechanism gridSpan does not cover.
        // The merged region must render as ONE wide cell, not content pinned to the first
        // narrow column with empty phantom cells to the right.
        var table = ThreeColTable(
            Row(Cell("H1"), Cell("H2"), Cell("H3")),
            Row(HMergeCell("MERGED", MergedCellValues.Restart),
                HMergeCell("", MergedCellValues.Continue),
                HMergeCell("", MergedCellValues.Continue)));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain("<td colspan=\"3\"");
        html.Should().Contain("MERGED");
        // 3 komórki nagłówka + 1 scalona — kontynuacje nie emitują <td>.
        System.Text.RegularExpressions.Regex.Matches(html, "<td").Count.Should().Be(4);
    }

    [Test]
    public void HMergeContinue_WithOmittedVal_IsTreatedAsContinue()
    {
        // Pominięty w:val na w:hMerge = "continue" (ECMA-376) — jak przy vMerge.
        var table = ThreeColTable(
            Row(Cell("H1"), Cell("H2"), Cell("H3")),
            Row(HMergeCell("AB", MergedCellValues.Restart), HMergeCell("", null), Cell("C")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain("<td colspan=\"2\"");
        html.Should().Contain("AB").And.Contain("C");
        System.Text.RegularExpressions.Regex.Matches(html, "<td").Count.Should().Be(5);
    }

    [Test]
    public void HMergeContinue_WithoutRestart_RendersNormally_NoContentLoss()
    {
        var table = ThreeColTable(
            Row(Cell("A"), HMergeCell("ORPHAN", MergedCellValues.Continue), Cell("C")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().Contain("ORPHAN");
        System.Text.RegularExpressions.Regex.Matches(html, "<td").Count.Should().Be(3);
    }

    private static TableCell HMergeCell(string text, MergedCellValues? val)
    {
        var hMerge = val == null ? new HorizontalMerge() : new HorizontalMerge { Val = val };
        return new TableCell(
            new TableCellProperties(hMerge),
            new Paragraph(new Run(new Text(text))));
    }

    // ---- Tabs inside table cells: no absolutely-positioned segments -------------------

    [Test]
    public void TabInsideTableCell_DoesNotUsePositionedSegments()
    {
        // Absolutne segmenty tabów (left:{stop}px na position:relative akapicie) w wąskiej
        // komórce wyjeżdżają poza komórkę i nakładają się na sąsiednią kolumnę (nagłówki
        // "Waluta"/"Termin spłaty" z dokumentu ING). W komórce tab renderuje się inline,
        // a stopy przeżywają w data-tab-stops (round-trip bez zmian).
        var cellPara = new Paragraph(
            new ParagraphProperties(new Tabs(new TabStop { Val = TabStopValues.Right, Position = 9000 })),
            new Run(new Text("Waluta")), new Run(new TabChar()), new Run(new Text("PLN")));
        var table = ThreeColTable(
            Row(new TableCell(new TableCellProperties(), cellPara), Cell("B"), Cell("C")));

        using var ms = Docx(table, new Paragraph());
        var html = _reader.Convert(ms).Html;

        html.Should().NotContain("docx-tab-seg");
        html.Should().NotContain("position:absolute");
        html.Should().Contain("data-tab-stops=\"9000:right\"");
        html.Should().Contain("Waluta").And.Contain("PLN");
    }

    private static Table ThreeColTable(params TableRow[] rows)
    {
        var t = new Table(
            new TableProperties(new TableStyle { Val = "TableGrid" }),
            new TableGrid(
                new GridColumn { Width = "3000" },
                new GridColumn { Width = "3000" },
                new GridColumn { Width = "3000" }));
        foreach (var r in rows) t.Append(r);
        return t;
    }

    private static TableRow Row(params TableCell[] cells) => new(cells.Cast<OpenXmlElement>().ToArray());

    private static TableCell Cell(string text, int gridSpan = 1)
    {
        var props = gridSpan > 1
            ? new TableCellProperties(new GridSpan { Val = gridSpan })
            : new TableCellProperties();
        return new TableCell(props, new Paragraph(new Run(new Text(text))));
    }

    // ---- Issue 5: paragraph spacing & manual line break -------------------------------

    [Test]
    public void ParagraphSpacing_MapsToDistinctCssProperties()
    {
        using var ms = Docx(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines
            { Before = "240", After = "120", Line = "360", LineRule = LineSpacingRuleValues.Auto }),
            new Run(new Text("l1")), new Run(new Break()), new Run(new Text("l2"))));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("margin-top:12pt;");     // 240 tw = 12 pt (odstęp PRZED)
        html.Should().Contain("margin-bottom:6pt;");   // 120 tw = 6 pt (odstęp PO)
        html.Should().Contain("line-height:1.5;");     // 360/240 = 1.5 (auto — interlinia)
        html.Should().Contain("<br/>");                // ręczne złamanie wiersza
    }

    // ---- Text boxes (wps:txbx / VML v:textbox) — no content loss ----------------------

    [Test]
    public void ModernTextBox_InAlternateContent_RendersTextOnce()
    {
        // wps:txbx (Choice) + VML v:textbox (Fallback) carry the SAME text; render exactly one.
        const string body = @"<w:p><w:r>
  <mc:AlternateContent>
    <mc:Choice Requires=""wps"">
      <w:drawing><wp:inline><wp:extent cx=""1000000"" cy=""500000""/>
        <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
          <wps:wsp><wps:txbx><w:txbxContent>
            <w:p><w:r><w:t>Tresc pola tekstowego</w:t></w:r></w:p>
          </w:txbxContent></wps:txbx></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing>
    </mc:Choice>
    <mc:Fallback>
      <w:pict><v:shape><v:textbox><w:txbxContent>
        <w:p><w:r><w:t>Tresc pola tekstowego</w:t></w:r></w:p>
      </w:txbxContent></v:textbox></v:shape></w:pict>
    </mc:Fallback>
  </mc:AlternateContent>
</w:r></w:p>";
        using var ms = DocxFromRawBody(body);

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("docx-textbox");
        html.Should().Contain("Tresc pola tekstowego");
        System.Text.RegularExpressions.Regex.Matches(html, "Tresc pola tekstowego").Count
            .Should().Be(1, "Choice i Fallback niosą tę samą treść — renderujemy raz");
    }

    [Test]
    public void AnchoredTextBox_IsAbsolutelyPositionedFromWordOffsets()
    {
        // Poziomo: page-relative 914400 EMU = 1 in = 96 px od lewej krawędzi strony.
        // Pionowo: obszar treści zaczyna się o górny margines (domyślnie 1 in) poniżej
        // krawędzi strony, więc page-relative 2 in (1828800) ląduje 1 in (96 px) niżej.
        const string body = @"<w:p><w:r>
  <w:drawing>
    <wp:anchor behindDoc=""0"" relativeHeight=""1"" allowOverlap=""1"" simplePos=""0""
      locked=""0"" layoutInCell=""1"">
      <wp:simplePos x=""0"" y=""0""/>
      <wp:positionH relativeFrom=""page""><wp:posOffset>914400</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom=""page""><wp:posOffset>1828800</wp:posOffset></wp:positionV>
      <wp:extent cx=""1828800"" cy=""457200""/>
      <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
        <wps:wsp><wps:txbx><w:txbxContent>
          <w:p><w:r><w:t>Zakotwiczone pole</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx></wps:wsp>
      </a:graphicData></a:graphic>
    </wp:anchor>
  </w:drawing>
</w:r></w:p>";
        using var ms = DocxFromRawBody(body);

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("docx-textbox").And.Contain("Zakotwiczone pole");
        html.Should().Contain("position:absolute;");
        html.Should().Contain("left:96px;").And.Contain("top:96px;");
        html.Should().Contain("width:192px;");   // 1828800 EMU = 192 px
    }

    [Test]
    public void LineShape_WithoutImage_RendersVisibleLine()
    {
        // Separator w stopce jako kształt-linia (straightConnector1) — wcześniej dropowany.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""5486400"" cy=""0""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:prstGeom prst=""straightConnector1""><a:avLst/></a:prstGeom>
        <a:ln w=""19050""><a:solidFill><a:srgbClr val=""FF6600""/></a:solidFill></a:ln>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        using var ms = DocxFromRawBody(body);

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("docx-line");
        html.Should().Contain("background:#FF6600;");
        html.Should().Contain("width:576px;");   // 5486400 EMU = 576 px
        html.Should().Contain("height:2px;");     // 19050 EMU ≈ 2 px
    }

    [Test]
    public void CustomGeometryShape_WithoutImage_RendersAsInlineSvgPath()
    {
        // Kształt DrawingML z własną ścieżką (a:custGeom) — np. wordmark „ING" / ikona „!".
        // Wcześniej dropowany w całości (RenderVectorShape zwracał ""), więc grafika z
        // oryginału NIE rysowała się w edytorze. Teraz → inline <svg><path> z kolorem kształtu.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""817245"" cy=""276860""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:custGeom><a:pathLst>
          <a:path w=""100"" h=""50"">
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo>
            <a:lnTo><a:pt x=""100"" y=""0""/></a:lnTo>
            <a:lnTo><a:pt x=""100"" y=""50""/></a:lnTo>
            <a:close/>
          </a:path>
        </a:pathLst></a:custGeom>
        <a:solidFill><a:srgbClr val=""000066""/></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        using var ms = DocxFromRawBody(body);

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("docx-custgeom");
        html.Should().Contain("<svg").And.Contain("viewBox=\"0 0 100 50\"");
        html.Should().Contain("<path d=\"M0 0 L100 0 L100 50 Z\"");
        html.Should().Contain("fill=\"#000066\"");
        // a:ln = noFill → brak obrysu (jak w Wordzie).
        html.Should().NotContain("stroke=");
    }

    [Test]
    public void VmlTextBox_Standalone_RendersText()
    {
        const string body = @"<w:p><w:r>
  <w:pict><v:shape><v:textbox><w:txbxContent>
    <w:p><w:r><w:t>Legacy textbox</w:t></w:r></w:p>
  </w:txbxContent></v:textbox></v:shape></w:pict>
</w:r></w:p>";
        using var ms = DocxFromRawBody(body);

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("docx-textbox").And.Contain("Legacy textbox");
    }

    private static MemoryStream DocxFromRawBody(string bodyInnerXml)
    {
        const string doc = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
  xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
  xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
  xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""
  xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""
  xmlns:v=""urn:schemas-microsoft-com:vml""
  xmlns:w10=""urn:schemas-microsoft-com:office:word"">
  <w:body>{BODY}</w:body>
</w:document>";

        var ms = new MemoryStream();
        using (var wpd = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wpd.AddMainDocumentPart();
            using var w = new StreamWriter(mainPart.GetStream(FileMode.Create), Encoding.UTF8);
            w.Write(doc.Replace("{BODY}", bodyInnerXml));
        }
        ms.Position = 0;
        return ms;
    }

    // ---- Issue 6: SVG data-URI display & sanitisation ---------------------------------

    [Test]
    public void ValidSvg_IsEmittedAsImageSvgXmlDataUri()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>"
                  + "<rect width='10' height='10' fill='red'/></svg>";
        using var ms = DocxWithSvg(svg, "image/svg+xml");

        var result = _reader.Convert(ms);

        result.Html.Should().Contain("src=\"data:image/svg+xml;base64,");
        result.Images.Should().ContainSingle().Which.ContentType.Should().Be("image/svg+xml");
    }

    [Test]
    public void MalformedSvgMimePrefix_IsNormalised()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect/></svg>";
        // Non-standard content type sometimes seen in source data: "img/svg+xml".
        using var ms = DocxWithSvg(svg, "img/svg+xml");

        var result = _reader.Convert(ms);

        result.Html.Should().Contain("data:image/svg+xml;base64,");
        result.Html.Should().NotContain("data:img/svg");
    }

    [Test]
    public void DangerousSvg_ScriptAndHandlersAreStripped()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg' onload='alert(1)' width='10' height='10'>"
                  + "<script>alert(2)</script><rect width='10' height='10' onclick='x()'/></svg>";
        using var ms = DocxWithSvg(svg, "image/svg+xml");

        var result = _reader.Convert(ms);

        var decoded = DecodeSingleImage(result);
        decoded.Should().NotContain("script");
        decoded.Should().NotContain("onload");
        decoded.Should().NotContain("onclick");
    }

    [Test]
    public void SvgLogoBuiltFromDefsAndUse_KeepsVisibleContentInDataUri()
    {
        // Regresja „puste białe logo": SVG złożone z <defs>+<use href='#id'> (typowy eksport
        // logo korporacyjnego) traciło wszystkie <use> w sanityzacji — w edytorze zostawał
        // obraz o poprawnych wymiarach bez żadnej widocznej treści.
        var svg =
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' viewBox='0 0 100 40'>"
            + "<defs><path id='lion' d='M10 10 C 20 0, 40 0, 50 10 Z' fill='#ff6200'/></defs>"
            + "<use xlink:href='#lion'/></svg>";
        using var ms = DocxWithSvg(svg, "image/svg+xml");

        var result = _reader.Convert(ms);

        result.Html.Should().Contain("data:image/svg+xml;base64,");
        var decoded = DecodeSingleImage(result);
        decoded.Should().Contain("<use").And.Contain("#lion");
    }

    [Test]
    public void SvgWithUtf8Bom_IsAcceptedAndEmbedded()
    {
        // Prefiks BOM w bajtach partu wywalał parser XML → cały poprawny SVG był pomijany,
        // a obraz znikał z edytora bez śladu.
        var svg = "\uFEFF<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>"
                  + "<rect width='10' height='10' fill='red'/></svg>";
        using var ms = DocxWithSvg(svg, "image/svg+xml");

        var result = _reader.Convert(ms);

        result.Images.Should().ContainSingle();
        result.Html.Should().Contain("data:image/svg+xml;base64,");
    }

    [Test]
    public void NonSvgDataInSvgPart_IsRejected_NotEmbedded()
    {
        using var ms = DocxWithSvg("this is not svg at all", "image/svg+xml");

        var result = _reader.Convert(ms);

        result.Images.Should().BeEmpty("niepoprawny SVG jest odrzucany, nie osadzany surowo");
    }

    private static string DecodeSingleImage(D2ViewerEditor.Domain.Models.DocumentContent content)
    {
        var img = content.Images.Single();
        return Encoding.UTF8.GetString(System.Convert.FromBase64String(img.Base64Data!));
    }

    private static MemoryStream DocxWithSvg(string svg, string contentType)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            var imagePart = mainPart.AddImagePart(contentType);
            using (var s = new MemoryStream(Encoding.UTF8.GetBytes(svg))) imagePart.FeedData(s);
            var rid = mainPart.GetIdOfPart(imagePart);
            body.Append(new Paragraph(new Run(BuildInlineDrawing(rid, 100000, 100000))));
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Drawing BuildInlineDrawing(string relationshipId, long cx, long cy)
        => new(new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = cx, Cy = cy },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1U, Name = "svg" },
            new DocumentFormat.OpenXml.Drawing.Graphic(
                new DocumentFormat.OpenXml.Drawing.GraphicData(
                    new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = 0U, Name = "svg" },
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                            new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId },
                            new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.Transform2D(
                                new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = cx, Cy = cy }),
                            new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
}

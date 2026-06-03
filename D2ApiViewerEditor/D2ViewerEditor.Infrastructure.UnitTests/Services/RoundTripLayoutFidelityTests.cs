using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Regression for the orginał_GOOD.docx → editor → zapisany_BAD.docx round-trip blow-up
/// (4 → 7 pages, enlarged tables). Tests the PIPELINE on synthetic documents that reproduce
/// the structures found in the analysis — not the specific file (no hardcoding, no committed
/// real document). Each test runs DOCX → DocumentContent (reader) → DOCX (writer) → DOCX and
/// asserts the authored layout survived.
///
/// Root causes verified here:
///  - body margins were inflated to Math.Max(top, headerHeight+720) → must stay as authored;
///  - header/footer distance was hardcoded 720 → must be reconstructed as (margin − band);
///  - table cells got 4px (≈60 twips) top/bottom padding → must be 0 (Word default);
///  - tables with no width were forced to pct 5000 (100%) → must stay auto.
/// </summary>
[TestFixture]
public class RoundTripLayoutFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private WordprocessingDocument RoundTrip(MemoryStream original)
    {
        var content = _reader.Convert(original);
        var bytes = _writer.Convert(content.Html, content.Metadata, content.Header, content.Footer,
            content.Margins, content.PageSize);
        var ms = new MemoryStream(bytes);
        return WordprocessingDocument.Open(ms, false);
    }

    private static PageMargin? FirstPageMargin(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()?.GetFirstChild<PageMargin>()
        ?? doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().FirstOrDefault()?.GetFirstChild<PageMargin>();

    // --- Margins ---------------------------------------------------------------

    [Test]
    public void Margins_NotInflated_StayCloseToAuthored()
    {
        // orginał_GOOD geometry: small top (567) + tiny header (6), small bottom (851) + footer (340).
        using var doc = RoundTrip(BuildDoc(top: 567, bottom: 851, header: 6, footer: 340, withHeaderFooter: true));
        var pgMar = FirstPageMargin(doc);

        pgMar.Should().NotBeNull();
        // BAD inflated top to 1281; the fix keeps it ~567 (±1 twip from cm rounding).
        ((int)pgMar!.Top!.Value).Should().BeInRange(560, 575);
        ((int)(uint)pgMar.Bottom!.Value).Should().BeInRange(845, 860);
    }

    [Test]
    public void HeaderFooterDistance_Reconstructed_NotHardcoded720()
    {
        using var doc = RoundTrip(BuildDoc(top: 567, bottom: 851, header: 6, footer: 340, withHeaderFooter: true));
        var pgMar = FirstPageMargin(doc)!;

        // Reconstructed as (margin − band). Was hardcoded 720 before → far smaller now.
        ((int)(uint)pgMar.Header!.Value).Should().BeLessThan(120);
        ((int)(uint)pgMar.Footer!.Value).Should().BeLessThan(450);
    }

    // --- Tables ----------------------------------------------------------------

    [Test]
    public void TableCellMargins_NoVerticalPadding_WordDefault()
    {
        using var doc = RoundTrip(BuildTableDoc(tableWidthAuto: true));
        var cells = doc.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToList();

        cells.Should().NotBeEmpty();
        foreach (var cell in cells)
        {
            var tcMar = cell.TableCellProperties?.TableCellMargin;
            // No top/bottom cell margin (was 60 twips ≈ 4px, which made every row taller).
            var top = tcMar?.TopMargin?.Width?.Value;
            var bottom = tcMar?.BottomMargin?.Width?.Value;
            if (top != null) int.Parse(top).Should().Be(0);
            if (bottom != null) int.Parse(bottom).Should().Be(0);
        }
    }

    [Test]
    public void TableWidth_AutoNotForcedToFullWidth()
    {
        using var doc = RoundTrip(BuildTableDoc(tableWidthAuto: true));
        var tblW = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().First()
            .GetFirstChild<TableProperties>()?.TableWidth;

        tblW.Should().NotBeNull();
        // BAD forced pct 5000 (100%). An auto-width table must stay auto, not pct.
        tblW!.Type!.Value.Should().NotBe(TableWidthUnitValues.Pct);
        tblW.Type.Value.Should().Be(TableWidthUnitValues.Auto);
    }

    // --- builders --------------------------------------------------------------

    private static MemoryStream BuildDoc(int top, int bottom, int header, int footer, bool withHeaderFooter)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            AppendStyles(mainPart);

            var sectPr = new SectionProperties();
            if (withHeaderFooter)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                headerPart.Header = new Header(new Paragraph(new Run(new Text("H"))));
                headerPart.Header.Save();
                var footerPart = mainPart.AddNewPart<FooterPart>();
                footerPart.Footer = new Footer(new Paragraph(new Run(new Text("F"))));
                footerPart.Footer.Save();
                sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });
                sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) });
            }

            body.Append(new Paragraph(new Run(new Text("Body"))));
            sectPr.Append(new DocumentFormat.OpenXml.Wordprocessing.PageSize { Width = 12240, Height = 15840 });
            sectPr.Append(new PageMargin
            {
                Top = top, Bottom = bottom, Left = 851, Right = 851,
                Header = (uint)header, Footer = (uint)footer, Gutter = 0
            });
            body.Append(sectPr);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static MemoryStream BuildTableDoc(bool tableWidthAuto)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            AppendStyles(mainPart);

            var tblPr = new TableProperties(
                new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "808080" }));

            // No tblCellMar, no per-cell tcMar — relies on Word default (top/bottom = 0).
            var table = new Table(
                tblPr,
                new TableGrid(new GridColumn { Width = "3402" }, new GridColumn { Width = "6981" }),
                new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth { Width = "3402", Type = TableWidthUnitValues.Dxa }), new Paragraph(new Run(new Text("A")))),
                    new TableCell(new TableCellProperties(new TableCellWidth { Width = "6981", Type = TableWidthUnitValues.Dxa }), new Paragraph(new Run(new Text("B"))))));
            body.Append(table);
            body.Append(new Paragraph()); // trailing paragraph Word requires after a table
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static void AppendStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
            new RunPropertiesBaseStyle(
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                new FontSize { Val = "22" }))));
        stylesPart.Styles.Save();
    }
}

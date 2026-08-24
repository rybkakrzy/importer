using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Fidelity of floating (wp:anchor) object positioning. The reader must resolve the
/// Word anchor into an editor-space offset that honours BOTH:
///   • <c>relativeFrom</c> (page / margin / column / …) — a margin/column offset is
///     measured from the text area, not the page edge (previously the raw posOffset was
///     used, so such objects landed one margin too far left / too high), and
///   • <c>wp:align</c> (right / center / …) — alignment with no explicit offset (the
///     letterhead pattern: logo aligned to the right margin) previously collapsed to 0.
///
/// Geometry: A4 (11906 × 16838 twips), 1440-twip (1") margins on every side.
///   page width  = 11906 twips = 7 560 310 EMU
///   left/right margin = 1440 twips = 914 400 EMU
///   content width = 7 560 310 − 914 400 − 914 400 = 5 731 510 EMU
/// </summary>
[TestFixture]
public class DocxAnchorPositionFidelityTests
{
    private DocxToHtmlConverter _converter = null!;

    [SetUp]
    public void Setup() => _converter = new DocxToHtmlConverter();

    private const long PageWidthEmu = 11906L * 635;   // 7 560 310
    private const long MarginEmu = 1440L * 635;        // 914 400
    private const long ContentWidthEmu = PageWidthEmu - 2 * MarginEmu; // 5 731 510

    private static readonly byte[] OnePixelPng = System.Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    /// <summary>
    /// Builds a one-paragraph A4 DOCX whose paragraph carries a floating (anchored) image.
    /// <paramref name="posHXml"/> / <paramref name="posVXml"/> are the full
    /// <c>wp:positionH</c> / <c>wp:positionV</c> elements so each test controls relativeFrom
    /// and offset-vs-align independently.
    /// </summary>
    private static MemoryStream BuildDocxWithAnchoredImage(string posHXml, string posVXml, long cx, long cy)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedBytes(OnePixelPng);
            var relId = mainPart.GetIdOfPart(imagePart);

            body.Append(new Paragraph(BuildAnchoredImageRun(relId, posHXml, posVXml, cx, cy)));

            body.Append(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720, Gutter = 0 }));

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Run BuildAnchoredImageRun(string relId, string posHXml, string posVXml, long cx, long cy)
    {
        var xml = $@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:drawing>
    <wp:anchor xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
               behindDoc=""0"" distT=""0"" distB=""0"" distL=""0"" distR=""0""
               simplePos=""0"" locked=""0"" layoutInCell=""1"" allowOverlap=""1"" relativeHeight=""1"">
      <wp:simplePos x=""0"" y=""0""/>
      {posHXml}
      {posVXml}
      <wp:extent cx=""{cx}"" cy=""{cy}""/>
      <wp:wrapNone/>
      <wp:docPr id=""2"" name=""Picture 2""/>
      <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
          <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
            <pic:nvPicPr><pic:cNvPr id=""2"" name=""Picture 2""/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill>
              <a:blip xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:embed=""{relId}""/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{cx}"" cy=""{cy}""/></a:xfrm>
              <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>
</w:r>";
        return new Run(xml);
    }

    private static string H(string relFrom, string inner) =>
        $@"<wp:positionH relativeFrom=""{relFrom}"">{inner}</wp:positionH>";
    private static string V(string relFrom, string inner) =>
        $@"<wp:positionV relativeFrom=""{relFrom}"">{inner}</wp:positionV>";
    private static string Off(long emu) => $"<wp:posOffset>{emu}</wp:posOffset>";
    private static string Align(string a) => $"<wp:align>{a}</wp:align>";

    private static long DataX(string html) => long.Parse(Regex.Match(html, "data-x-emu=\"(-?\\d+)\"").Groups[1].Value);
    private static long DataY(string html) => long.Parse(Regex.Match(html, "data-y-emu=\"(-?\\d+)\"").Groups[1].Value);

    // ── Horizontal relativeFrom ──────────────────────────────────────────────

    [Test]
    public void Horizontal_Page_Offset_MeasuredFromPageEdge()
    {
        using var stream = BuildDocxWithAnchoredImage(
            H("page", Off(1_000_000)), V("paragraph", Off(0)), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataX(html).Should().Be(1_000_000, "page-relative offset maps directly from the left page edge");
    }

    [Test]
    public void Horizontal_Margin_Offset_AddsLeftMargin()
    {
        // Same numeric posOffset as the page case, but measured from the LEFT MARGIN:
        // the resolved editor coordinate must be a full margin further right.
        using var stream = BuildDocxWithAnchoredImage(
            H("margin", Off(0)), V("paragraph", Off(0)), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataX(html).Should().Be(MarginEmu, "a margin/column offset starts at the text-area left edge, not the page edge");
    }

    [Test]
    public void Horizontal_Column_AlignRight_PinsToRightMargin()
    {
        // The classic letterhead: logo aligned to the right margin, no explicit offset.
        const long objW = 1_000_000;
        using var stream = BuildDocxWithAnchoredImage(
            H("column", Align("right")), V("paragraph", Off(0)), cx: objW, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        // base (left margin) + (content width − object width)
        DataX(html).Should().Be(MarginEmu + (ContentWidthEmu - objW));
    }

    [Test]
    public void Horizontal_Margin_AlignCenter_CentersInContentArea()
    {
        const long objW = 1_000_000;
        using var stream = BuildDocxWithAnchoredImage(
            H("margin", Align("center")), V("paragraph", Off(0)), cx: objW, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataX(html).Should().Be(MarginEmu + (ContentWidthEmu - objW) / 2);
    }

    [Test]
    public void Horizontal_Page_AlignRight_PinsToPageRightEdge()
    {
        const long objW = 1_000_000;
        using var stream = BuildDocxWithAnchoredImage(
            H("page", Align("right")), V("paragraph", Off(0)), cx: objW, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataX(html).Should().Be(PageWidthEmu - objW);
    }

    // ── Vertical relativeFrom (editor origin = top of the text area) ──────────

    [Test]
    public void Vertical_Paragraph_Offset_IsContentRelative()
    {
        using var stream = BuildDocxWithAnchoredImage(
            H("page", Off(0)), V("paragraph", Off(500_000)), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataY(html).Should().Be(500_000, "a paragraph/margin vertical offset is already relative to the content top");
    }

    [Test]
    public void Vertical_Page_Offset_SubtractsTopMargin()
    {
        // A page-relative vertical offset equal to the top margin sits exactly at the
        // content top → editor y = 0 (the body band begins one margin below the page top).
        using var stream = BuildDocxWithAnchoredImage(
            H("page", Off(0)), V("page", Off(MarginEmu)), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataY(html).Should().Be(0);
    }

    // ── Regression guard: offset still wins when both are absent → base, not crash ──

    [Test]
    public void Horizontal_NoOffsetNoAlign_FallsBackToReferenceBase()
    {
        using var stream = BuildDocxWithAnchoredImage(
            H("margin", ""), V("paragraph", ""), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataX(html).Should().Be(MarginEmu);
        DataY(html).Should().Be(0);
        html.Should().Contain("data-pos-mode=\"front\"");
    }

    // ── Anchors inside the header/footer band ────────────────────────────────
    // The reference paragraph of a band anchor lives in the BAND, not the body text
    // area: header paragraphs start at the header distance from the page top, footer
    // paragraphs at the bottom edge of the content area. Resolving them against the
    // body origin dropped a header logo one (margin − distance) too low, over the
    // first body lines.

    private const long HeaderDistanceEmu = 720L * 635;   // 457 200
    private const long PageHeightEmu = 16838L * 635;     // 10 692 130

    /// <summary>DOCX whose DEFAULT header (or footer) carries the anchored image.</summary>
    private static MemoryStream BuildDocxWithAnchoredImageInBand(
        string posVXml, long cx, long cy, bool footer = false)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            body.Append(new Paragraph(new Run(new Text("body"))));

            var sectPr = new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720, Gutter = 0 });

            var posH = H("page", Off(0));
            if (footer)
            {
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var imagePart = footerPart.AddImagePart(ImagePartType.Png);
                imagePart.FeedBytes(OnePixelPng);
                footerPart.Footer = new Footer(new Paragraph(
                    BuildAnchoredImageRun(footerPart.GetIdOfPart(imagePart), posH, posVXml, cx, cy)));
                sectPr.PrependChild(new FooterReference
                { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) });
            }
            else
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var imagePart = headerPart.AddImagePart(ImagePartType.Png);
                imagePart.FeedBytes(OnePixelPng);
                headerPart.Header = new Header(new Paragraph(
                    BuildAnchoredImageRun(headerPart.GetIdOfPart(imagePart), posH, posVXml, cx, cy)));
                sectPr.PrependChild(new HeaderReference
                { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });
            }

            body.Append(sectPr);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Header_Vertical_Paragraph_Offset_ResolvesAgainstHeaderBand()
    {
        // Logo anchored at the top of the header paragraph (the letterhead pattern).
        // Editor origin is the content top (= top margin), the header paragraph sits at
        // the header distance → contract y must be distance − margin (negative), NOT 0.
        using var stream = BuildDocxWithAnchoredImageInBand(V("paragraph", Off(0)), cx: 1_000_000, cy: 1_000_000);

        var headerHtml = _converter.Convert(stream).Header!.Html;

        DataY(headerHtml).Should().Be(HeaderDistanceEmu - MarginEmu);
    }

    [Test]
    public void Header_Vertical_Margin_Offset_KeepsPageLayoutOrigin()
    {
        // relativeFrom="margin" keeps the page's margin geometry even inside the band.
        using var stream = BuildDocxWithAnchoredImageInBand(V("margin", Off(0)), cx: 1_000_000, cy: 1_000_000);

        var headerHtml = _converter.Convert(stream).Header!.Html;

        DataY(headerHtml).Should().Be(0);
    }

    [Test]
    public void Footer_Vertical_Paragraph_Offset_ResolvesAgainstFooterBand()
    {
        // Footer paragraphs start at the bottom edge of the content area
        // (page height − bottom margin), mirroring the GUI band geometry.
        using var stream = BuildDocxWithAnchoredImageInBand(
            V("paragraph", Off(0)), cx: 1_000_000, cy: 300_000, footer: true);

        var footerHtml = _converter.Convert(stream).Footer!.Html;

        DataY(footerHtml).Should().Be(PageHeightEmu - MarginEmu - MarginEmu);
    }

    [Test]
    public void Body_Vertical_Paragraph_Offset_StaysContentRelative_AfterBandConversion()
    {
        // The band context must not leak into the body: converting a document whose
        // header holds an anchor leaves body anchors resolved against the content top.
        using var stream = BuildDocxWithAnchoredImage(
            H("page", Off(0)), V("paragraph", Off(500_000)), cx: 1_000_000, cy: 300_000);

        var html = _converter.Convert(stream).Html;

        DataY(html).Should().Be(500_000);
    }
}

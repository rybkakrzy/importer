using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Phase 1b fidelity tests — geometry, font family inheritance, image EMU sizing,
/// aspect ratio preservation, and paragraph alignment in footer (the typical
/// left/middle/right layout pattern). Structure mirrors the real reference DOCX
/// (`stupki.docx`) without committing its content.
/// </summary>
[TestFixture]
public class DocxToHtmlConverterFidelityTests
{
    private DocxToHtmlConverter _converter = null!;

    [SetUp]
    public void Setup() => _converter = new DocxToHtmlConverter();

    // 1 px PNG (red), base64; cheapest legal image that round-trips.
    private static readonly byte[] OnePixelPng = System.Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    /// <summary>
    /// Builds a DOCX shaped like the reference: A4 page, 1417-twip margins, header
    /// margin 708 (header band ≈ 1.25 cm), default header with an inline image of
    /// the given EMU size, default footer with three paragraphs aligned left/center/right.
    /// </summary>
    private static MemoryStream BuildReferenceDocx(long imageCx, long imageCy)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            // docDefaults — Calibri-like via theme is not available offline, so write an
            // explicit RunFonts (asciiTheme=minorHAnsi mimics stupki, but the converter
            // resolves themes via theme1.xml; for a deterministic test we set ascii=Calibri).
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                    new FontSize { Val = "24" }))));
            stylesPart.Styles.Save();

            // Default header: paragraph with an inline image (extent cx × cy EMU).
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var imagePart = headerPart.AddImagePart(ImagePartType.Png);
            using (var s = imagePart.GetStream()) s.Write(OnePixelPng, 0, OnePixelPng.Length);
            var imgRelId = headerPart.GetIdOfPart(imagePart);
            headerPart.Header = new Header(new Paragraph(BuildInlineImageRun(imgRelId, imageCx, imageCy)));
            headerPart.Header.Save();

            // Default footer: three paragraphs, left / center / right alignment.
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                AlignedParagraph(JustificationValues.Left, "L"),
                AlignedParagraph(JustificationValues.Center, "C"),
                AlignedParagraph(JustificationValues.Right, "R"));
            footerPart.Footer.Save();

            var sectPr = new SectionProperties();
            sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) });
            sectPr.Append(new PageSize { Width = 11906, Height = 16838 });
            sectPr.Append(new PageMargin
            {
                Top = 1417, Bottom = 1417, Left = 1417, Right = 1417,
                Header = 708, Footer = 708, Gutter = 0
            });
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Paragraph AlignedParagraph(JustificationValues jc, string text) =>
        new(new ParagraphProperties(new Justification { Val = jc }),
            new Run(new Text(text)));

    private static Run BuildInlineImageRun(string relId, long cx, long cy)
    {
        // Minimal DrawingML inline picture XML (the converter only inspects
        // a:blip/@r:embed and wp:extent/@cx/@cy, so the surrounding structure
        // just needs to parse as valid OOXML.)
        var xml = $@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:drawing>
    <wp:inline xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" distT=""0"" distB=""0"" distL=""0"" distR=""0"">
      <wp:extent cx=""{cx}"" cy=""{cy}""/>
      <wp:docPr id=""1"" name=""Picture 1""/>
      <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
          <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
            <pic:nvPicPr><pic:cNvPr id=""1"" name=""Picture 1""/><pic:cNvPicPr/></pic:nvPicPr>
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
    </wp:inline>
  </w:drawing>
</w:r>";
        return new Run(xml);
    }

    // --- Geometry ---

    [Test]
    public void Test01_Header_Height_ReflectsPageMarginMinusHeaderDistance()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        // (top - header) twips → cm. 709 / 1440 * 2.54 ≈ 1.250
        content.Header!.Height.Should().BeApproximately(1.25, 0.02);
    }

    [Test]
    public void Test02_Footer_Height_ReflectsBottomMarginMinusFooterDistance()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        content.Footer!.Height.Should().BeApproximately(1.25, 0.02);
    }

    [Test]
    public void Test02b_PageMargins_PreservedFromPgMar()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        // 1417 twips / 567 ≈ 2.50 cm on every side.
        content.Margins!.Top.Should().BeApproximately(2.5, 0.02);
        content.Margins.Bottom.Should().BeApproximately(2.5, 0.02);
        content.Margins.Left.Should().BeApproximately(2.5, 0.02);
        content.Margins.Right.Should().BeApproximately(2.5, 0.02);
    }

    // --- Font family inheritance (tests 7/8) ---

    [Test]
    public void Test07_Header_FontFamily_InheritedFromDocDefaults()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        // Container CSS for the header carries docDefault font-family.
        content.Header!.Html.Should().Contain("font-family:'Calibri'");
    }

    [Test]
    public void Test08_Footer_FontFamily_InheritedFromDocDefaults()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        content.Footer!.Html.Should().Contain("font-family:'Calibri'");
    }

    // --- Images (tests 9/10/11) ---

    [Test]
    public void Test09_Header_Image_RealisticSize_FromExtentEmu()
    {
        // EMU sizes from the real reference logo.
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        // EmuToPx = emu / 914400 * 96 → 133 × 34 px (1px rounding).
        content.Header!.Html.Should().MatchRegex(@"width:13[23]px;height:3[34]px");
    }

    [Test]
    public void Test11_Image_AspectRatio_Preserved_AcrossSizes()
    {
        // 4:1 EMU aspect.
        using var stream = BuildReferenceDocx(imageCx: 2_000_000, imageCy: 500_000);

        var content = _converter.Convert(stream);

        var w = ExtractInt(content.Header!.Html, @"width:(\d+)px");
        var h = ExtractInt(content.Header.Html, @"height:(\d+)px");
        w.Should().BeGreaterThan(0);
        h.Should().BeGreaterThan(0);
        // Aspect within ±2 px tolerance — pure integer rounding from EmuToPx.
        ((double)w / h).Should().BeApproximately(4.0, 0.06);
    }

    [Test]
    public void Test10_Image_OriginalEmu_EmittedAsDataAttributes_ForLosslessRoundTrip()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        // EMU is the canonical unit — keep it so the save-side can re-emit exact size.
        content.Header!.Html.Should().Contain("data-width-emu=\"1272540\"");
        content.Header.Html.Should().Contain("data-height-emu=\"327354\"");
    }

    // --- Footer "layout" (test 12) ---

    [Test]
    public void Test12_Footer_ThreeParagraphAlignment_PreservesLeftCenterRight()
    {
        using var stream = BuildReferenceDocx(imageCx: 1272540, imageCy: 327354);

        var content = _converter.Convert(stream);

        var html = content.Footer!.Html;
        // Order matters: L then C then R, each on its own paragraph with the matching align.
        var lIdx = html.IndexOf("text-align:left", System.StringComparison.OrdinalIgnoreCase);
        var cIdx = html.IndexOf("text-align:center", System.StringComparison.OrdinalIgnoreCase);
        var rIdx = html.IndexOf("text-align:right", System.StringComparison.OrdinalIgnoreCase);
        // At minimum centre and right must be emitted; left is the default and may be implicit.
        cIdx.Should().BeGreaterThan(-1, "centre paragraph alignment must survive import");
        rIdx.Should().BeGreaterThan(-1, "right paragraph alignment must survive import");
        cIdx.Should().BeLessThan(rIdx, "footer paragraph order must be preserved (L, C, R)");
        if (lIdx >= 0) lIdx.Should().BeLessThan(cIdx);
        // Text content is preserved in order regardless of how left alignment is encoded.
        html.IndexOf(">L<").Should().BeLessThan(html.IndexOf(">C<"));
        html.IndexOf(">C<").Should().BeLessThan(html.IndexOf(">R<"));
    }

    private static int ExtractInt(string s, string pattern)
    {
        var m = Regex.Match(s, pattern);
        return m.Success ? int.Parse(m.Groups[1].Value) : -1;
    }
}

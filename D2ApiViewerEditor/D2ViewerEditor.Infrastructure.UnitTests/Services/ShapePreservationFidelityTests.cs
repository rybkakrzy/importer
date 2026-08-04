using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Pass-through grafik XML (ADR-0056): kształty DrawingML (custGeom/prstGeom), VML,
/// OLE (w:object) i grupy (wpg:wgp) niosą oryginalny OOXML w data-docx-xml (+ części
/// relacji w data-docx-rels), a writer odtwarza je 1:1. Wcześniej były preview-only
/// albo dropowane — każdy autosave TRWALE usuwał je z wersji edytowalnej (v2).
/// </summary>
[TestFixture]
public class ShapePreservationFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    // 1x1 przezroczysty PNG (nagłówek podglądu OLE).
    private const string OnePxPngB64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream DocxFromRawBody(string bodyInnerXml)
    {
        const string doc = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
  xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
  xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
  xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""
  xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape""
  xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup""
  xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""
  xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
  xmlns:v=""urn:schemas-microsoft-com:vml""
  xmlns:o=""urn:schemas-microsoft-com:office:office""
  xmlns:w10=""urn:schemas-microsoft-com:office:word"">
  <w:body>{BODY}</w:body>
</w:document>";
        var ms = new MemoryStream();
        using (var docx = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = docx.AddMainDocumentPart();
            using var w = new StreamWriter(mainPart.GetStream(FileMode.Create), Encoding.UTF8);
            w.Write(doc.Replace("{BODY}", bodyInnerXml));
        }
        ms.Position = 0;
        return ms;
    }

    private const string CustGeomShapeBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:custGeom><a:pathLst>
          <a:path w=""100"" h=""50""><a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo>
            <a:lnTo><a:pt x=""100"" y=""0""/></a:lnTo><a:lnTo><a:pt x=""100"" y=""50""/></a:lnTo><a:close/></a:path>
        </a:pathLst></a:custGeom>
        <a:solidFill><a:srgbClr val=""000066""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    // ── Kształt custGeom: marker + round-trip ──────────────────────────────

    [Test]
    public void Read_CustomGeometryShape_CarriesPassThroughMarker()
    {
        var html = _reader.Convert(DocxFromRawBody(CustGeomShapeBody)).Html;

        html.Should().Contain("docx-custgeom");
        html.Should().Contain("data-docx-xml=\"", "oryginalny w:drawing musi jechać w pass-through");
        html.Should().Contain("contenteditable=\"false\"", "karetka w SVG zniszczyłaby podgląd");
    }

    [Test]
    public void RoundTrip_CustomGeometryShape_SurvivesSaveWithoutDuplication()
    {
        var html = _reader.Convert(DocxFromRawBody(CustGeomShapeBody)).Html;
        var docx = _writer.Convert(html);

        using (var doc = WordprocessingDocument.Open(new MemoryStream(docx), false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<A.CustomGeometry>().Should().HaveCount(1,
                "kształt nie może ani zniknąć (stan sprzed ADR-0056), ani się zduplikować");
            body.Descendants<A.SolidFill>()
                .SelectMany(f => f.Elements<A.RgbColorModelHex>())
                .Select(c => c.Val?.Value)
                .Should().Contain("000066", "wypełnienie oryginału musi wrócić 1:1");
        }

        // Drugi cykl (autosave po autosave) — nadal dokładnie jeden kształt.
        var html2 = _reader.Convert(new MemoryStream(docx)).Html;
        var docx2 = _writer.Convert(html2);
        using var doc2 = WordprocessingDocument.Open(new MemoryStream(docx2), false);
        doc2.MainDocumentPart!.Document!.Body!.Descendants<A.CustomGeometry>()
            .Should().HaveCount(1, "kolejne zapisy nie mogą mnożyć ani gubić kształtu");
    }

    // ── Nieznany preset: niewidoczny placeholder + round-trip ──────────────

    private const string UnknownPresetBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:prstGeom prst=""bentConnector3""><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val=""FF0000""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    [Test]
    public void Read_UnknownPreset_EmitsInvisiblePreservedPlaceholder()
    {
        var html = _reader.Convert(DocxFromRawBody(UnknownPresetBody)).Html;

        html.Should().Contain("docx-preserved", "dotąd kształt znikał bez śladu");
        html.Should().Contain("data-docx-xml=\"");
        // Placeholder zachowuje układ, ale nie maluje tuszu (reguła no-placeholder).
        html.Should().NotContain("background:#FF0000").And.NotContain("<svg");
    }

    [Test]
    public void RoundTrip_UnknownPreset_SurvivesSave()
    {
        var html = _reader.Convert(DocxFromRawBody(UnknownPresetBody)).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var preset = doc.MainDocumentPart!.Document!.Body!
            .Descendants<A.PresetGeometry>().Single();
        preset.Preset!.InnerText.Should().Be("bentConnector3");
    }

    // ── VML: hook translatora + round-trip w:pict ──────────────────────────

    private const string VmlOvalBody = @"<w:p><w:r>
  <w:pict><v:oval style=""width:120pt;height:40pt"" fillcolor=""#ff6100"" strokecolor=""#333333""/></w:pict>
</w:r></w:p>";

    [Test]
    public void Read_VmlOval_RendersSvgPreviewWithPassThrough()
    {
        var html = _reader.Convert(DocxFromRawBody(VmlOvalBody)).Html;

        html.Should().Contain("data-vml-shape=\"oval\"", "hook ConvertVmlShapeForEditor (GRAPHICS_CONVERSION §10)");
        html.Should().Contain("data:image/svg+xml", "podgląd SVG z serwisu grafik");
        html.Should().Contain("data-docx-xml=\"", "oryginalny w:pict do round-tripu");
    }

    [Test]
    public void RoundTrip_VmlOval_RestoresPict()
    {
        var html = _reader.Convert(DocxFromRawBody(VmlOvalBody)).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var oval = doc.MainDocumentPart!.Document!.Body!.Descendants<V.Oval>().Single();
        oval.FillColor!.Value.Should().Be("#ff6100");
        oval.Ancestors<Picture>().Should().NotBeEmpty("kształt musi wrócić jako w:pict");
    }

    [Test]
    public void Read_VmlShapeWithPath_PreservedInsteadOfSilentDrop()
    {
        // v:shape z własnym path — poza podzbiorem translatora; dotąd CICHY drop.
        const string body = @"<w:p><w:r>
  <w:pict><v:shape style=""width:60pt;height:60pt"" path=""m0,0l100,100e"" fillcolor=""#00ff00""/></w:pict>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("docx-preserved").And.Contain("data-docx-xml=\"");

        var docx = _writer.Convert(html);
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<V.Shape>()
            .Should().HaveCount(1, "w:pict/v:shape musi przetrwać zapis");
    }

    // ── OLE (w:object): podgląd + odtworzenie części relacji ───────────────

    private static MemoryStream DocxWithOleObject()
    {
        const string doc = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
  xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
  xmlns:v=""urn:schemas-microsoft-com:vml""
  xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <w:body><w:p><w:r>
    <w:object><v:shape style=""width:90pt;height:45pt""><v:imagedata r:id=""rIdPreview""/></v:shape>
      <o:OLEObject Type=""Embed"" ProgID=""Excel.Sheet.12"" r:id=""rIdOle""/></w:object>
  </w:r></w:p></w:body>
</w:document>";
        var ms = new MemoryStream();
        using (var docx = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = docx.AddMainDocumentPart();
            using (var w = new StreamWriter(mainPart.GetStream(FileMode.Create), Encoding.UTF8))
                w.Write(doc);

            var imagePart = mainPart.AddImagePart(ImagePartType.Png, "rIdPreview");
            using (var s = new MemoryStream(System.Convert.FromBase64String(OnePxPngB64)))
                imagePart.FeedData(s);

            var olePart = mainPart.AddNewPart<EmbeddedObjectPart>(
                "application/vnd.openxmlformats-officedocument.oleObject", "rIdOle");
            using (var s = new MemoryStream(Encoding.UTF8.GetBytes("fake-ole-binary")))
                olePart.FeedData(s);
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_OleObject_RendersImagePreviewWithRels()
    {
        var html = _reader.Convert(DocxWithOleObject()).Html;

        html.Should().Contain("data-ole-preview=\"1\"", "Word pokazuje statyczny podgląd v:imagedata");
        html.Should().Contain("data-docx-xml=\"");
        html.Should().Contain("data-docx-rels=\"", "binarium OLE + obraz podglądu muszą jechać z fragmentem");
    }

    [Test]
    public void RoundTrip_OleObject_RestoresObjectAndParts()
    {
        var html = _reader.Convert(DocxWithOleObject()).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var mainPart = doc.MainDocumentPart!;
        var obj = mainPart.Document!.Body!.Descendants<EmbeddedObject>().Single();

        // Wszystkie rId fragmentu muszą wskazywać ISTNIEJĄCE części nowego pakietu.
        const string relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var relIds = obj.Descendants()
            .SelectMany(e => e.GetAttributes())
            .Where(a => a.NamespaceUri == relNs && !string.IsNullOrEmpty(a.Value))
            .Select(a => a.Value!)
            .Distinct()
            .ToList();
        relIds.Should().HaveCount(2);
        foreach (var rid in relIds)
        {
            var part = mainPart.GetPartById(rid);
            part.Should().NotBeNull();
        }
        mainPart.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedObjectPart>()
            .Should().HaveCount(1, "binarium OLE musi zostać odtworzone");
    }

    // ── Grupa kształtów: render dzieci + round-trip ────────────────────────

    private const string GroupBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""1828800"" cy=""914400""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x=""0"" y=""0""/><a:ext cx=""1828800"" cy=""914400""/>
          <a:chOff x=""0"" y=""0""/><a:chExt cx=""3657600"" cy=""1828800""/>
        </a:xfrm></wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1828800"" cy=""1828800""/></a:xfrm>
          <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val=""FF6100""/></a:solidFill>
        </wps:spPr></wps:wsp>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""1828800"" y=""0""/><a:ext cx=""1828800"" cy=""1828800""/></a:xfrm>
          <a:prstGeom prst=""ellipse""><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val=""0000FF""/></a:solidFill>
        </wps:spPr></wps:wsp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    [Test]
    public void Read_GroupShape_RendersBothChildrenPositioned()
    {
        var html = _reader.Convert(DocxFromRawBody(GroupBody)).Html;

        html.Should().Contain("docx-group", "grupa nie może kolapsować do pierwszego dziecka");
        html.Should().Contain("data-docx-xml=\"");
        // Dwoje dzieci: rect (FF6100) i ellipse (0000FF), pozycjonowane transformacją chOff/chExt
        // (skala 0.5: dziecko 1828800 EMU → 914400 EMU → 96 px; drugie od 96 px).
        html.Should().Contain("#FF6100").And.Contain("#0000FF");
        html.Should().Contain("left:0px").And.Contain("left:96px");
    }

    [Test]
    public void RoundTrip_GroupShape_RestoresGroupWithBothChildren()
    {
        var html = _reader.Convert(DocxFromRawBody(GroupBody)).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var group = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Wpg.WordprocessingGroup>().Single();
        group.Elements<Wps.WordprocessingShape>().Should().HaveCount(2,
            "oba dzieci grupy muszą wrócić (dotąd zostawał najwyżej pierwszy blip)");
    }

    // ── Zagnieżdżone grupy / graphicFrame / grpFill ────────────────────────

    private const string NestedGroupBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""1828800"" cy=""914400""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x=""0"" y=""0""/><a:ext cx=""1828800"" cy=""914400""/>
          <a:chOff x=""0"" y=""0""/><a:chExt cx=""3657600"" cy=""1828800""/>
        </a:xfrm></wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""1828800"" cy=""1828800""/></a:xfrm>
          <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val=""FF6100""/></a:solidFill>
        </wps:spPr></wps:wsp>
        <wpg:grpSp>
          <wpg:grpSpPr><a:xfrm>
            <a:off x=""1828800"" y=""0""/><a:ext cx=""1828800"" cy=""1828800""/>
            <a:chOff x=""0"" y=""0""/><a:chExt cx=""914400"" cy=""914400""/>
          </a:xfrm></wpg:grpSpPr>
          <wps:wsp><wps:spPr>
            <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm>
            <a:prstGeom prst=""ellipse""><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val=""0000FF""/></a:solidFill>
          </wps:spPr></wps:wsp>
        </wpg:grpSp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    [Test]
    public void Read_NestedGroup_RendersInnerChildren()
    {
        // wpg:grpSp w wpg:wgp — dotąd CAŁA grupa degradowała do niewidocznego placeholdera.
        var html = _reader.Convert(DocxFromRawBody(NestedGroupBody)).Html;

        html.Should().Contain("docx-group", "grupa z pod-grupą musi się renderować, nie znikać");
        html.Should().Contain("#FF6100").And.Contain("#0000FF");
        // Pod-grupa: off 1828800 EMU × skala 0.5 = 96 px; skala składana 0.5 × ext/chExt(2.0) = 1.0
        // → elipsa 914400 EMU = 96 px.
        html.Should().Contain("left:96px").And.Contain("<ellipse");
    }

    [Test]
    public void Read_GroupWithGraphicFrame_DegradesToInvisiblePlaceholder()
    {
        // Wykres w grupie nie ma reprezentacji webowej — render częściowy (grupa bez ramki)
        // kłamałby wizualnie; cała grupa musi zdegradować do pass-through.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""1828800"" cy=""914400""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x=""0"" y=""0""/><a:ext cx=""1828800"" cy=""914400""/>
        </a:xfrm></wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm>
          <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val=""FF6100""/></a:solidFill>
        </wps:spPr></wps:wsp>
        <wpg:graphicFrame/>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("data-preserved=\"group\"");
        html.Should().NotContain("docx-group", "render częściowy grupy jest zabroniony");
    }

    // ── Dzieci sub-pikselowe / oś zerowa / placeholder pływający (ADR-0058; repro:
    // logo bankowe w stopce = grupa custGeom z detalami < 1px, separator extent cy=0) ──

    private const string SubPixelGroupBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""714515"" cy=""714528""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x=""0"" y=""0""/><a:ext cx=""714515"" cy=""714528""/>
          <a:chOff x=""0"" y=""0""/><a:chExt cx=""714515"" cy=""714528""/>
        </a:xfrm></wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""714515"" cy=""714528""/></a:xfrm>
          <a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w=""714515"" h=""714528"">
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo><a:lnTo><a:pt x=""714515"" y=""0""/></a:lnTo>
            <a:lnTo><a:pt x=""714515"" y=""714528""/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>
        </wps:spPr><wps:style>
          <a:lnRef idx=""0""><a:srgbClr val=""000000""><a:alpha val=""0""/></a:srgbClr></a:lnRef>
          <a:fillRef idx=""1""><a:srgbClr val=""FF6100""/></a:fillRef>
          <a:effectRef idx=""0""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:effectRef><a:fontRef idx=""none""/>
        </wps:style></wps:wsp>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""300000"" y=""300000""/><a:ext cx=""8000"" cy=""8000""/></a:xfrm>
          <a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w=""8000"" h=""8000"">
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo><a:lnTo><a:pt x=""8000"" y=""0""/></a:lnTo>
            <a:lnTo><a:pt x=""8000"" y=""8000""/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>
          <a:solidFill><a:srgbClr val=""FFFFFF""/></a:solidFill>
        </wps:spPr></wps:wsp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    [Test]
    public void Read_GroupWithSubPixelChild_RendersWholeGroup()
    {
        var html = _reader.Convert(DocxFromRawBody(SubPixelGroupBody)).Html;

        html.Should().Contain("docx-group", "sub-pikselowy detal gasił CAŁĄ grupę (logo w stopce)");
        html.Should().NotContain("data-preserved=\"group\"");
        html.Should().Contain("#FF6100", "wypełnienie z wps:style/a:fillRef musi się malować");
        html.Should().Contain("width:0.84px", "detal renderuje się w ułamkowych px, nie znika");
        // lnRef z alpha=0 (przezroczysty) = kształt BEZ obrysu — fallback lnRef nie może malować
        // czarnych konturów na sylwetce logo.
        html.Should().NotContain("stroke=\"#000000\"");
    }

    [Test]
    public void Read_ZeroHeightLine_LnRefOnlyStroke_RendersStroke()
    {
        // Realny wariant separatora stopki: a:ln deklaruje TYLKO grubość (bez solidFill),
        // kolor kreski siedzi w wps:style/a:lnRef — dotąd strokeHex=null → linia „bez tuszu"
        // była pomijana i cała grupa znikała.
        var body = ZeroHeightLineGroupBody
            .Replace(@"<a:ln w=""12700"" cap=""flat""><a:solidFill><a:srgbClr val=""FF5D00""/></a:solidFill></a:ln>",
                @"<a:ln w=""12700"" cap=""flat""/>")
            .Replace("</wps:spPr></wps:wsp>",
                @"</wps:spPr><wps:style>
  <a:lnRef idx=""1""><a:srgbClr val=""FF5D00""/></a:lnRef>
  <a:fillRef idx=""0""><a:srgbClr val=""FFFFFF""><a:alpha val=""0""/></a:srgbClr></a:fillRef>
  <a:effectRef idx=""0""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:effectRef><a:fontRef idx=""none""/>
</wps:style></wps:wsp>");
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("docx-group");
        html.Should().Contain("stroke=\"#FF5D00\"", "kolor kreski musi przyjść z wps:style/a:lnRef");
    }

    [Test]
    public void Read_TopLevelZeroHeightLine_RendersVisibleStroke()
    {
        // Linia jako pojedynczy kształt (bez grupy): extent cy=0 + obrys z lnRef — oś zerowa
        // dostaje grubość kreski zamiast gasić render (jak w dzieciach grup).
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""4571924"" cy=""0""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""4571924"" cy=""0""/></a:xfrm>
        <a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w=""4571924"" h=""0"">
          <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo><a:lnTo><a:pt x=""4571924"" y=""0""/></a:lnTo>
        </a:path></a:pathLst></a:custGeom>
        <a:ln w=""12700"" cap=""flat""/>
      </wps:spPr><wps:style>
        <a:lnRef idx=""1""><a:srgbClr val=""FF5D00""/></a:lnRef>
        <a:fillRef idx=""0""><a:srgbClr val=""FFFFFF""><a:alpha val=""0""/></a:srgbClr></a:fillRef>
        <a:effectRef idx=""0""><a:scrgbClr r=""0"" g=""0"" b=""0""/></a:effectRef><a:fontRef idx=""none""/>
      </wps:style></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("docx-custgeom", "linia o zerowej wysokości musi się renderować");
        html.Should().NotContain("data-preserved=");
        html.Should().Contain("stroke=\"#FF5D00\"");
        html.Should().Contain("vector-effect=\"non-scaling-stroke\"");
    }

    private const string ZeroHeightLineGroupBody = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""4571924"" cy=""0""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x=""0"" y=""0""/><a:ext cx=""4571924"" cy=""0""/>
          <a:chOff x=""0"" y=""0""/><a:chExt cx=""4571924"" cy=""0""/>
        </a:xfrm></wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""4571924"" cy=""0""/></a:xfrm>
          <a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w=""4571924"" h=""0"">
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo><a:lnTo><a:pt x=""4571924"" y=""0""/></a:lnTo>
          </a:path></a:pathLst></a:custGeom>
          <a:ln w=""12700"" cap=""flat""><a:solidFill><a:srgbClr val=""FF5D00""/></a:solidFill></a:ln>
        </wps:spPr></wps:wsp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";

    [Test]
    public void Read_ZeroHeightLineGroup_RendersVisibleStroke()
    {
        var html = _reader.Convert(DocxFromRawBody(ZeroHeightLineGroupBody)).Html;

        html.Should().Contain("docx-group", "grupa-linia o zerowej wysokości musi się renderować");
        html.Should().NotContain("data-preserved=\"group\"");
        html.Should().Contain("stroke=\"#FF5D00\"");
        html.Should().Contain("vector-effect=\"non-scaling-stroke\"",
            "stroke-width w px musi przetrwać viewBox w jednostkach EMU");
    }

    [Test]
    public void Read_FloatingUnrepresentableGraphic_PlaceholderHasZeroFootprint()
    {
        const string body = @"<w:p><w:r>
  <w:drawing><wp:anchor distT=""0"" distB=""0"" distL=""114300"" distR=""114300"" simplePos=""0""
      relativeHeight=""251658240"" behindDoc=""0"" locked=""0"" layoutInCell=""1"" allowOverlap=""1"">
    <wp:simplePos x=""0"" y=""0""/>
    <wp:positionH relativeFrom=""page""><wp:posOffset>6130290</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom=""page""><wp:posOffset>9729977</wp:posOffset></wp:positionV>
    <wp:extent cx=""714515"" cy=""714528""/><wp:wrapSquare wrapText=""bothSides""/>
    <wp:docPr id=""1"" name=""G""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""714515"" cy=""714528""/></a:xfrm></wpg:grpSpPr>
        <wpg:graphicFrame/>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:anchor></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("data-preserved=\"group\"", "graphicFrame w grupie nadal degraduje");
        var placeholder = System.Text.RegularExpressions.Regex
            .Match(html, "<span class=\"docx-preserved\"[^>]*>").Value;
        placeholder.Should().Contain("width:0;height:0;",
            "pływający placeholder nie może rezerwować miejsca w linii");
        placeholder.Should().NotContain("width:75px");
    }

    [Test]
    public void Read_InlineUnrepresentableGraphic_PlaceholderKeepsExtentSize()
    {
        // Kontrast do testu wyżej: grafika INLINE realnie zajmuje miejsce w linii.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""714515"" cy=""714528""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""714515"" cy=""714528""/></a:xfrm></wpg:grpSpPr>
        <wpg:graphicFrame/>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        var placeholder = System.Text.RegularExpressions.Regex
            .Match(html, "<span class=\"docx-preserved\"[^>]*>").Value;
        placeholder.Should().Contain("width:75px;height:75px;");
    }

    [Test]
    public void RoundTrip_SubPixelGroup_RestoresOriginalXml()
    {
        var html = _reader.Convert(DocxFromRawBody(SubPixelGroupBody)).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var group = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Wpg.WordprocessingGroup>().Single();
        group.Elements<Wps.WordprocessingShape>().Should().HaveCount(2,
            "oba dzieci (w tym sub-pikselowe) muszą wrócić z pass-through 1:1");
    }

    [Test]
    public void Read_GroupChildWithGrpFill_InheritsGroupFill()
    {
        // a:grpFill — dziecko maluje się wypełnieniem GRUPY; dotąd fill → currentColor.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""914400"" cy=""914400""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"">
      <wpg:wgp>
        <wpg:grpSpPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm>
          <a:solidFill><a:srgbClr val=""FF6100""/></a:solidFill>
        </wpg:grpSpPr>
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""914400"" cy=""914400""/></a:xfrm>
          <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
          <a:grpFill/>
        </wps:spPr></wps:wsp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("fill=\"#FF6100\"", "grpFill dziedziczy solidFill grupy");
        html.Should().NotContain("currentColor");
    }

    // ── mc:AlternateContent: Fallback VML zamiast niewidocznego placeholdera ──

    private const string AcChartWithVmlFallbackBody = @"<w:p><w:r>
  <mc:AlternateContent>
    <mc:Choice Requires=""wpg"">
      <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
        <a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
          <c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing>
    </mc:Choice>
    <mc:Fallback>
      <w:pict><v:oval style=""width:120pt;height:40pt"" fillcolor=""#ff6100""/></w:pict>
    </mc:Fallback>
  </mc:AlternateContent>
</w:r></w:p>";

    [Test]
    public void Read_AcUnrenderableChoice_PrefersVisibleVmlFallback()
    {
        // Dotąd niewidoczny placeholder z Choice KOŃCZYŁ przetwarzanie — użytkownik nie
        // widział nic, mimo że Fallback niósł renderowalną wersję VML.
        var html = _reader.Convert(DocxFromRawBody(AcChartWithVmlFallbackBody)).Html;

        html.Should().Contain("data-vml-shape=\"oval\"", "podgląd musi przyjść z gałęzi Fallback");

        // Marker pass-through musi nieść CAŁE mc:AlternateContent (obie gałęzie), nie sam w:pict.
        var m = System.Text.RegularExpressions.Regex.Match(html, "data-docx-xml=\"([^\"]+)\"");
        m.Success.Should().BeTrue();
        var decoded = Encoding.UTF8.GetString(System.Convert.FromBase64String(m.Groups[1].Value));
        decoded.Should().Contain("AlternateContent");
    }

    [Test]
    public void RoundTrip_AcUnrenderableChoice_RestoresBothBranches()
    {
        var html = _reader.Convert(DocxFromRawBody(AcChartWithVmlFallbackBody)).Html;
        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        body.Descendants<AlternateContent>().Should().HaveCount(1,
            "zapis nie może zdegradować dokumentu do samego w:pict");
        body.Descendants<V.Oval>().Should().HaveCount(1);
        body.Descendants<Drawing>().Should().HaveCount(1, "gałąź DrawingML (wykres) musi przetrwać");
    }

    // ── custGeom: guides, per-path fill, przestrzeń EMU ────────────────────

    [Test]
    public void Read_CustomGeometryWithGuides_EvaluatesFormulas()
    {
        // Współrzędne = NAZWY guide'ów z gdLst (fmla), ścieżka bez w/h = przestrzeń EMU —
        // dotąd oba przypadki kończyły się pustym renderem.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""952500""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:custGeom>
          <a:avLst/>
          <a:gdLst>
            <a:gd name=""halfW"" fmla=""*/ w 1 2""/>
            <a:gd name=""halfH"" fmla=""*/ h 1 2""/>
          </a:gdLst>
          <a:pathLst><a:path>
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo>
            <a:lnTo><a:pt x=""halfW"" y=""halfH""/></a:lnTo>
            <a:lnTo><a:pt x=""0"" y=""halfH""/></a:lnTo>
            <a:close/>
          </a:path></a:pathLst>
        </a:custGeom>
        <a:solidFill><a:srgbClr val=""336699""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("docx-custgeom");
        // w = 952500 EMU (100 px) → halfW = 476250.
        html.Should().Contain("476250").And.Contain("viewBox=\"0 0 952500 952500\"");
    }

    [Test]
    public void Read_CustomGeometryPathFillNone_KeepsSubpathUnfilled()
    {
        // a:path fill=""none"" = subścieżka bez tuszu — dotąd wszystkie subścieżki łączyły
        // się w jeden <path> ze wspólnym wypełnieniem.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""952500""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:custGeom><a:pathLst>
          <a:path w=""100"" h=""100"">
            <a:moveTo><a:pt x=""0"" y=""0""/></a:moveTo><a:lnTo><a:pt x=""100"" y=""100""/></a:lnTo><a:close/>
          </a:path>
          <a:path w=""100"" h=""100"" fill=""none"">
            <a:moveTo><a:pt x=""0"" y=""100""/></a:moveTo><a:lnTo><a:pt x=""100"" y=""0""/></a:lnTo>
          </a:path>
        </a:pathLst></a:custGeom>
        <a:solidFill><a:srgbClr val=""112233""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("fill=\"#112233\"").And.Contain("fill=\"none\"");
        System.Text.RegularExpressions.Regex.Matches(html, "<path ").Count.Should().Be(2);
    }

    // ── Transformacje koloru (lumMod itp.) ─────────────────────────────────

    [Test]
    public void Read_SrgbColorWithLumMod_AppliesLuminanceTransform()
    {
        // „Szary, ciemniejszy 50%" — dotąd transformacje odcienia były ignorowane
        // i kształt malował się kolorem bazowym.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:prstGeom prst=""triangle""><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val=""808080""><a:lumMod val=""50000""/></a:srgbClr></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("#404040", "lumMod 50% z #808080 = #404040");
        html.Should().NotContain("#808080");
    }

    // ── Tabela presetów / łuki / rotacja ───────────────────────────────────

    [Test]
    public void Read_PresetRightArrow_RendersPolygonSvg()
    {
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:prstGeom prst=""rightArrow""><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val=""008000""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("data-shape=\"rightArrow\"");
        html.Should().Contain("<polygon").And.Contain("fill=\"#008000\"");
        html.Should().Contain("data-docx-xml=\"", "podgląd ≠ źródło prawdy — oryginał w pass-through");
    }

    [Test]
    public void Read_CustomGeometryWithArc_EmitsSvgArcCommand()
    {
        // a:arcTo — dotąd pomijane (kształty z zaokrągleniami rozjeżdżały się).
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""952500""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:custGeom><a:pathLst>
          <a:path w=""100"" h=""100""><a:moveTo><a:pt x=""50"" y=""0""/></a:moveTo>
            <a:arcTo wR=""50"" hR=""50"" stAng=""16200000"" swAng=""10800000""/>
            <a:close/></a:path>
        </a:pathLst></a:custGeom>
        <a:solidFill><a:srgbClr val=""123456""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("docx-custgeom");
        html.Should().MatchRegex(@"d=""[^""]*A\s*50\b", "łuk musi wyemitować komendę SVG A z promieniami");
    }

    [Test]
    public void Read_RotatedShape_EmitsCssTransform()
    {
        // a:xfrm rot=2700000 (45°) — dotąd rotacja była ignorowana.
        const string body = @"<w:p><w:r>
  <w:drawing><wp:inline><wp:extent cx=""952500"" cy=""476250""/>
    <a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"">
      <wps:wsp><wps:spPr>
        <a:xfrm rot=""2700000""><a:off x=""0"" y=""0""/><a:ext cx=""952500"" cy=""476250""/></a:xfrm>
        <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val=""336699""/></a:solidFill>
      </wps:spPr></wps:wsp>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing>
</w:r></w:p>";
        var html = _reader.Convert(DocxFromRawBody(body)).Html;

        html.Should().Contain("transform:rotate(45deg)");
    }

    // ── Bezpieczeństwo / degradacja markera ────────────────────────────────

    [Test]
    public void Write_CorruptedMarker_DegradesWithoutException()
    {
        const string html = "<p><div class=\"docx-shape\" data-docx-xml=\"!!!nie-base64!!!\"></div>tekst</p>";

        var act = () => _writer.Convert(html);

        var docx = act.Should().NotThrow().Subject;
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>().Should().BeEmpty();
        doc.MainDocumentPart.Document.Body.InnerText.Should().Contain("tekst");
    }

    [Test]
    public void Write_NonWhitelistedRootInMarker_IsRejected()
    {
        // sectPr przemycone w markerze — root spoza białej listy (drawing/pict/object/AC)
        // nie może zostać wstrzyknięty do dokumentu.
        var evil = System.Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<w:sectPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"));
        var html = $"<p><span class=\"docx-preserved\" data-docx-xml=\"{evil}\"></span>ok</p>";

        var docx = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        // Dokładnie jeden sectPr — standardowy, na końcu body; żaden wstrzyknięty w akapit.
        doc.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>()
            .Should().HaveCount(1);
        doc.MainDocumentPart.Document.Body.Elements<Paragraph>()
            .SelectMany(p => p.Descendants<SectionProperties>())
            .Should().BeEmpty();
    }

    [Test]
    public void Write_MarkerWithDtd_IsRejected()
    {
        var evil = System.Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<!DOCTYPE foo [<!ENTITY xxe SYSTEM \"file:///etc/passwd\">]>" +
            "<w:drawing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">&xxe;</w:drawing>"));
        var html = $"<p><span class=\"docx-preserved\" data-docx-xml=\"{evil}\"></span>ok</p>";

        var act = () => _writer.Convert(html);

        var docx = act.Should().NotThrow().Subject;
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>().Should().BeEmpty();
    }
}

using System.Buffers.Binary;
using System.Linq;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// End-to-end: DOCX z obrazem EMF (a:blip → media part image/x-emf) musi trafić do edytora jako
/// renderowalna grafika (placeholder SVG albo osadzony raster), NIE jako `data:image/x-emf`,
/// którego przeglądarka nie wyświetli. To regresja realnego buga „grafiki EMF nie renderują się".
/// </summary>
[TestFixture]
public class GraphicConversionIntegrationTests
{
    [Test]
    public void DocxWithEmfImage_RendersAsRenderableDataUrl_NotRawEmf()
    {
        var docx = BuildDocxWithEmf();
        var reader = new DocxToHtmlConverter();

        var content = reader.Convert(new MemoryStream(docx));

        content.Html.Should().Contain("<img");
        // Renderowany `src` NIE może być surowym EMF (przeglądarka go nie wyświetli). Oryginalny
        // EMF wolno nieść w `data-original-src` (do round-tripu przez writer) — to nie jest `src`.
        content.Html.Should().NotContain(" src=\"data:image/x-emf");
        // Przezroczysty blank SVG (brak osadzonego rastra w syntetyku) oznaczony atrybutem diagnostycznym.
        content.Html.Should().Contain("data:image/svg+xml;base64,");
        content.Html.Should().Contain("data-legacy-graphic=\"blank\"");

        // Regresja: ZADEN widoczny placeholder w wygenerowanym HTML/SVG (req 1).
        AssertNoVisiblePlaceholder(content.Html);
    }

    /// <summary>
    /// Skanuje wygenerowane assety: dekoduje wszystkie data:image/svg+xml z HTML i sprawdza, że żaden
    /// nie zawiera szarego tła / ramki / tekstu „requires conversion"/„podgląd w Word". Blank fallback
    /// EMF/WMF musi być całkowicie przezroczysty (brak udawanej treści).
    /// </summary>
    private static void AssertNoVisiblePlaceholder(string html)
    {
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                     html, "data:image/svg\\+xml;base64,([A-Za-z0-9+/=]+)"))
        {
            var svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(m.Groups[1].Value));
            if (svg.Contains("data-legacy-graphic")) continue; // (attr nie trafia do treści SVG; defensywnie)
            svg.Should().NotContain("<text");
            foreach (var banned in new[] { "requires conversion", "podgląd w word", "konwersj" })
                svg.ToLowerInvariant().Should().NotContain(banned);
        }
        html.ToLowerInvariant().Should().NotContain("requires conversion");
    }

    /// <summary>
    /// Round-trip realnego buga „EMF psuje DOCX": DOCX z EMF → reader (placeholder + data-original-src)
    /// → writer. Wynikowy DOCX MUSI być poprawny (OpenXmlValidator: zero błędów = brak ostrzeżenia
    /// „dokument uszkodzony" w Word) i zawierać prawdziwy part EMF, NIE goły blip SVG.
    /// </summary>
    [Test]
    public void EmfRoundTrip_ProducesValidDocx_WithEmfPart_NotCorruptSvgBlip()
    {
        var docx = BuildDocxWithEmf();
        var reader = new DocxToHtmlConverter();
        var content = reader.Convert(new MemoryStream(docx));

        // Reader niesie oryginalny metafile, by writer mógł go odtworzyć.
        content.Html.Should().Contain("data-original-src=\"data:image/x-emf;base64,");

        var writer = new HtmlToDocxConverter();
        var outBytes = writer.Convert(content.Html);

        using var outMs = new MemoryStream(outBytes);
        using var outDoc = WordprocessingDocument.Open(outMs, false);

        // 1. Poprawność pakietu — to jest „nie uszkodzony w Word".
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        var descriptions = validator.Validate(outDoc)
            .Select(e => $"{e.Id} @ {e.Path?.XPath}: {e.Description}")
            .ToList();
        descriptions.Should().BeEmpty(because: "eksportowany DOCX nie może mieć błędów schematu (Word zgłasza uszkodzenie)");

        // 2. Zapisany obraz to prawdziwy EMF (Word renderuje natywnie), nie goły SVG blip.
        var imageParts = outDoc.MainDocumentPart!.ImageParts.ToList();
        imageParts.Should().Contain(p => p.ContentType == "image/x-emf");
        imageParts.Should().NotContain(p => p.ContentType == "image/svg+xml");
    }

    [Test]
    public void DocxWithWmfImage_RendersRenderableDataUrl_NotRawWmf_NoPlaceholder()
    {
        var docx = BuildDocxWithImages((ImagePartType.Wmf, BuildPlaceableWmf(0, 0, 1440, 720, 1440)));
        var content = new DocxToHtmlConverter().Convert(new MemoryStream(docx));

        content.Html.Should().Contain("<img");
        content.Html.Should().NotContain(" src=\"data:image/x-wmf");
        content.Html.Should().Contain("data-legacy-graphic=\"blank\"");
        AssertNoVisiblePlaceholder(content.Html);
    }

    [Test]
    public void DocxWithMixedNativeAndNonNativeImages_RendersBothRenderably()
    {
        var docx = BuildDocxWithImages(
            (ImagePartType.Png, MinimalPng(20, 20)),
            (ImagePartType.Emf, BuildEmf(10000, 5000)));
        var content = new DocxToHtmlConverter().Convert(new MemoryStream(docx));

        // Web-native PNG przechodzi bez zmian; EMF → przezroczysty blank (renderowalny), bez raw x-emf w src.
        content.Html.Should().Contain("data:image/png;base64,");
        content.Html.Should().Contain("data:image/svg+xml;base64,");
        content.Html.Should().NotContain(" src=\"data:image/x-emf");
        AssertNoVisiblePlaceholder(content.Html);
    }

    [Test]
    public void DocxWithDuplicatedEmf_RendersConsistently_AndDeduplicates()
    {
        var emf = BuildEmf(10000, 5000);
        var docx = BuildDocxWithImages((ImagePartType.Emf, emf), (ImagePartType.Emf, (byte[])emf.Clone()));
        var content = new DocxToHtmlConverter().Convert(new MemoryStream(docx));

        // Oba wystąpienia renderują się jako blank SVG (identyczna konwersja, deduplikowana w cache).
        System.Text.RegularExpressions.Regex.Matches(content.Html, "data-legacy-graphic=\"blank\"")
            .Count.Should().Be(2);
        content.Html.Should().NotContain(" src=\"data:image/x-emf");
        AssertNoVisiblePlaceholder(content.Html);
    }

    private static byte[] BuildDocxWithImages(params (PartTypeInfo type, byte[] data)[] images)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            uint id = 1;
            foreach (var (type, data) in images)
            {
                var part = main.AddImagePart(type);
                using (var s = new MemoryStream(data)) part.FeedData(s);
                var relId = main.GetIdOfPart(part);
                body.AppendChild(new Paragraph(new Run(BuildInlineDrawing(relId, id++))));
            }
            main.Document = new Document(body);
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static Drawing BuildInlineDrawing(string relId, uint id) =>
        new(new DW.Inline(
            new DW.Extent { Cx = 990000L, Cy = 495000L },
            new DW.DocProperties { Id = id, Name = $"img{id}" },
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"img{id}" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 990000L, Cy = 495000L }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

    private static byte[] BuildPlaceableWmf(short left, short top, short right, short bottom, ushort inch)
    {
        var d = new byte[40];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 0x9AC6CDD7);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(6), left);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(8), top);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(10), right);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(12), bottom);
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(14), inch);
        return d;
    }

    private static byte[] MinimalPng(int w, int h)
    {
        var bytes = new List<byte> { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        var ihdr = new byte[25];
        BinaryPrimitives.WriteUInt32BigEndian(ihdr, 13);
        System.Text.Encoding.ASCII.GetBytes("IHDR").CopyTo(ihdr, 4);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(8), (uint)w);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(12), (uint)h);
        ihdr[16] = 8; ihdr[17] = 2;
        bytes.AddRange(ihdr);
        var iend = new byte[12];
        System.Text.Encoding.ASCII.GetBytes("IEND").CopyTo(iend, 4);
        bytes.AddRange(iend);
        return bytes.ToArray();
    }

    private static byte[] BuildDocxWithEmf()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var emfPart = main.AddImagePart(ImagePartType.Emf);
            using (var s = new MemoryStream(BuildEmf(10000, 5000)))
                emfPart.FeedData(s);
            var relId = main.GetIdOfPart(emfPart);

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = 990000L, Cy = 495000L },
                    new DW.DocProperties { Id = 1U, Name = "emf" },
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = "emf" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = 990000L, Cy = 495000L }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            main.Document = new Document(new Body(new Paragraph(new Run(drawing))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildEmf(int frameRight, int frameBottom)
    {
        var d = new byte[88];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(4), 88);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(32), frameRight);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(36), frameBottom);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(40), 0x464D4520);
        return d;
    }
}

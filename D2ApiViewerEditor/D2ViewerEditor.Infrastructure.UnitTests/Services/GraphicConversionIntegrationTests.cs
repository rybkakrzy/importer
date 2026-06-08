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
        // Placeholder SVG (brak osadzonego rastra w syntetyku) oznaczony atrybutem diagnostycznym.
        content.Html.Should().Contain("data:image/svg+xml;base64,");
        content.Html.Should().Contain("data-legacy-graphic=\"placeholder\"");
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

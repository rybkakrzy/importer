using System.Buffers.Binary;
using System.Linq;
using System.Text;
using D2ViewerEditor.Domain.Models;
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
/// Etap 1 tłumacza wektorowego EMF/WMF → SVG (strategia `vector-translate` w
/// GraphicConversionService): czysto wektorowe metafile z obsługiwanym podzbiorem rekordów GDI
/// renderują się w edytorze jako realne SVG (kształty, kolory pióra/pędzla), a nie przezroczysty
/// blank. Oryginalny metafile nadal jedzie do DOCX przez pass-through (data-original-src).
/// </summary>
[TestFixture]
public class MetafileVectorTranslationTests
{
    // ---- EMF: kształty -----------------------------------------------------------------

    [Test]
    public void Emf_Rectangle_WithPenAndBrush_TranslatesToSvgRect()
    {
        var emf = EmfWith(
            Rec(38, 1, 0, 2, 0, 0x000000FF),   // CREATEPEN ih=1 solid 2px czerwony (COLORREF 0x00BBGGRR)
            Rec(37, 1),                         // SELECTOBJECT pen
            Rec(39, 2, 0, 0x0000FF00, 0),       // CREATEBRUSHINDIRECT ih=2 solid zielony
            Rec(37, 2),                         // SELECTOBJECT brush
            Rec(43, 10, 10, 110, 60));          // RECTANGLE l t r b

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        result.Web.Should().NotBeNull();
        result.Web!.MimeType.Should().Be("image/svg+xml");
        result.Web.IsBlankFallback.Should().BeFalse("wektor z obsługiwanymi rekordami nie może być blankiem");
        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Converted);
        result.Diagnostics.Fidelity.Should().Be(GraphicFidelity.Lossy);
        result.Diagnostics.AttemptedStrategies.Should().Contain("vector-translate");

        var svg = Encoding.UTF8.GetString(result.Web.Data);
        svg.Should().Contain("<rect");
        svg.Should().Contain("fill=\"#00ff00\"");
        svg.Should().Contain("stroke=\"#ff0000\"");
        svg.Should().Contain("stroke-width=\"2\"");
    }

    [Test]
    public void Emf_Polygon16_TranslatesToSvgPolygon()
    {
        var emf = EmfWith(Poly16Rec(86, 10, 10, 90, 10, 50, 80)); // POLYGON16, 3 punkty

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        result.Web!.MimeType.Should().Be("image/svg+xml");
        var svg = Encoding.UTF8.GetString(result.Web.Data);
        svg.Should().Contain("<polygon");
        svg.Should().Contain("10,10");
        svg.Should().Contain("50,80");
    }

    [Test]
    public void Emf_MoveToLineTo_TranslatesToSvgLine()
    {
        var emf = EmfWith(
            Rec(27, 5, 5),      // MOVETOEX
            Rec(54, 100, 50));  // LINETO

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        var svg = Encoding.UTF8.GetString(result.Web!.Data);
        svg.Should().Contain("<line");
        svg.Should().Contain("x2=\"100\"");
        svg.Should().Contain("y2=\"50\"");
    }

    [Test]
    public void Emf_Path_StrokePath_TranslatesToSvgPath()
    {
        var emf = EmfWith(
            Rec(59),             // BEGINPATH
            Rec(27, 10, 10),     // MOVETOEX
            Rec(54, 50, 10),     // LINETO
            Rec(54, 50, 40),     // LINETO
            Rec(61),             // CLOSEFIGURE
            Rec(60),             // ENDPATH
            Rec(64));            // STROKEPATH

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        var svg = Encoding.UTF8.GetString(result.Web!.Data);
        svg.Should().Contain("<path");
        svg.Should().Contain("M 10 10");
        svg.Should().Contain("L 50 40");
        svg.Should().Contain("Z");
        svg.Should().Contain("fill=\"none\"", "STROKEPATH rysuje tylko kontur");
    }

    [Test]
    public void Emf_SetWorldTransform_TranslationIsApplied()
    {
        var emf = EmfWith(
            XformRec(35, 1, 0, 0, 1, 5, 7),   // SETWORLDTRANSFORM: przesunięcie (5,7)
            Rec(43, 0, 0, 10, 10));            // RECTANGLE 0,0,10,10

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        var svg = Encoding.UTF8.GetString(result.Web!.Data);
        svg.Should().Contain("x=\"5\"");
        svg.Should().Contain("y=\"7\"");
    }

    // ---- regresje: kiedy tłumacz NIE może fabrykować treści ------------------------------

    [Test]
    public void Emf_HeaderOnly_NoDrawableRecords_StaysBlankFallback()
    {
        var emf = EmfHeaderOnly(10000, 5000);

        var result = new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        result.Web!.IsBlankFallback.Should().BeTrue("brak rekordów rysujących → nadal przezroczysty blank");
        result.Diagnostics.AttemptedStrategies.Should().Contain("vector-translate");
    }

    [Test]
    public void Emf_UnknownRecordsOnly_DoesNotThrow_FallsBackToBlank()
    {
        var emf = EmfWith(
            Rec(9999, 1, 2, 3),      // nieistniejący typ rekordu
            Rec(84, 0, 0, 0, 0));    // EXTTEXTOUTW — poza podzbiorem etapu 1

        var act = () => new GraphicConversionService().ConvertForEditor(EmfSource(emf));

        var result = act.Should().NotThrow().Subject;
        result.Web!.IsBlankFallback.Should().BeTrue();
    }

    // ---- WMF ------------------------------------------------------------------------------

    [Test]
    public void Wmf_Rectangle_WithPenAndBrush_TranslatesToSvgRect()
    {
        var wmf = WmfWith(
            WmfRec(0x02FA, 0, 1, 0, 0x00FF, 0),                      // CREATEPENINDIRECT czerwony
            WmfRec(0x012D, 0),                                        // SELECTOBJECT slot 0
            WmfRec(0x02FC, 0, unchecked((short)0xFF00), 0, 0),        // CREATEBRUSHINDIRECT zielony
            WmfRec(0x012D, 1),                                        // SELECTOBJECT slot 1
            WmfRec(0x041B, 60, 110, 10, 10));                         // RECTANGLE [B,R,T,L]

        var result = new GraphicConversionService().ConvertForEditor(new GraphicSource
        {
            Data = wmf,
            ContentType = "image/x-wmf"
        });

        result.Web!.MimeType.Should().Be("image/svg+xml");
        result.Web.IsBlankFallback.Should().BeFalse();
        result.Diagnostics.AttemptedStrategies.Should().Contain("vector-translate");

        var svg = Encoding.UTF8.GetString(result.Web.Data);
        svg.Should().Contain("<rect");
        svg.Should().Contain("stroke=\"#ff0000\"");
        svg.Should().Contain("fill=\"#00ff00\"");
    }

    [Test]
    public void Wmf_Polygon_TranslatesToSvgPolygon()
    {
        var wmf = WmfWith(WmfRec(0x0324, 3, 10, 10, 90, 10, 50, 80)); // POLYGON: count, x,y ×3

        var result = new GraphicConversionService().ConvertForEditor(new GraphicSource
        {
            Data = wmf,
            ContentType = "image/x-wmf"
        });

        var svg = Encoding.UTF8.GetString(result.Web!.Data);
        svg.Should().Contain("<polygon");
        svg.Should().Contain("90,10");
    }

    // ---- integracja: DOCX → edytor → DOCX ---------------------------------------------------

    [Test]
    public void DocxWithVectorEmf_RendersSvgInEditor_AndExportsOriginalEmf()
    {
        var emf = EmfWith(
            Rec(39, 1, 0, 0x00FF0000, 0),   // niebieski pędzel (COLORREF: BB=FF)
            Rec(37, 1),
            Rec(42, 20, 20, 120, 80));      // ELLIPSE

        var docx = BuildDocxWithEmf(emf);
        var content = new DocxToHtmlConverter().Convert(new MemoryStream(docx));

        // Podgląd: realne SVG w src, już NIE przezroczysty blank.
        content.Html.Should().Contain(" src=\"data:image/svg+xml;base64,");
        content.Html.Should().NotContain("data-legacy-graphic=\"blank\"");
        content.Html.Should().Contain("data-original-src=\"data:image/x-emf;base64,");

        var svgB64 = System.Text.RegularExpressions.Regex
            .Match(content.Html, "src=\"data:image/svg\\+xml;base64,([A-Za-z0-9+/=]+)\"").Groups[1].Value;
        var svg = Encoding.UTF8.GetString(Convert.FromBase64String(svgB64));
        svg.Should().Contain("<ellipse");
        svg.Should().Contain("fill=\"#0000ff\"");

        // Eksport: do DOCX wraca ORYGINALNY EMF (pass-through), nie SVG.
        var outBytes = new HtmlToDocxConverter().Convert(content.Html);
        using var outDoc = WordprocessingDocument.Open(new MemoryStream(outBytes), false);
        var parts = outDoc.MainDocumentPart!.ImageParts.ToList();
        parts.Should().Contain(p => p.ContentType == "image/x-emf");
        parts.Should().NotContain(p => p.ContentType == "image/svg+xml");
    }

    // ---- buildery EMF -----------------------------------------------------------------------

    private static GraphicSource EmfSource(byte[] emf) => new()
    {
        Data = emf,
        ContentType = "image/x-emf"
    };

    /// <summary>Nagłówek EMR_HEADER (88 B, bounds 0,0,200,100) + rekordy + EMR_EOF.</summary>
    private static byte[] EmfWith(params byte[][] records)
    {
        var total = 88 + records.Sum(r => r.Length) + 20;
        var d = new byte[total];
        W32(d, 0, 1);                       // iType = EMR_HEADER
        W32(d, 4, 88);                      // nSize nagłówka
        W32(d, 16, 200); W32(d, 20, 100);   // rclBounds right/bottom (left/top = 0)
        W32(d, 32, 5000); W32(d, 36, 2500); // rclFrame (0.01 mm)
        W32(d, 40, 0x464D4520);             // " EMF"
        int o = 88;
        foreach (var r in records) { r.CopyTo(d, o); o += r.Length; }
        W32(d, o, 14); W32(d, o + 4, 20); W32(d, o + 12, 16); W32(d, o + 16, 20); // EMR_EOF
        return d;
    }

    private static byte[] EmfHeaderOnly(int frameRight, int frameBottom)
    {
        var d = new byte[88];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(4), 88);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(32), frameRight);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(36), frameBottom);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(40), 0x464D4520);
        return d;
    }

    private static byte[] Rec(uint type, params uint[] dwords)
    {
        var d = new byte[8 + dwords.Length * 4];
        W32(d, 0, type);
        W32(d, 4, (uint)d.Length);
        for (int i = 0; i < dwords.Length; i++) W32(d, 8 + i * 4, dwords[i]);
        return d;
    }

    private static byte[] XformRec(uint type, float m11, float m12, float m21, float m22, float dx, float dy)
    {
        var d = new byte[32];
        W32(d, 0, type);
        W32(d, 4, 32);
        WF(d, 8, m11); WF(d, 12, m12); WF(d, 16, m21); WF(d, 20, m22); WF(d, 24, dx); WF(d, 28, dy);
        return d;
    }

    /// <summary>Rekord poly 16-bit: bounds(16B) + licznik + punkty s16 (x,y par).</summary>
    private static byte[] Poly16Rec(uint type, params short[] xy)
    {
        int n = xy.Length / 2;
        int size = 8 + 16 + 4 + n * 4;
        var d = new byte[size];
        W32(d, 0, type);
        W32(d, 4, (uint)size);
        W32(d, 24, (uint)n);
        for (int i = 0; i < xy.Length; i++)
            BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(28 + i * 2), xy[i]);
        return d;
    }

    // ---- buildery WMF -------------------------------------------------------------------------

    /// <summary>Placeable header (bbox 0,0,200,100 @96) + nagłówek WMF + rekordy + META_EOF.</summary>
    private static byte[] WmfWith(params byte[][] records)
    {
        var total = 22 + 18 + records.Sum(r => r.Length) + 6;
        var d = new byte[total];
        W32(d, 0, 0x9AC6CDD7);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(10), 200);  // right
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(12), 100);  // bottom
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(14), 96);  // inch
        int o = 22;
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(o), 1);      // type = memory metafile
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(o + 2), 9);  // headerSize w słowach
        o += 18;
        foreach (var r in records) { r.CopyTo(d, o); o += r.Length; }
        W32(d, o, 3); // META_EOF: rozmiar 3 słowa, funkcja 0x0000
        return d;
    }

    private static byte[] WmfRec(ushort func, params short[] parms)
    {
        int size = 6 + parms.Length * 2;
        var d = new byte[size];
        W32(d, 0, (uint)(size / 2));
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(4), func);
        for (int i = 0; i < parms.Length; i++)
            BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(6 + i * 2), parms[i]);
        return d;
    }

    // ---- DOCX builder ---------------------------------------------------------------------------

    private static byte[] BuildDocxWithEmf(byte[] emf)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var part = main.AddImagePart(ImagePartType.Emf);
            using (var s = new MemoryStream(emf)) part.FeedData(s);
            var relId = main.GetIdOfPart(part);

            var drawing = new Drawing(new DW.Inline(
                new DW.Extent { Cx = 990000L, Cy = 495000L },
                new DW.DocProperties { Id = 1U, Name = "emf" },
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "emf" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 495000L }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            main.Document = new Document(new Body(new Paragraph(new Run(drawing))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static void W32(byte[] d, int o, uint v) => BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(o), v);
    private static void WF(byte[] d, int o, float v) => BinaryPrimitives.WriteSingleLittleEndian(d.AsSpan(o), v);
}

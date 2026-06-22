using System.Buffers.Binary;
using System.Text;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Konwerter grafik (pure-managed, bez LibreOffice/GDI). Detekcja, EMF/WMF header → wymiary,
/// ekstrakcja osadzonego rastra, placeholder, VML shape → SVG, limity. Patrz .ai/GRAPHICS_CONVERSION.md.
/// </summary>
[TestFixture]
public class GraphicConversionServiceTests
{
    private GraphicConversionService _svc = null!;

    [SetUp]
    public void Setup() => _svc = new GraphicConversionService();

    // ---- detection ----------------------------------------------------------

    [Test]
    public void Detect_Png_Jpeg_Gif_Bmp_FromMagicBytes()
    {
        _svc.Detect(MinimalPng(10, 20)).Should().Be(GraphicKind.Png);
        _svc.Detect(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }).Should().Be(GraphicKind.Jpeg);
        _svc.Detect(Encoding.ASCII.GetBytes("GIF89a")).Should().Be(GraphicKind.Gif);
        _svc.Detect(new byte[] { (byte)'B', (byte)'M', 0, 0 }).Should().Be(GraphicKind.Bmp);
    }

    [Test]
    public void Detect_Emf_Wmf_Svg()
    {
        _svc.Detect(BuildEmf(10000, 5000)).Should().Be(GraphicKind.Emf);
        _svc.Detect(BuildPlaceableWmf(0, 0, 1440, 720, 1440)).Should().Be(GraphicKind.Wmf);
        _svc.Detect(Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'></svg>")).Should().Be(GraphicKind.Svg);
    }

    [Test]
    public void Detect_Tiff_Webp_Ico_FromMagicBytes()
    {
        _svc.Detect(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0, 0 }).Should().Be(GraphicKind.Tiff);   // little-endian TIFF
        _svc.Detect(new byte[] { 0x4D, 0x4D, 0x00, 0x2A, 0, 0 }).Should().Be(GraphicKind.Tiff);   // big-endian TIFF
        _svc.Detect(BuildWebpHeader()).Should().Be(GraphicKind.Webp);
        _svc.Detect(new byte[] { 0x00, 0x00, 0x01, 0x00, 0x01, 0x00 }).Should().Be(GraphicKind.Ico);
    }

    [Test]
    public void Detect_UnknownMedia_IsUnknown()
        => _svc.Detect(new byte[] { 0x00, 0x11, 0x22, 0x33, 0x44 }, "application/octet-stream").Should().Be(GraphicKind.Unknown);

    [Test]
    public void Detect_FallsBackToContentTypeHint_WhenMagicUnknown()
    {
        _svc.Detect(new byte[] { 1, 2, 3 }, "image/x-emf").Should().Be(GraphicKind.Emf);
        _svc.Detect(new byte[] { 1, 2, 3 }, "image/x-wmf").Should().Be(GraphicKind.Wmf);
        _svc.Detect(new byte[] { 1, 2, 3 }, "application/octet-stream").Should().Be(GraphicKind.Unknown);
    }

    // ---- raster passthrough -------------------------------------------------

    [Test]
    public void Png_IsPassedThrough_Lossless_WithDimensions()
    {
        var result = _svc.ConvertForEditor(new GraphicSource { Data = MinimalPng(64, 48) });

        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.PassThrough);
        result.Diagnostics.Fidelity.Should().Be(GraphicFidelity.Lossless);
        result.Web!.MimeType.Should().Be("image/png");
        result.Web.WidthPx.Should().Be(64);
        result.Web.HeightPx.Should().Be(48);
        result.Web.IsBlankFallback.Should().BeFalse();
    }

    // ---- EMF/WMF ------------------------------------------------------------

    [Test]
    public void Emf_WithoutEmbeddedRaster_GivesTransparentBlank_NotVisiblePlaceholder()
    {
        // rclFrame 10000 x 5000 (0.01mm) = 100mm x 50mm ≈ 378 x 189 px @96dpi.
        var result = _svc.ConvertForEditor(new GraphicSource { Data = BuildEmf(10000, 5000), ContentType = "image/x-emf" });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Emf);
        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Fallback);
        result.Diagnostics.Fidelity.Should().Be(GraphicFidelity.Fallback);
        result.Web!.MimeType.Should().Be("image/svg+xml");
        result.Web.IsBlankFallback.Should().BeTrue();
        result.Web.WidthPx.Should().Be(378);     // wymiary z layoutu zachowane (stabilność układu)
        result.Web.HeightPx.Should().Be(189);
        result.PreserveOriginalPart.Should().BeTrue(); // legacy part rides along to DOCX
        // Strukturalne raportowanie błędu (req 8): powód + lista prób strategii.
        result.Diagnostics.FailureReason.Should().NotBeNullOrEmpty();
        result.Diagnostics.AttemptedStrategies.Should().Contain("embedded-raster").And.Contain("dib-rasterize");

        AssertTransparentNoPlaceholder(result.Web.Data);
    }

    /// <summary>
    /// Regresja realnego buga: NIGDY szare tło / tekst „requires conversion" / „podgląd w Word"
    /// dla EMF/WMF bez rastra. Wynik to przezroczysty, pusty SVG (zero widocznej treści).
    /// </summary>
    [Test]
    public void EmfAndWmf_BlankFallback_ContainsNoPlaceholderTextOrFill()
    {
        var emf = _svc.ConvertForEditor(new GraphicSource { Data = BuildEmf(10000, 5000), ContentType = "image/x-emf" });
        var wmf = _svc.ConvertForEditor(new GraphicSource { Data = BuildPlaceableWmf(0, 0, 1440, 720, 1440), ContentType = "image/x-wmf" });

        AssertTransparentNoPlaceholder(emf.Web!.Data);
        AssertTransparentNoPlaceholder(wmf.Web!.Data);
    }

    private static void AssertTransparentNoPlaceholder(byte[] svgBytes)
    {
        var svg = Encoding.UTF8.GetString(svgBytes);
        svg.Should().StartWith("<svg").And.EndWith("</svg>");
        svg.Should().NotContain("<text");                 // brak jakiegokolwiek tekstu
        svg.Should().NotContain("<rect");                 // brak rysowanego tła/ramki
        svg.Should().NotContain("fill=");                 // brak wypełnienia (szare tło)
        svg.Should().NotContain("stroke");                // brak ramki
        foreach (var banned in new[] { "requires conversion", "podgląd w Word", "placeholder", "EMF —", "WMF —", "konwersj" })
            svg.ToLowerInvariant().Should().NotContain(banned.ToLowerInvariant());
    }

    [Test]
    public void Emf_WithEmbeddedPng_ExtractsRaster()
    {
        var emf = BuildEmf(10000, 5000);
        var png = MinimalPng(30, 40);
        var combined = emf.Concat(png).ToArray(); // EMF+ commonly wraps a PNG

        var result = _svc.ConvertForEditor(new GraphicSource { Data = combined, ContentType = "image/x-emf" });

        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Converted);
        result.Diagnostics.Fidelity.Should().Be(GraphicFidelity.Lossy);
        result.Web!.MimeType.Should().Be("image/png");
        result.Web.WidthPx.Should().Be(30);
        result.Web.HeightPx.Should().Be(40);
    }

    [Test]
    public void Emf_WithEmbeddedDib_IsRasterizedToPng()
    {
        // EMF z rekordem EMR_STRETCHDIBITS niosącym DIB 2x2 24bpp → realny PNG (nie placeholder).
        var emf = BuildEmfWithStretchDibits(2, 2);

        var result = _svc.ConvertForEditor(new GraphicSource { Data = emf, ContentType = "image/x-emf" });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Emf);
        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Converted);
        result.Web!.MimeType.Should().Be("image/png");
        result.Web.IsBlankFallback.Should().BeFalse();      // KLUCZOWE: realny raster, nie blank
        result.Web.WidthPx.Should().Be(2);
        result.Web.HeightPx.Should().Be(2);
        result.PreserveOriginalPart.Should().BeTrue();      // oryginał EMF nadal jedzie do DOCX (eksport wektorowy)
        result.Web.Data.AsSpan(0, 8).ToArray()
            .Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }); // sygnatura PNG
    }

    [Test]
    public void Wmf_WithEmbeddedDib_IsRasterizedToPng_ViaGenericScan()
    {
        // WMF (placeable) z dołączonym DIB-em → generyczny skan BITMAPINFOHEADER → PNG (nie placeholder).
        var wmf = BuildPlaceableWmf(0, 0, 96, 48, 96).Concat(BuildDib(2, 2)).ToArray();

        var result = _svc.ConvertForEditor(new GraphicSource { Data = wmf, ContentType = "image/x-wmf" });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Wmf);
        result.Web!.MimeType.Should().Be("image/png");
        result.Web.IsBlankFallback.Should().BeFalse();
    }

    [Test]
    public void Emf_WithCorruptDib_FallsBackToPlaceholder_WithoutThrowing()
    {
        // EMR_STRETCHDIBITS deklaruje DIB, ale nagłówek jest niepoprawny (biSize≠40) i offsety poza
        // rekordem → ani ścieżka rekordowa, ani skan generyczny nie znajdą rastra → placeholder, bez wyjątku.
        var emf = BuildEmfWithCorruptDibits();

        var result = _svc.ConvertForEditor(new GraphicSource { Data = emf, ContentType = "image/x-emf" });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Emf);
        result.Web!.IsBlankFallback.Should().BeTrue();
        result.Web.MimeType.Should().Be("image/svg+xml");
        result.PreserveOriginalPart.Should().BeTrue();
        AssertTransparentNoPlaceholder(result.Web.Data);
    }

    [Test]
    public void Wmf_Placeable_GivesTransparentBlank_WithDimensions()
    {
        // bbox 0,0,1440,720 with 1440 units/inch = 1in x 0.5in = 96 x 48 px.
        var result = _svc.ConvertForEditor(new GraphicSource { Data = BuildPlaceableWmf(0, 0, 1440, 720, 1440) });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Wmf);
        result.Web!.IsBlankFallback.Should().BeTrue();
        result.Web.WidthPx.Should().Be(96);
        result.Web.HeightPx.Should().Be(48);
        result.PreserveOriginalPart.Should().BeTrue();
        AssertTransparentNoPlaceholder(result.Web.Data);
    }

    // ---- VML shapes ---------------------------------------------------------

    [Test]
    public void Vml_Rect_RendersSvg()
    {
        var xml = "<v:rect xmlns:v='urn:schemas-microsoft-com:vml' style='width:100pt;height:50pt' fillcolor='#ff0000' strokecolor='#000000'/>";
        var result = _svc.ConvertVmlShapeForEditor(xml);

        result.Should().NotBeNull();
        var svg = Encoding.UTF8.GetString(result!.Web!.Data);
        svg.Should().Contain("<rect");
        svg.Should().Contain("#ff0000");
        result.Diagnostics.InputKind.Should().Be(GraphicKind.Vml);
    }

    [Test]
    public void Vml_Oval_And_Line_AreSupported()
    {
        var oval = _svc.ConvertVmlShapeForEditor("<v:oval xmlns:v='urn:schemas-microsoft-com:vml' style='width:40pt;height:40pt'/>");
        var line = _svc.ConvertVmlShapeForEditor("<v:line xmlns:v='urn:schemas-microsoft-com:vml' style='width:80pt;height:0pt'/>");
        Encoding.UTF8.GetString(oval!.Web!.Data).Should().Contain("<ellipse");
        Encoding.UTF8.GetString(line!.Web!.Data).Should().Contain("<line");
    }

    [Test]
    public void Vml_WithImageData_ReturnsNull_DeferringToPartConversion()
    {
        var xml = "<v:shape xmlns:v='urn:schemas-microsoft-com:vml' style='width:50pt;height:50pt'>" +
                  "<v:imagedata r:id='rId5' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/></v:shape>";
        _svc.ConvertVmlShapeForEditor(xml).Should().BeNull();
    }

    [Test]
    public void Vml_UnsupportedShape_ReturnsNull()
        => _svc.ConvertVmlShapeForEditor("<v:curve xmlns:v='urn:schemas-microsoft-com:vml' style='width:10pt;height:10pt'/>")
            .Should().BeNull();

    // ---- limits -------------------------------------------------------------

    [Test]
    public void EmptyInput_IsRejected()
    {
        var result = _svc.ConvertForEditor(new GraphicSource { Data = Array.Empty<byte>() });
        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Rejected);
        result.Web.Should().BeNull();
    }

    [Test]
    public void OversizeInput_IsRejected()
    {
        var opts = new GraphicConversionOptions { MaxInputBytes = 10 };
        var result = _svc.ConvertForEditor(new GraphicSource { Data = new byte[100] }, opts);
        result.Diagnostics.Status.Should().Be(GraphicConversionStatus.Rejected);
    }

    // ---- caching / deduplication --------------------------------------------

    [Test]
    public void IdenticalInput_IsDeduplicated_FromContentHashCache()
    {
        var emf = BuildEmf(10000, 5000);
        var first = _svc.ConvertForEditor(new GraphicSource { Data = (byte[])emf.Clone(), ContentType = "image/x-emf" });
        var second = _svc.ConvertForEditor(new GraphicSource { Data = (byte[])emf.Clone(), ContentType = "image/x-emf" });

        // Drugi przebieg trafia w cache po hashu treści → ten sam (zdeduplikowany) obiekt wyniku.
        ReferenceEquals(first, second).Should().BeTrue();
        ReferenceEquals(first.Web, second.Web).Should().BeTrue();
    }

    [Test]
    public void Diagnostics_ExposesCacheKey_AndAttemptedStrategies()
    {
        var result = _svc.ConvertForEditor(new GraphicSource
        {
            Data = MinimalPng(8, 8), SourcePath = "/word/media/image1.png"
        });

        result.Diagnostics.CacheKey.Should().NotBeNullOrEmpty();
        result.Diagnostics.AttemptedStrategies.Should().NotBeEmpty();
        result.Diagnostics.SourcePath.Should().Be("/word/media/image1.png");
    }

    [Test]
    public void DifferentContent_ProducesDifferentCacheKeys()
    {
        var a = _svc.ConvertForEditor(new GraphicSource { Data = MinimalPng(8, 8) });
        var b = _svc.ConvertForEditor(new GraphicSource { Data = MinimalPng(16, 16) });
        a.Diagnostics.CacheKey.Should().NotBe(b.Diagnostics.CacheKey);
    }

    // ---- TIFF (non-browser-native raster) -----------------------------------

    [Test]
    public void Tiff_WithoutDecodableData_GivesTransparentBlank_NotPlaceholder()
    {
        // Goły nagłówek TIFF bez ciała → Skia nie zdekoduje → przezroczysty blank (bez placeholdera),
        // oryginał zachowany do DOCX. NIGDY surowy image/tiff w `src`.
        var tiff = new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x08, 0, 0, 0 };
        var result = _svc.ConvertForEditor(new GraphicSource { Data = tiff, ContentType = "image/tiff" });

        result.Diagnostics.InputKind.Should().Be(GraphicKind.Tiff);
        result.Web!.MimeType.Should().Be("image/svg+xml");
        result.Web.IsBlankFallback.Should().BeTrue();
        result.PreserveOriginalPart.Should().BeTrue();
        result.Diagnostics.FailureReason.Should().NotBeNullOrEmpty();
        AssertTransparentNoPlaceholder(result.Web.Data);
    }

    // ---- synthetic graphic builders ----------------------------------------

    private static byte[] MinimalPng(int w, int h)
    {
        var bytes = new List<byte> { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // signature
        var ihdr = new byte[25]; // len(4)+"IHDR"(4)+data(13)+crc(4)
        BinaryPrimitives.WriteUInt32BigEndian(ihdr, 13);
        Encoding.ASCII.GetBytes("IHDR").CopyTo(ihdr, 4);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(8), (uint)w);   // offset 16 in file
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(12), (uint)h);  // offset 20 in file
        ihdr[16] = 8; ihdr[17] = 2; // bit depth / colour type
        bytes.AddRange(ihdr);
        var iend = new byte[12];
        Encoding.ASCII.GetBytes("IEND").CopyTo(iend, 4);
        bytes.AddRange(iend);
        return bytes.ToArray();
    }

    /// <summary>Minimalny nagłówek RIFF/WEBP (12 B) — wystarczy do detekcji po sygnaturze.</summary>
    private static byte[] BuildWebpHeader()
    {
        var d = new byte[12];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(d, 0);
        Encoding.ASCII.GetBytes("WEBP").CopyTo(d, 8);
        return d;
    }

    private static byte[] BuildEmf(int frameRight, int frameBottom)
    {
        var d = new byte[88];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 1);                       // iType = EMR_HEADER
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(4), 88);            // nSize
        // rclFrame (offset 24): 0,0,right,bottom in 0.01 mm
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(24), 0);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(28), 0);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(32), frameRight);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(36), frameBottom);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(40), 0x464D4520);   // " EMF"
        return d;
    }

    /// <summary>
    /// EMF: EMR_HEADER + EMR_STRETCHDIBITS (z osadzonym DIB w×h 24bpp BI_RGB) + EMR_EOF.
    /// Odzwierciedla realny przypadek Worda, w którym EMF opakowuje bitmapę.
    /// </summary>
    private static byte[] BuildEmfWithStretchDibits(int w, int h)
    {
        var header = BuildEmf(10000, 5000); // 88-bajtowy EMR_HEADER (Detect rozpozna EMF)

        int stride = ((w * 24 + 31) / 32) * 4;
        int cbBits = stride * h;
        const int cbBmi = 40;                 // BITMAPINFOHEADER, 24bpp → bez palety
        const int dibStart = 80;              // BitmapBuffer w EMR_STRETCHDIBITS
        int recSize = dibStart + cbBmi + cbBits;
        recSize = (recSize + 3) & ~3;         // wyrównanie do 4

        var rec = new byte[recSize];
        BinaryPrimitives.WriteUInt32LittleEndian(rec, 81);                 // iType = EMR_STRETCHDIBITS
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(4), (uint)recSize);
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(48), dibStart);          // offBmiSrc
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(52), cbBmi);             // cbBmiSrc
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(56), dibStart + cbBmi);  // offBitsSrc
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(60), (uint)cbBits);      // cbBitsSrc
        // BITMAPINFOHEADER
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(dibStart), 40);          // biSize
        BinaryPrimitives.WriteInt32LittleEndian(rec.AsSpan(dibStart + 4), w);        // biWidth
        BinaryPrimitives.WriteInt32LittleEndian(rec.AsSpan(dibStart + 8), h);        // biHeight
        BinaryPrimitives.WriteUInt16LittleEndian(rec.AsSpan(dibStart + 12), 1);      // biPlanes
        BinaryPrimitives.WriteUInt16LittleEndian(rec.AsSpan(dibStart + 14), 24);     // biBitCount
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(dibStart + 16), 0);      // BI_RGB
        // pixel bits — niezerowe (kolor), by dekoder dał sensowny obraz
        for (int i = dibStart + cbBmi; i < dibStart + cbBmi + cbBits; i++) rec[i] = 0x80;

        var eof = new byte[20];
        BinaryPrimitives.WriteUInt32LittleEndian(eof, 14);                 // iType = EMR_EOF
        BinaryPrimitives.WriteUInt32LittleEndian(eof.AsSpan(4), 20);       // nSize

        return header.Concat(rec).Concat(eof).ToArray();
    }

    /// <summary>Goły DIB (BITMAPINFOHEADER 40B, 24bpp BI_RGB + bity) — do testów skanu generycznego.</summary>
    private static byte[] BuildDib(int w, int h)
    {
        int stride = ((w * 24 + 31) / 32) * 4;
        int cbBits = stride * h;
        var dib = new byte[40 + cbBits];
        BinaryPrimitives.WriteUInt32LittleEndian(dib, 40);                 // biSize
        BinaryPrimitives.WriteInt32LittleEndian(dib.AsSpan(4), w);         // biWidth
        BinaryPrimitives.WriteInt32LittleEndian(dib.AsSpan(8), h);         // biHeight
        BinaryPrimitives.WriteUInt16LittleEndian(dib.AsSpan(12), 1);       // biPlanes
        BinaryPrimitives.WriteUInt16LittleEndian(dib.AsSpan(14), 24);      // biBitCount
        BinaryPrimitives.WriteUInt32LittleEndian(dib.AsSpan(16), 0);       // BI_RGB
        for (int i = 40; i < dib.Length; i++) dib[i] = 0x80;
        return dib;
    }

    /// <summary>EMF z rekordem EMR_STRETCHDIBITS o niepoprawnym DIB (biSize=0, offsety poza rekordem).</summary>
    private static byte[] BuildEmfWithCorruptDibits()
    {
        var header = BuildEmf(10000, 5000);
        const int recSize = 88; // wyrównane do 4
        var rec = new byte[recSize];
        BinaryPrimitives.WriteUInt32LittleEndian(rec, 81);                  // EMR_STRETCHDIBITS
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(4), recSize);
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(48), 9000);     // offBmiSrc poza rekordem
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(52), 40);       // cbBmiSrc
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(56), 9100);     // offBitsSrc poza rekordem
        BinaryPrimitives.WriteUInt32LittleEndian(rec.AsSpan(60), 16);       // cbBitsSrc
        return header.Concat(rec).ToArray();
    }

    private static byte[] BuildPlaceableWmf(short left, short top, short right, short bottom, ushort inch)
    {
        var d = new byte[40];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 0x9AC6CDD7);              // Aldus placeable magic
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(6), left);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(8), top);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(10), right);
        BinaryPrimitives.WriteInt16LittleEndian(d.AsSpan(12), bottom);
        BinaryPrimitives.WriteUInt16LittleEndian(d.AsSpan(14), inch);
        return d;
    }
}

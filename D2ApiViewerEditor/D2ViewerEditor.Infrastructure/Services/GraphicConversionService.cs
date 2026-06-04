using System.Buffers.Binary;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Pure-managed konwerter grafik (bez LibreOffice/GDI/System.Drawing → Linux/GCP-safe).
/// Patrz <see cref="IGraphicConversionService"/> i .ai/GRAPHICS_CONVERSION.md.
/// </summary>
public sealed class GraphicConversionService : IGraphicConversionService
{
    private const double EmuPerPixel = 9525.0;          // 914400 EMU/inch / 96 px/inch
    private const double EmfFrameUnitsPerMm = 100.0;    // rclFrame is in 0.01 mm
    private const double MmPerInch = 25.4;
    private const double DefaultDpi = 96.0;

    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    public GraphicKind Detect(ReadOnlySpan<byte> data, string? contentType = null)
    {
        // Magic bytes win over content-type (which can be spoofed / generic).
        if (data.Length >= 8 && data[..8].SequenceEqual(PngSignature)) return GraphicKind.Png;
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return GraphicKind.Jpeg;
        if (data.Length >= 6 && data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F') return GraphicKind.Gif;
        if (data.Length >= 2 && data[0] == (byte)'B' && data[1] == (byte)'M') return GraphicKind.Bmp;
        if (IsEmf(data)) return GraphicKind.Emf;
        if (IsWmf(data)) return GraphicKind.Wmf;
        if (LooksLikeSvg(data)) return GraphicKind.Svg;

        // Fall back to content-type hints for ambiguous/headerless cases.
        var ct = contentType?.ToLowerInvariant() ?? string.Empty;
        if (ct.Contains("emf")) return GraphicKind.Emf;
        if (ct.Contains("wmf") || ct.Contains("windows-metafile")) return GraphicKind.Wmf;
        if (ct.Contains("svg")) return GraphicKind.Svg;
        if (ct.Contains("png")) return GraphicKind.Png;
        if (ct.Contains("jpeg") || ct.Contains("jpg")) return GraphicKind.Jpeg;
        return GraphicKind.Unknown;
    }

    public GraphicConversionResult ConvertForEditor(
        GraphicSource source, GraphicConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= GraphicConversionOptions.Default;
        var sw = Stopwatch.StartNew();
        var warnings = new List<string>();
        var lost = new List<string>();

        if (source.Data == null || source.Data.Length == 0)
            return Rejected(GraphicKind.Unknown, sw, "Puste dane grafiki.");
        if (source.Data.Length > options.MaxInputBytes)
            return Rejected(GraphicKind.Unknown, sw, $"Przekroczono limit {options.MaxInputBytes} B.");

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(options.Timeout);

        var kind = Detect(source.Data, source.ContentType);
        try
        {
            switch (kind)
            {
                case GraphicKind.Png:
                case GraphicKind.Jpeg:
                case GraphicKind.Gif:
                case GraphicKind.Bmp:
                {
                    var (w, h) = ReadRasterSize(kind, source.Data);
                    return PassThroughWeb(kind, MimeFor(kind), source.Data, w, h, sw);
                }
                case GraphicKind.Svg:
                {
                    var svg = SanitizeSvg(Encoding.UTF8.GetString(source.Data));
                    if (svg == null) return Fallback(kind, source, options, sw, warnings, lost, "SVG nieparsowalny/niebezpieczny.");
                    var bytes = Encoding.UTF8.GetBytes(svg);
                    var (w, h) = ReadSvgSize(svg, source);
                    return new GraphicConversionResult
                    {
                        Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = bytes, WidthPx = w, HeightPx = h },
                        PreserveOriginalPart = false,
                        Diagnostics = Diag(kind, "image/svg+xml", GraphicConversionStatus.Converted, GraphicFidelity.Lossless, sw, warnings, lost)
                    };
                }
                case GraphicKind.Emf:
                case GraphicKind.Wmf:
                {
                    var (w, h) = kind == GraphicKind.Emf ? ReadEmfSize(source.Data) : ReadWmfSize(source.Data);
                    (w, h) = ResolveDims(w, h, source, options);

                    // Best-effort: many EMF/EMF+ wrap a PNG/JPEG — extract and show that raster.
                    var embedded = TryExtractEmbeddedRaster(source.Data);
                    if (embedded != null)
                    {
                        var ek = Detect(embedded);
                        var (ew, eh) = ReadRasterSize(ek, embedded);
                        warnings.Add($"{kind} nie jest rasteryzowane — pokazano osadzony {ek} z metafile.");
                        lost.Add("Wektorowe elementy metafile poza osadzonym rastrem.");
                        return new GraphicConversionResult
                        {
                            Web = new WebGraphicRepresentation { MimeType = MimeFor(ek), Data = embedded, WidthPx = ew > 0 ? ew : w, HeightPx = eh > 0 ? eh : h },
                            PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
                            Diagnostics = Diag(kind, MimeFor(ek), GraphicConversionStatus.Converted, GraphicFidelity.Lossy, sw, warnings, lost)
                        };
                    }

                    // No safe Linux-pure rasterizer for metafile → controlled placeholder; original
                    // part preserved so Word still renders the real graphic on save.
                    warnings.Add($"{kind} nie jest renderowane w przeglądarce — placeholder; oryginał zachowany w DOCX (pass-through).");
                    lost.Add("Podgląd wektorowy metafile (renderowany dopiero w Word).");
                    var placeholder = BuildPlaceholderSvg(kind, w, h);
                    return new GraphicConversionResult
                    {
                        Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = placeholder, WidthPx = w, HeightPx = h, IsPlaceholder = true },
                        PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
                        Diagnostics = Diag(kind, "image/svg+xml", GraphicConversionStatus.Fallback, GraphicFidelity.Fallback, sw, warnings, lost)
                    };
                }
                default:
                    return Fallback(kind, source, options, sw, warnings, lost, "Nieznany/nieobsługiwany format grafiki.");
            }
        }
        catch (OperationCanceledException)
        {
            return Rejected(kind, sw, "Przekroczono limit czasu konwersji.");
        }
        catch (Exception ex)
        {
            // Niezaufane wejście nie może wywrócić importu — mapujemy na fallback.
            return Fallback(kind, source, options, sw, warnings, lost, $"Błąd konwersji: {ex.GetType().Name}.");
        }
    }

    public GraphicConversionResult? ConvertVmlShapeForEditor(
        string vmlXml, GraphicConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= GraphicConversionOptions.Default;
        var sw = Stopwatch.StartNew();
        if (string.IsNullOrWhiteSpace(vmlXml)) return null;

        XElement root;
        try { root = SafeParse(vmlXml); }
        catch { return null; }

        var localName = root.Name.LocalName.ToLowerInvariant();
        // v:imagedata jest obsługiwane przez part (ConvertForEditor) — tu tylko kształty wektorowe.
        if (root.Descendants().Any(e => e.Name.LocalName.Equals("imagedata", StringComparison.OrdinalIgnoreCase)))
            return null;

        var style = root.Attribute("style")?.Value ?? string.Empty;
        var (w, h) = ParseVmlStyleSize(style);
        if (w <= 0) w = 200;
        if (h <= 0) h = 150;
        w = Math.Min(w, options.MaxPlaceholderWidthPx);
        h = Math.Min(h, options.MaxPlaceholderHeightPx);

        var fill = SafeColor(root.Attribute("fillcolor")?.Value) ?? "#cccccc";
        var stroke = SafeColor(root.Attribute("strokecolor")?.Value) ?? "#333333";
        var strokeWidthRaw = root.Attribute("strokeweight")?.Value;
        var strokeW = ParsePtToPx(strokeWidthRaw) is { } sp && sp > 0 ? sp : 1;

        string? shape = localName switch
        {
            "rect" => $"<rect x='0' y='0' width='{w}' height='{h}' fill='{fill}' stroke='{stroke}' stroke-width='{strokeW}'/>",
            "roundrect" => $"<rect x='0' y='0' width='{w}' height='{h}' rx='{Math.Max(4, w / 10)}' fill='{fill}' stroke='{stroke}' stroke-width='{strokeW}'/>",
            "oval" => $"<ellipse cx='{w / 2.0}' cy='{h / 2.0}' rx='{w / 2.0 - strokeW}' ry='{h / 2.0 - strokeW}' fill='{fill}' stroke='{stroke}' stroke-width='{strokeW}'/>",
            "line" => $"<line x1='0' y1='0' x2='{w}' y2='{h}' stroke='{stroke}' stroke-width='{strokeW}'/>",
            _ => null
        };
        if (shape == null) return null; // nieobsługiwany kształt → caller użyje placeholdera

        var svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' viewBox='0 0 {w} {h}'>{shape}</svg>";
        var bytes = Encoding.UTF8.GetBytes(svg);
        return new GraphicConversionResult
        {
            Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = bytes, WidthPx = w, HeightPx = h },
            PreserveOriginalPart = true, // VML zachowujemy w DOCX przez pass-through, póki nieedytowany
            Diagnostics = Diag(GraphicKind.Vml, "image/svg+xml", GraphicConversionStatus.Converted, GraphicFidelity.Lossy, sw,
                new[] { "Odwzorowano bezpieczny podzbiór VML (kształt + fill/stroke)." },
                new[] { "Gradienty/cienie/ścieżki/tekst VML." })
        };
    }

    public string? SanitizeSvg(string svgXml)
    {
        if (string.IsNullOrWhiteSpace(svgXml)) return null;
        XElement root;
        try { root = SafeParse(svgXml); }
        catch { return null; }
        if (!root.Name.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase)) return null;

        // Usuń niebezpieczne elementy.
        var killTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "script", "foreignObject", "iframe", "use", "set", "animate", "animateTransform", "handler" };
        root.Descendants().Where(e => killTags.Contains(e.Name.LocalName)).ToList().ForEach(e => e.Remove());

        foreach (var el in root.DescendantsAndSelf())
        {
            foreach (var attr in el.Attributes().ToList())
            {
                var name = attr.Name.LocalName.ToLowerInvariant();
                var val = attr.Value.Trim();
                if (name.StartsWith("on")) { attr.Remove(); continue; }              // on* handlers
                if (name is "href" or "xlink:href" or "src")
                {
                    // Tylko data: i bezpieczne fragmenty (#id). Blokuj http(s)/file/javascript.
                    if (!(val.StartsWith("data:", StringComparison.OrdinalIgnoreCase) || val.StartsWith("#")))
                        attr.Remove();
                    continue;
                }
                if (val.Contains("javascript:", StringComparison.OrdinalIgnoreCase)) attr.Remove();
            }
        }

        var result = root.ToString(SaveOptions.DisableFormatting);
        return string.IsNullOrWhiteSpace(result) ? null : result;
    }

    // ---- format detection helpers -------------------------------------------------

    private static bool IsEmf(ReadOnlySpan<byte> d)
    {
        // EMR_HEADER: iType==1 (offset 0) and dSignature==0x464D4520 (" EMF", offset 40).
        if (d.Length < 44) return false;
        return BinaryPrimitives.ReadUInt32LittleEndian(d) == 1
            && BinaryPrimitives.ReadUInt32LittleEndian(d[40..]) == 0x464D4520;
    }

    private static bool IsWmf(ReadOnlySpan<byte> d)
    {
        if (d.Length < 4) return false;
        var head = BinaryPrimitives.ReadUInt32LittleEndian(d);
        if (head == 0x9AC6CDD7) return true;                       // Aldus placeable
        // Standard METAFILEHEADER: type 1/2 + headerSize 9 words.
        var type = BinaryPrimitives.ReadUInt16LittleEndian(d);
        var headerWords = BinaryPrimitives.ReadUInt16LittleEndian(d[2..]);
        return (type == 1 || type == 2) && headerWords == 9;
    }

    private static bool LooksLikeSvg(ReadOnlySpan<byte> d)
    {
        var n = Math.Min(d.Length, 512);
        var head = Encoding.UTF8.GetString(d[..n]).TrimStart('﻿', ' ', '\t', '\r', '\n');
        return head.StartsWith("<?xml") && head.Contains("<svg", StringComparison.OrdinalIgnoreCase)
            || head.StartsWith("<svg", StringComparison.OrdinalIgnoreCase);
    }

    // ---- dimension parsers --------------------------------------------------------

    private static (int w, int h) ReadRasterSize(GraphicKind kind, byte[] d) => kind switch
    {
        GraphicKind.Png => ReadPngSize(d),
        GraphicKind.Gif => (BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(6)), BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(8))),
        GraphicKind.Bmp => d.Length >= 26 ? (BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(18)), Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(22)))) : (0, 0),
        GraphicKind.Jpeg => ReadJpegSize(d),
        _ => (0, 0)
    };

    private static (int w, int h) ReadPngSize(byte[] d) =>
        d.Length >= 24
            ? ((int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(16)), (int)BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(20)))
            : (0, 0);

    private static (int w, int h) ReadJpegSize(byte[] d)
    {
        int i = 2;
        while (i + 9 < d.Length)
        {
            if (d[i] != 0xFF) { i++; continue; }
            var marker = d[i + 1];
            // SOF0..SOF15 except DHT(C4), JPG(C8), DAC(CC) carry frame size.
            if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
            {
                int h = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(i + 5));
                int w = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(i + 7));
                return (w, h);
            }
            if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7)) { i += 2; continue; }
            int len = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(i + 2));
            i += 2 + len;
        }
        return (0, 0);
    }

    private static (int w, int h) ReadEmfSize(byte[] d)
    {
        if (d.Length < 40) return (0, 0);
        // rclFrame (offset 24): RECTL in 0.01 mm → px @96dpi.
        int left = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(24));
        int top = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(28));
        int right = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(32));
        int bottom = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(36));
        double wMm = Math.Abs(right - left) / EmfFrameUnitsPerMm;
        double hMm = Math.Abs(bottom - top) / EmfFrameUnitsPerMm;
        return ((int)Math.Round(wMm / MmPerInch * DefaultDpi), (int)Math.Round(hMm / MmPerInch * DefaultDpi));
    }

    private static (int w, int h) ReadWmfSize(byte[] d)
    {
        // Only the Aldus placeable header carries a bounding box.
        if (d.Length >= 22 && BinaryPrimitives.ReadUInt32LittleEndian(d) == 0x9AC6CDD7)
        {
            short left = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(6));
            short top = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(8));
            short right = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(10));
            short bottom = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(12));
            ushort inch = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(14));
            if (inch == 0) inch = 1440;
            double wIn = Math.Abs(right - left) / (double)inch;
            double hIn = Math.Abs(bottom - top) / (double)inch;
            return ((int)Math.Round(wIn * DefaultDpi), (int)Math.Round(hIn * DefaultDpi));
        }
        return (0, 0);
    }

    private static (int w, int h) ReadSvgSize(string svg, GraphicSource source)
    {
        var wm = Regex.Match(svg, @"\bwidth\s*=\s*[""']([\d.]+)", RegexOptions.IgnoreCase);
        var hm = Regex.Match(svg, @"\bheight\s*=\s*[""']([\d.]+)", RegexOptions.IgnoreCase);
        int w = wm.Success ? (int)double.Parse(wm.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
        int h = hm.Success ? (int)double.Parse(hm.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
        if (w <= 0 && source.TargetWidthEmu is > 0) w = (int)(source.TargetWidthEmu.Value / EmuPerPixel);
        if (h <= 0 && source.TargetHeightEmu is > 0) h = (int)(source.TargetHeightEmu.Value / EmuPerPixel);
        return (w, h);
    }

    private static (int w, int h) ResolveDims(int w, int h, GraphicSource source, GraphicConversionOptions o)
    {
        if (w <= 0 && source.TargetWidthEmu is > 0) w = (int)(source.TargetWidthEmu.Value / EmuPerPixel);
        if (h <= 0 && source.TargetHeightEmu is > 0) h = (int)(source.TargetHeightEmu.Value / EmuPerPixel);
        if (w <= 0) w = 200;
        if (h <= 0) h = 150;
        return (Math.Min(w, o.MaxPlaceholderWidthPx), Math.Min(h, o.MaxPlaceholderHeightPx));
    }

    // ---- embedded raster extraction ----------------------------------------------

    private static byte[]? TryExtractEmbeddedRaster(byte[] d)
    {
        int png = IndexOf(d, PngSignature, 0);
        if (png >= 0)
        {
            int iend = IndexOf(d, Encoding.ASCII.GetBytes("IEND"), png);
            if (iend > png && iend + 8 <= d.Length) return d[png..(iend + 8)];
        }
        // JPEG: SOI..EOI
        int soi = IndexOf(d, new byte[] { 0xFF, 0xD8, 0xFF }, 0);
        if (soi >= 0)
        {
            for (int i = soi + 2; i + 1 < d.Length; i++)
                if (d[i] == 0xFF && d[i + 1] == 0xD9) return d[soi..(i + 2)];
        }
        return null;
    }

    private static int IndexOf(byte[] haystack, byte[] needle, int start)
    {
        for (int i = Math.Max(0, start); i + needle.Length <= haystack.Length; i++)
        {
            bool ok = true;
            for (int j = 0; j < needle.Length; j++)
                if (haystack[i + j] != needle[j]) { ok = false; break; }
            if (ok) return i;
        }
        return -1;
    }

    // ---- placeholder + parsing helpers -------------------------------------------

    private static byte[] BuildPlaceholderSvg(GraphicKind kind, int w, int h)
    {
        var label = kind == GraphicKind.Emf ? "EMF" : kind == GraphicKind.Wmf ? "WMF" : "grafika";
        // Bez script/zewnętrznych zasobów. Tekst neutralny, ramka przerywana.
        var svg =
            $"<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' viewBox='0 0 {w} {h}'>" +
            $"<rect x='0.5' y='0.5' width='{w - 1}' height='{h - 1}' fill='#f3f4f6' stroke='#9ca3af' stroke-width='1' stroke-dasharray='6 4'/>" +
            $"<text x='{w / 2}' y='{h / 2}' font-family='sans-serif' font-size='14' fill='#6b7280' text-anchor='middle' dominant-baseline='middle'>{label} — podgląd w Word</text>" +
            "</svg>";
        return Encoding.UTF8.GetBytes(svg);
    }

    private static XElement SafeParse(string xml)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,   // blokuje DTD/encje (XXE/billion laughs)
            XmlResolver = null,                        // brak rozwiązywania zewnętrznych zasobów
            MaxCharactersFromEntities = 0,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true
        };
        using var sr = new StringReader(xml);
        using var reader = XmlReader.Create(sr, settings);
        return XElement.Load(reader);
    }

    private static (int w, int h) ParseVmlStyleSize(string style)
    {
        int w = 0, h = 0;
        var wm = Regex.Match(style, @"width:\s*([\d.]+)pt", RegexOptions.IgnoreCase);
        var hm = Regex.Match(style, @"height:\s*([\d.]+)pt", RegexOptions.IgnoreCase);
        if (wm.Success) w = (int)(double.Parse(wm.Groups[1].Value, CultureInfo.InvariantCulture) * DefaultDpi / 72.0);
        if (hm.Success) h = (int)(double.Parse(hm.Groups[1].Value, CultureInfo.InvariantCulture) * DefaultDpi / 72.0);
        return (w, h);
    }

    private static double? ParsePtToPx(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var m = Regex.Match(raw, @"([\d.]+)");
        if (!m.Success) return null;
        var pt = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        return raw.Contains("pt", StringComparison.OrdinalIgnoreCase) ? pt * DefaultDpi / 72.0 : pt;
    }

    private static string? SafeColor(string? c)
    {
        if (string.IsNullOrWhiteSpace(c)) return null;
        c = c.Trim();
        if (Regex.IsMatch(c, "^#[0-9A-Fa-f]{3,6}$")) return c;
        if (Regex.IsMatch(c, "^[a-zA-Z]{3,20}$")) return c.ToLowerInvariant(); // named color
        return null;
    }

    private static string MimeFor(GraphicKind k) => k switch
    {
        GraphicKind.Png => "image/png",
        GraphicKind.Jpeg => "image/jpeg",
        GraphicKind.Gif => "image/gif",
        GraphicKind.Bmp => "image/bmp",
        GraphicKind.Svg => "image/svg+xml",
        _ => "application/octet-stream"
    };

    // ---- result factories ---------------------------------------------------------

    private static GraphicConversionResult PassThroughWeb(GraphicKind kind, string mime, byte[] data, int w, int h, Stopwatch sw) =>
        new()
        {
            Web = new WebGraphicRepresentation { MimeType = mime, Data = data, WidthPx = w, HeightPx = h },
            PreserveOriginalPart = false,
            Diagnostics = Diag(kind, mime, GraphicConversionStatus.PassThrough, GraphicFidelity.Lossless, sw,
                Array.Empty<string>(), Array.Empty<string>())
        };

    private GraphicConversionResult Fallback(GraphicKind kind, GraphicSource source, GraphicConversionOptions o,
        Stopwatch sw, List<string> warnings, List<string> lost, string warning)
    {
        warnings.Add(warning);
        var (w, h) = ResolveDims(0, 0, source, o);
        var placeholder = BuildPlaceholderSvg(kind, w, h);
        return new GraphicConversionResult
        {
            Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = placeholder, WidthPx = w, HeightPx = h, IsPlaceholder = true },
            PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
            Diagnostics = Diag(kind, "image/svg+xml", GraphicConversionStatus.Unsupported, GraphicFidelity.Unsupported, sw, warnings, lost)
        };
    }

    private static GraphicConversionResult Rejected(GraphicKind kind, Stopwatch sw, string reason) =>
        new()
        {
            Web = null,
            PreserveOriginalPart = false,
            Diagnostics = Diag(kind, string.Empty, GraphicConversionStatus.Rejected, GraphicFidelity.Unsupported, sw,
                new[] { reason }, Array.Empty<string>())
        };

    private static GraphicConversionDiagnostics Diag(GraphicKind kind, string mime, GraphicConversionStatus status,
        GraphicFidelity fidelity, Stopwatch sw, IReadOnlyList<string> warnings, IReadOnlyList<string> lost)
    {
        sw.Stop();
        return new GraphicConversionDiagnostics
        {
            InputKind = kind,
            OutputMimeType = mime,
            Status = status,
            Fidelity = fidelity,
            ElapsedMs = sw.ElapsedMilliseconds,
            Warnings = warnings,
            LostProperties = lost
        };
    }
}

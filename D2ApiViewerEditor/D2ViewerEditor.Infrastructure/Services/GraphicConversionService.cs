using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using SkiaSharp;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Pure-managed konwerter grafik dokumentowych (bez LibreOffice/GDI/System.Drawing → Linux/GCP-safe).
///
/// Pipeline: <c>Detect</c> (klasyfikacja) → <c>ConvertForEditor</c> (wybór strategii + łańcuch
/// fallbacków + cache po hashu treści). Strategie produkują wyłącznie formaty renderowalne w
/// przeglądarce (PNG/SVG/web-native raster). Gdy żadna realna strategia nie zwróci rastra dla
/// metafile czysto wektorowego, zwracamy PRZEZROCZYSTĄ grafikę zachowującą układ — NIGDY widoczny
/// placeholder ("requires conversion"/szare tło). Oryginalny part jedzie do DOCX (pass-through),
/// więc Word renderuje prawdziwy wektor. Patrz <see cref="IGraphicConversionService"/> i
/// .ai/GRAPHICS_CONVERSION.md.
/// </summary>
public sealed class GraphicConversionService : IGraphicConversionService
{
    private const double EmuPerPixel = 9525.0;          // 914400 EMU/inch / 96 px/inch
    private const double EmfFrameUnitsPerMm = 100.0;    // rclFrame is in 0.01 mm
    private const double MmPerInch = 25.4;
    private const double DefaultDpi = 96.0;
    private const int MaxCacheEntries = 1024;           // ograniczenie pamięci cache (bounded)

    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    // Cache deduplikujący identyczne assety w obrębie instancji (np. ten sam obraz wstawiony N razy
    // w dokumencie). Klucz = hash treści + parametry wpływające na wynik. ConcurrentDictionary →
    // bez globalnego locka. Wynik konwersji jest niezmienny (immutable), więc współdzielenie bezpieczne.
    private readonly ConcurrentDictionary<string, GraphicConversionResult> _cache = new();

    public GraphicKind Detect(ReadOnlySpan<byte> data, string? contentType = null)
    {
        // Magic bytes win over content-type (which can be spoofed / generic).
        if (data.Length >= 8 && data[..8].SequenceEqual(PngSignature)) return GraphicKind.Png;
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return GraphicKind.Jpeg;
        if (data.Length >= 6 && data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F') return GraphicKind.Gif;
        if (IsWebp(data)) return GraphicKind.Webp;
        if (data.Length >= 2 && data[0] == (byte)'B' && data[1] == (byte)'M') return GraphicKind.Bmp;
        if (IsIco(data)) return GraphicKind.Ico;
        if (IsTiff(data)) return GraphicKind.Tiff;
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
        if (ct.Contains("gif")) return GraphicKind.Gif;
        if (ct.Contains("webp")) return GraphicKind.Webp;
        if (ct.Contains("tiff") || ct.Contains("tif")) return GraphicKind.Tiff;
        if (ct.Contains("icon")) return GraphicKind.Ico;
        if (ct.Contains("bmp")) return GraphicKind.Bmp;
        return GraphicKind.Unknown;
    }

    /// <summary>Formaty renderowalne natywnie przez aktualne przeglądarki w &lt;img&gt; (bez konwersji).</summary>
    private static bool IsBrowserNativeRaster(GraphicKind kind) => kind is
        GraphicKind.Png or GraphicKind.Jpeg or GraphicKind.Gif or GraphicKind.Bmp
        or GraphicKind.Webp or GraphicKind.Ico;

    public GraphicConversionResult ConvertForEditor(
        GraphicSource source, GraphicConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= GraphicConversionOptions.Default;
        var sw = Stopwatch.StartNew();

        if (source.Data == null || source.Data.Length == 0)
            return Rejected(GraphicKind.Unknown, source, string.Empty, sw, "Puste dane grafiki.");
        if (source.Data.Length > options.MaxInputBytes)
            return Rejected(GraphicKind.Unknown, source, string.Empty, sw, $"Przekroczono limit {options.MaxInputBytes} B.");

        var cacheKey = BuildCacheKey(source, options);
        if (_cache.TryGetValue(cacheKey, out var cached))
            return cached;

        var result = Execute(source, options, cacheKey, sw, cancellationToken);

        // Bounded cache: przestajemy dokładać po osiągnięciu limitu (zapobiega nieograniczonemu
        // wzrostowi pamięci na dokumentach z setkami unikalnych grafik).
        if (_cache.Count < MaxCacheEntries)
            _cache.TryAdd(cacheKey, result);
        return result;
    }

    private GraphicConversionResult Execute(
        GraphicSource source, GraphicConversionOptions options, string cacheKey, Stopwatch sw, CancellationToken ct)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(options.Timeout);

        var attempted = new List<string>();
        var warnings = new List<string>();
        var lost = new List<string>();
        var kind = Detect(source.Data, source.ContentType);

        // Word zapisuje metafile także w wariancie skompresowanym GZIP (EMZ/WMZ — content type
        // image/x-emz / image/x-wmz). Sygnatura 1F 8B nie pasuje do żadnego formatu graficznego,
        // więc bez dekompresji taki part kończył jako Unknown → broken image w edytorze.
        if (kind == GraphicKind.Unknown && IsGzip(source.Data))
        {
            attempted.Add("gzip-decompress");
            var inner = TryGunzip(source.Data, options.MaxInputBytes);
            if (inner != null)
            {
                warnings.Add("Skompresowany media part (EMZ/WMZ/GZIP) — zdekompresowano.");
                source = new GraphicSource
                {
                    Data = inner,
                    ContentType = source.ContentType,
                    FileName = source.FileName,
                    SourcePath = source.SourcePath,
                    Origin = source.Origin,
                    TargetWidthEmu = source.TargetWidthEmu,
                    TargetHeightEmu = source.TargetHeightEmu
                };
                kind = Detect(inner, source.ContentType);
            }
        }

        try
        {
            if (IsBrowserNativeRaster(kind))
            {
                var (w, h) = ReadRasterSize(kind, source.Data);
                return PassThroughWeb(kind, MimeFor(kind), source.Data, w, h, source, cacheKey, sw);
            }

            switch (kind)
            {
                case GraphicKind.Svg:
                    return ConvertSvg(source, cacheKey, sw, attempted, warnings, lost);

                case GraphicKind.Tiff:
                    return ConvertRasterNonNative(GraphicKind.Tiff, source, options, cacheKey, sw, attempted, warnings, lost);

                case GraphicKind.Emf:
                case GraphicKind.Wmf:
                    return ConvertMetafile(kind, source, options, cacheKey, sw, attempted, warnings, lost);

                default:
                    // Nieznany media part: ostatnia próba przez dekoder rastra (Skia) — pokrywa
                    // formaty, których nie wykryto po sygnaturze, a Skia potrafi je odczytać.
                    return ConvertRasterNonNative(GraphicKind.Unknown, source, options, cacheKey, sw, attempted, warnings, lost);
            }
        }
        catch (OperationCanceledException)
        {
            return Rejected(kind, source, cacheKey, sw, "Przekroczono limit czasu konwersji.");
        }
        catch (Exception ex)
        {
            // Niezaufane wejście nie może wywrócić importu — mapujemy na przezroczysty fallback
            // i zachowujemy szczegóły w diagnostyce (bez logowania pełnych danych dokumentu).
            attempted.Add("exception");
            return BlankFallback(kind, source, options, cacheKey, sw, attempted, warnings, lost,
                GraphicConversionStatus.Unsupported, $"{ex.GetType().Name}: {ex.Message}");
        }
    }

    // ---- per-format strategies ----------------------------------------------------

    private GraphicConversionResult ConvertSvg(
        GraphicSource source, string cacheKey, Stopwatch sw,
        List<string> attempted, List<string> warnings, List<string> lost)
    {
        attempted.Add("svg-sanitize");
        var svg = SanitizeSvg(Encoding.UTF8.GetString(source.Data));
        if (svg == null)
            return BlankFallback(GraphicKind.Svg, source, GraphicConversionOptions.Default, cacheKey, sw,
                attempted, warnings, lost, GraphicConversionStatus.Unsupported, "SVG nieparsowalny/niebezpieczny.");

        var bytes = Encoding.UTF8.GetBytes(svg);
        var (w, h) = ReadSvgSize(svg, source);
        return new GraphicConversionResult
        {
            Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = bytes, WidthPx = w, HeightPx = h },
            PreserveOriginalPart = false,
            Diagnostics = Diag(GraphicKind.Svg, "image/svg+xml", GraphicConversionStatus.Converted,
                GraphicFidelity.Lossless, source, cacheKey, sw, attempted, null, warnings, lost)
        };
    }

    /// <summary>
    /// Raster nie-natywny dla przeglądarki (TIFF) lub nieznany: dekodujemy przez SkiaSharp do PNG.
    /// Gdy Skia nie ma kodeka (np. TIFF bez libtiff w obrazie kontenera) → przezroczysty fallback,
    /// oryginał zachowany do DOCX. Bez widocznego placeholdera.
    /// </summary>
    private GraphicConversionResult ConvertRasterNonNative(
        GraphicKind kind, GraphicSource source, GraphicConversionOptions options, string cacheKey, Stopwatch sw,
        List<string> attempted, List<string> warnings, List<string> lost)
    {
        attempted.Add("skia-decode");
        var png = TryDecodeRasterToPng(source.Data);
        if (png != null)
        {
            var (w, h) = ReadPngSize(png);
            if (kind == GraphicKind.Tiff) lost.Add("Wielostronicowość/CMYK TIFF poza pierwszą klatką.");
            warnings.Add($"{kind}: zdekodowano do PNG (SkiaSharp).");
            return new GraphicConversionResult
            {
                Web = new WebGraphicRepresentation { MimeType = "image/png", Data = png, WidthPx = w, HeightPx = h },
                PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart && kind != GraphicKind.Unknown,
                Diagnostics = Diag(kind, "image/png", GraphicConversionStatus.Converted, GraphicFidelity.Lossy,
                    source, cacheKey, sw, attempted, null, warnings, lost)
            };
        }

        var reason = kind == GraphicKind.Tiff
            ? "Brak kodeka TIFF w SkiaSharp na tej platformie."
            : "Nierozpoznany/nieobsługiwany format media partu.";
        return BlankFallback(kind, source, options, cacheKey, sw, attempted, warnings, lost,
            GraphicConversionStatus.Unsupported, reason);
    }

    /// <summary>
    /// Łańcuch strategii dla metafile EMF/WMF (od najwierniejszej):
    /// 1) osadzony gotowy raster (PNG/JPEG) → pokaż wprost,
    /// 2) osadzony DIB (StretchDIBits itp.) → DIB→BMP→SkiaSharp→PNG,
    /// 3) brak rastra (czysty wektor) → przezroczysty fallback (NIE placeholder); oryginał do DOCX.
    /// Brak pure-managed rasteryzera wektora EMF/WMF bez GDI — patrz .ai/GRAPHICS_CONVERSION.md.
    /// </summary>
    private GraphicConversionResult ConvertMetafile(
        GraphicKind kind, GraphicSource source, GraphicConversionOptions options, string cacheKey, Stopwatch sw,
        List<string> attempted, List<string> warnings, List<string> lost)
    {
        var (w, h) = kind == GraphicKind.Emf ? ReadEmfSize(source.Data) : ReadWmfSize(source.Data);
        (w, h) = ResolveDims(w, h, source, options);

        attempted.Add("embedded-raster");
        var embedded = TryExtractEmbeddedRaster(source.Data);
        if (embedded != null)
        {
            var ek = Detect(embedded);
            var (ew, eh) = ReadRasterSize(ek, embedded);
            warnings.Add($"{kind}: pokazano osadzony {ek} z metafile.");
            lost.Add("Wektorowe elementy metafile poza osadzonym rastrem.");
            return MetafileRaster(kind, MimeFor(ek), embedded, ew > 0 ? ew : w, eh > 0 ? eh : h,
                source, cacheKey, sw, attempted, warnings, lost);
        }

        attempted.Add("dib-rasterize");
        var png = TryRasterizeMetafileToPng(source.Data, kind);
        if (png != null)
        {
            var (pw, ph) = ReadPngSize(png);
            warnings.Add($"{kind}: zrasteryzowano osadzoną bitmapę do PNG (SkiaSharp).");
            lost.Add("Elementy wektorowe metafile poza zrasteryzowaną bitmapą.");
            return MetafileRaster(kind, "image/png", png, pw > 0 ? pw : w, ph > 0 ? ph : h,
                source, cacheKey, sw, attempted, warnings, lost);
        }

        // Etap 1 własnego tłumacza wektorowego (pure-managed): podzbiór rekordów GDI → SVG.
        // Tylko podgląd (Lossy) — oryginalny metafile i tak jedzie do DOCX przez pass-through.
        attempted.Add("vector-translate");
        var vector = MetafileVectorTranslator.Translate(kind, source.Data, w, h);
        if (vector != null && SanitizeSvg(vector.Svg) is { } safeSvg)
        {
            warnings.Add($"{kind}: rekordy wektorowe przetłumaczone na SVG (etap 1 — podzbiór GDI).");
            if (vector.SkippedRecords > 0)
            {
                warnings.Add($"{kind}: pominięto {vector.SkippedRecords} rekordów spoza podzbioru.");
                lost.Add("Rekordy GDI spoza podzbioru etapu 1 (tekst, clipping, ROP, EMF+).");
            }
            return new GraphicConversionResult
            {
                Web = new WebGraphicRepresentation
                {
                    MimeType = "image/svg+xml",
                    Data = Encoding.UTF8.GetBytes(safeSvg),
                    WidthPx = w,
                    HeightPx = h
                },
                PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
                Diagnostics = Diag(kind, "image/svg+xml", GraphicConversionStatus.Converted, GraphicFidelity.Lossy,
                    source, cacheKey, sw, attempted, null, warnings, lost)
            };
        }

        // Brak osadzonego rastra i nic do przetłumaczenia → przezroczysty, niewidoczny element
        // zachowujący wymiary/układ. Oryginał zachowany do eksportu (data-original-src).
        lost.Add("Podgląd wektorowego metafile (renderowany w Word z zachowanego oryginału).");
        return BlankFallback(kind, source, options, cacheKey, sw, attempted, warnings, lost,
            GraphicConversionStatus.Fallback,
            $"{kind} czysto wektorowy bez osadzonego rastra i bez rekordów obsługiwanych przez tłumacz SVG.",
            preComputedDims: (w, h));
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
        if (shape == null) return null; // nieobsługiwany kształt → caller rozwiąże inaczej

        var svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' viewBox='0 0 {w} {h}'>{shape}</svg>";
        var bytes = Encoding.UTF8.GetBytes(svg);
        return new GraphicConversionResult
        {
            Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = bytes, WidthPx = w, HeightPx = h },
            PreserveOriginalPart = true, // VML zachowujemy w DOCX przez pass-through, póki nieedytowany
            Diagnostics = Diag(GraphicKind.Vml, "image/svg+xml", GraphicConversionStatus.Converted, GraphicFidelity.Lossy,
                null, string.Empty, sw,
                new[] { "vml-shape" },
                null,
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

    // ---- cache key ----------------------------------------------------------------

    /// <summary>
    /// Klucz cache = SHA-256(treść) + parametry wpływające na wynik (content-type, origin, wymiary
    /// docelowe, limity). Deterministyczny: identyczny asset → identyczny wynik → deduplikacja.
    /// </summary>
    private static string BuildCacheKey(GraphicSource source, GraphicConversionOptions o)
    {
        var hash = Convert.ToHexString(SHA256.HashData(source.Data));
        return string.Create(CultureInfo.InvariantCulture,
            $"{hash}|ct={source.ContentType}|o={(int)source.Origin}|tw={source.TargetWidthEmu}|th={source.TargetHeightEmu}|mw={o.MaxPlaceholderWidthPx}|mh={o.MaxPlaceholderHeightPx}|mb={o.MaxInputBytes}");
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

    private static bool IsWebp(ReadOnlySpan<byte> d) =>
        d.Length >= 12 && d[0] == (byte)'R' && d[1] == (byte)'I' && d[2] == (byte)'F' && d[3] == (byte)'F'
        && d[8] == (byte)'W' && d[9] == (byte)'E' && d[10] == (byte)'B' && d[11] == (byte)'P';

    private static bool IsIco(ReadOnlySpan<byte> d) =>
        // ICONDIR: reserved=0, type=1 (icon) or 2 (cursor), count>=1.
        d.Length >= 6 && d[0] == 0 && d[1] == 0 && (d[2] == 1 || d[2] == 2) && d[3] == 0
        && BinaryPrimitives.ReadUInt16LittleEndian(d[4..]) >= 1;

    private static bool IsTiff(ReadOnlySpan<byte> d)
    {
        if (d.Length < 4) return false;
        // "II" 0x2A00 (little-endian) lub "MM" 0x002A (big-endian).
        if (d[0] == 0x49 && d[1] == 0x49 && d[2] == 0x2A && d[3] == 0x00) return true;
        if (d[0] == 0x4D && d[1] == 0x4D && d[2] == 0x00 && d[3] == 0x2A) return true;
        return false;
    }

    private static bool IsGzip(ReadOnlySpan<byte> d) => d.Length >= 3 && d[0] == 0x1F && d[1] == 0x8B;

    /// <summary>Dekompresja GZIP z twardym limitem rozmiaru wyjścia (ochrona przed decompression bomb).</summary>
    private static byte[]? TryGunzip(byte[] data, int maxOutputBytes)
    {
        try
        {
            using var input = new MemoryStream(data, writable: false);
            using var gz = new GZipStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            var buffer = new byte[81920];
            int read;
            while ((read = gz.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (output.Length + read > maxOutputBytes) return null; // bomba/oversize → odrzuć
                output.Write(buffer, 0, read);
            }
            return output.Length > 0 ? output.ToArray() : null;
        }
        catch { return null; }
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
        GraphicKind.Webp or GraphicKind.Ico or GraphicKind.Tiff => TrySkiaDims(d),
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

    /// <summary>Tani odczyt wymiarów rastra przez nagłówek kodeka SkiaSharp (bez pełnego dekodowania).</summary>
    private static (int w, int h) TrySkiaDims(byte[] d)
    {
        try
        {
            using var codec = SKCodec.Create(new MemoryStream(d, writable: false));
            if (codec != null && codec.Info.Width > 0 && codec.Info.Height > 0)
                return (codec.Info.Width, codec.Info.Height);
        }
        catch { /* nieobsługiwany kodek → wymiary z layoutu DOCX */ }
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

    /// <summary>Buduje wynik dla metafile pokazanego jako realny raster (nie placeholder).</summary>
    private GraphicConversionResult MetafileRaster(
        GraphicKind kind, string mime, byte[] raster, int w, int h, GraphicSource source, string cacheKey,
        Stopwatch sw, List<string> attempted, List<string> warnings, List<string> lost) => new()
    {
        Web = new WebGraphicRepresentation { MimeType = mime, Data = raster, WidthPx = w, HeightPx = h, IsBlankFallback = false },
        // Oryginalny metafile jedzie do DOCX (pass-through) — Word renderuje wektor; PNG to tylko podgląd.
        PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
        Diagnostics = Diag(kind, mime, GraphicConversionStatus.Converted, GraphicFidelity.Lossy,
            source, cacheKey, sw, attempted, null, warnings, lost)
    };

    // ---- metafile rasterization (pure-managed via SkiaSharp) ----------------------

    /// <summary>Dekoduje dowolny raster znany SkiaSharp i re-enkoduje do PNG (TIFF/Unknown rescue).</summary>
    private static byte[]? TryDecodeRasterToPng(byte[] data)
    {
        try
        {
            using var skbmp = SKBitmap.Decode(data);
            if (skbmp == null || skbmp.Width <= 0 || skbmp.Height <= 0) return null;
            using var image = SKImage.FromBitmap(skbmp);
            using var encoded = image.Encode(SKEncodedImageFormat.Png, 100);
            return encoded?.ToArray();
        }
        catch { return null; }
    }

    /// <summary>
    /// Rasteryzuje metafile do PNG, wydobywając osadzony DIB (device-independent bitmap) z rekordów
    /// EMF/WMF i dekodując go przez SkiaSharp (cross-platform; bez GDI/System.Drawing/LibreOffice).
    /// Pokrywa najczęstszy realny przypadek — EMF/WMF opakowujący bitmapę (np. StretchDIBits). Zwraca
    /// null dla metafile czysto wektorowego (brak DIB) lub gdy dekodowanie się nie powiedzie.
    /// </summary>
    private static byte[]? TryRasterizeMetafileToPng(byte[] data, GraphicKind kind)
    {
        var dib = kind == GraphicKind.Emf ? TryExtractEmfDib(data) : null;
        dib ??= TryFindDibGeneric(data); // fallback (także dla WMF) — skan po nagłówku BITMAPINFOHEADER
        if (dib == null) return null;

        var bmp = WrapDibInBmpFile(dib);
        return bmp == null ? null : DecodeToPng(bmp);
    }

    /// <summary>
    /// Iteruje rekordy EMF i zwraca największy DIB niesiony przez rekordy rastrowe (offsety pól
    /// offBmiSrc/cbBmiSrc/offBitsSrc/cbBitsSrc są względem początku rekordu — MS-EMF).
    /// </summary>
    private static byte[]? TryExtractEmfDib(byte[] d)
    {
        byte[]? best = null;
        int o = 0, guard = 0;
        while (o + 8 <= d.Length && guard++ < 200_000)
        {
            uint iType = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o));
            uint nSize = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o + 4));
            if (nSize < 8 || (long)o + nSize > d.Length) break; // uszkodzony strumień → przerwij

            // Pozycja pól DIB zależy od typu rekordu (MS-EMF 2.3.1):
            //  STRETCHDIBITS(81)/SETDIBITSTODEVICE(80) → offBmiSrc na offsecie 48,
            //  BITBLT(76)/STRETCHBLT(77)/ALPHABLEND(114) → offBmiSrc na offsecie 84.
            int? bmiFieldPos = iType switch
            {
                80 or 81 => 48,
                76 or 77 or 114 => 84,
                _ => null
            };
            if (bmiFieldPos is int p && p + 16 <= nSize)
            {
                uint offBmi = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o + p));
                uint cbBmi = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o + p + 4));
                uint offBits = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o + p + 8));
                uint cbBits = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o + p + 12));
                var dib = SliceDib(d, o, offBmi, cbBmi, offBits, cbBits, nSize);
                if (dib != null && (best == null || dib.Length > best.Length)) best = dib;
            }

            if (iType == 14) break; // EMR_EOF
            o += (int)nSize;
        }
        return best;
    }

    /// <summary>DIB → BMP → PNG (SkiaSharp). Współdzielone z tłumaczem wektorowym (StretchDIBits → &lt;image&gt;).</summary>
    internal static byte[]? DibToPng(byte[] dib)
    {
        var bmp = WrapDibInBmpFile(dib);
        return bmp == null ? null : DecodeToPng(bmp);
    }

    /// <summary>Wycina blok (BITMAPINFOHEADER+palety + bity) z rekordu wg offsetów względnych.</summary>
    internal static byte[]? SliceDib(byte[] d, int recStart, uint offBmi, uint cbBmi, uint offBits, uint cbBits, uint recSize)
    {
        if (cbBmi < 40 || cbBits == 0) return null;
        long bmiAbs = (long)recStart + offBmi, bitsAbs = (long)recStart + offBits;
        if (offBmi + cbBmi > recSize || offBits + cbBits > recSize) return null;
        if (bmiAbs + cbBmi > d.Length || bitsAbs + cbBits > d.Length) return null;

        var bmi = d.AsSpan((int)bmiAbs, (int)cbBmi);
        // Sanity-check nagłówka: biSize=40 (BITMAPINFOHEADER) i sensowny bitCount/compression.
        uint biSize = BinaryPrimitives.ReadUInt32LittleEndian(bmi);
        if (biSize != 40) return null;
        var dib = new byte[cbBmi + cbBits];
        bmi.CopyTo(dib);
        d.AsSpan((int)bitsAbs, (int)cbBits).CopyTo(dib.AsSpan((int)cbBmi));
        return dib;
    }

    /// <summary>
    /// Konserwatywny skan: szuka prawidłowego BITMAPINFOHEADER i odtwarza DIB z policzonych rozmiarów
    /// (palety + bitów). Fallback dla WMF i nietypowych EMF; przy niejasności zwraca null.
    /// </summary>
    private static byte[]? TryFindDibGeneric(byte[] d)
    {
        for (int i = 0; i + 40 <= d.Length; i += 2)
        {
            uint biSize = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(i));
            if (biSize != 40) continue;
            int width = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(i + 4));
            int height = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(i + 8));
            ushort planes = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(i + 12));
            ushort bitCount = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(i + 14));
            uint compression = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(i + 16));
            uint clrUsed = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(i + 32));

            if (planes != 1) continue;
            if (bitCount is not (1 or 4 or 8 or 16 or 24 or 32)) continue;
            if (compression is not (0u or 3u)) continue;          // tylko BI_RGB / BI_BITFIELDS (bez RLE)
            int absH = Math.Abs(height);
            if (width <= 0 || width > 20000 || absH <= 0 || absH > 20000) continue;

            int masks = compression == 3 ? 12 : 0;                 // BI_BITFIELDS → 3×DWORD po nagłówku
            int paletteEntries = bitCount <= 8 ? (clrUsed != 0 ? (int)clrUsed : 1 << bitCount) : 0;
            int cbBmi = 40 + masks + paletteEntries * 4;
            int stride = ((width * bitCount + 31) / 32) * 4;
            long cbBits = (long)stride * absH;
            if (cbBits <= 0 || cbBits > 64L * 1024 * 1024) continue;
            if (i + cbBmi + cbBits > d.Length) continue;

            var dib = new byte[cbBmi + cbBits];
            d.AsSpan(i, (int)(cbBmi + cbBits)).CopyTo(dib);
            return dib;
        }
        return null;
    }

    /// <summary>Opakowuje DIB w plik BMP (BITMAPFILEHEADER + DIB), gotowy do dekodowania.</summary>
    private static byte[]? WrapDibInBmpFile(byte[] dib)
    {
        if (dib.Length < 40) return null;
        uint cbBmi = BinaryPrimitives.ReadUInt32LittleEndian(dib);          // biSize = 40
        ushort bitCount = BinaryPrimitives.ReadUInt16LittleEndian(dib.AsSpan(14));
        uint compression = BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(16));
        uint clrUsed = BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(32));
        int masks = compression == 3 ? 12 : 0;
        int paletteEntries = bitCount <= 8 ? (clrUsed != 0 ? (int)clrUsed : 1 << bitCount) : 0;
        long headerAndPalette = cbBmi + masks + paletteEntries * 4L;
        if (headerAndPalette > dib.Length) headerAndPalette = Math.Min(40 + masks, dib.Length);

        const int FileHeader = 14;
        var bmp = new byte[FileHeader + dib.Length];
        bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(bmp.AsSpan(2), (uint)(FileHeader + dib.Length)); // bfSize
        BinaryPrimitives.WriteUInt32LittleEndian(bmp.AsSpan(10), (uint)(FileHeader + headerAndPalette)); // bfOffBits
        dib.CopyTo(bmp.AsSpan(FileHeader));
        return bmp;
    }

    /// <summary>Dekoduje BMP przez SkiaSharp i re-enkoduje do PNG. Zwraca null przy błędzie.</summary>
    private static byte[]? DecodeToPng(byte[] bmp)
    {
        using var skbmp = SKBitmap.Decode(bmp);
        if (skbmp == null || skbmp.Width <= 0 || skbmp.Height <= 0) return null;
        using var image = SKImage.FromBitmap(skbmp);
        using var encoded = image.Encode(SKEncodedImageFormat.Png, 100);
        return encoded?.ToArray();
    }

    // ---- transparent fallback + parsing helpers ----------------------------------

    /// <summary>
    /// Niewidoczna, w pełni przezroczysta grafika SVG o wymiarach z layoutu DOCX. Zachowuje miejsce
    /// w układzie BEZ rysowania jakiejkolwiek widocznej treści (brak szarego tła, ramki, tekstu
    /// „requires conversion"). To honest internal-failure state — realny wektor jest w oryginalnym
    /// part (pass-through do DOCX), a tu nie udajemy treści, której nie potrafimy zrenderować w web.
    /// </summary>
    private static byte[] BuildTransparentSvg(int w, int h)
    {
        if (w <= 0) w = 1;
        if (h <= 0) h = 1;
        var svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' " +
                  $"viewBox='0 0 {w} {h}' aria-hidden='true'></svg>";
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
        GraphicKind.Webp => "image/webp",
        GraphicKind.Ico => "image/x-icon",
        GraphicKind.Tiff => "image/tiff",
        GraphicKind.Svg => "image/svg+xml",
        _ => "application/octet-stream"
    };

    // ---- result factories ---------------------------------------------------------

    private GraphicConversionResult PassThroughWeb(
        GraphicKind kind, string mime, byte[] data, int w, int h, GraphicSource source, string cacheKey, Stopwatch sw) =>
        new()
        {
            Web = new WebGraphicRepresentation { MimeType = mime, Data = data, WidthPx = w, HeightPx = h },
            PreserveOriginalPart = false,
            Diagnostics = Diag(kind, mime, GraphicConversionStatus.PassThrough, GraphicFidelity.Lossless,
                source, cacheKey, sw, new[] { "pass-through" }, null, Array.Empty<string>(), Array.Empty<string>())
        };

    /// <summary>
    /// Przezroczysty (niewidoczny) fallback — gdy żadna realna strategia nie zwróciła rastra.
    /// NIGDY nie produkuje widocznego placeholdera. Diagnostyka niesie powód i listę prób.
    /// </summary>
    private GraphicConversionResult BlankFallback(
        GraphicKind kind, GraphicSource source, GraphicConversionOptions o, string cacheKey, Stopwatch sw,
        List<string> attempted, List<string> warnings, List<string> lost,
        GraphicConversionStatus status, string failureReason, (int w, int h)? preComputedDims = null)
    {
        warnings.Add(failureReason);
        var (w, h) = preComputedDims ?? ResolveDims(0, 0, source, o);
        var blank = BuildTransparentSvg(w, h);
        var fidelity = status == GraphicConversionStatus.Fallback ? GraphicFidelity.Fallback : GraphicFidelity.Unsupported;
        return new GraphicConversionResult
        {
            Web = new WebGraphicRepresentation { MimeType = "image/svg+xml", Data = blank, WidthPx = w, HeightPx = h, IsBlankFallback = true },
            PreserveOriginalPart = source.Origin == GraphicOrigin.LegacyDocxPart,
            Diagnostics = Diag(kind, "image/svg+xml", status, fidelity, source, cacheKey, sw, attempted, failureReason, warnings, lost)
        };
    }

    private GraphicConversionResult Rejected(GraphicKind kind, GraphicSource source, string cacheKey, Stopwatch sw, string reason) =>
        new()
        {
            Web = null,
            PreserveOriginalPart = false,
            Diagnostics = Diag(kind, string.Empty, GraphicConversionStatus.Rejected, GraphicFidelity.Unsupported,
                source, cacheKey, sw, new[] { "reject" }, reason, new[] { reason }, Array.Empty<string>())
        };

    private static GraphicConversionDiagnostics Diag(
        GraphicKind kind, string mime, GraphicConversionStatus status, GraphicFidelity fidelity,
        GraphicSource? source, string cacheKey, Stopwatch sw,
        IReadOnlyList<string> attempted, string? failureReason,
        IReadOnlyList<string> warnings, IReadOnlyList<string> lost)
    {
        sw.Stop();
        return new GraphicConversionDiagnostics
        {
            InputKind = kind,
            SourcePath = source?.SourcePath,
            OutputMimeType = mime,
            Status = status,
            Fidelity = fidelity,
            ElapsedMs = sw.ElapsedMilliseconds,
            CacheKey = cacheKey,
            AttemptedStrategies = attempted,
            FailureReason = failureReason,
            Warnings = warnings,
            LostProperties = lost
        };
    }
}

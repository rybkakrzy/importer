using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Etap 1 własnego tłumacza wektorowego EMF/WMF → SVG (pure-managed, bez GDI/System.Drawing/
/// LibreOffice → identyczne zachowanie Windows i Linux/GCP). Obsługiwany podzbiór rekordów GDI:
/// pióra/pędzle (kolor, grubość, styl kreski), MoveTo/LineTo, Rectangle/Ellipse/RoundRect,
/// Polygon/Polyline/PolyBezier (warianty 16- i 32-bitowe), PolyPolygon, ścieżki
/// (BeginPath/EndPath/Fill/Stroke), transformacje świata (SetWorldTransform/ModifyWorldTransform),
/// StretchDIBits jako &lt;image&gt;. Rekordy spoza podzbioru (tekst, clipping, ROP, EMF+) są
/// pomijane i raportowane licznikiem — konsument decyduje, czy wynik jest użyteczny.
///
/// Tłumaczenie jest podglądem (Lossy) — oryginalny metafile ZAWSZE jedzie do DOCX przez
/// pass-through (data-original-src / ConvertPreservingPackage), więc Word renderuje prawdziwy wektor.
/// Wejście niezaufane: twarde limity rekordów/punktów/rozmiaru wyjścia, każdy wyjątek → null
/// (łańcuch strategii przechodzi na blank fallback).
/// </summary>
internal static class MetafileVectorTranslator
{
    private const int MaxRecords = 200_000;
    private const int MaxPointsPerPoly = 100_000;
    private const int MaxSvgElements = 20_000;
    private const int MaxOutputChars = 2_000_000;

    internal sealed class MetafileSvg
    {
        public required string Svg { get; init; }
        public int SkippedRecords { get; init; }
        /// <summary>Liczba przetworzonych rekordów metafile (diagnostyka).</summary>
        public int RecordCount { get; init; }
        /// <summary>Czy zastosowano mapowanie window→viewport (metafile zdefiniował oba zakresy).</summary>
        public bool UsedWindowViewport { get; init; }
        /// <summary>
        /// Treść po zmapowaniu NIE pokryła się z rclBounds → viewBox wzięto z bbox treści (safety-net).
        /// Sygnał, że mapowanie page→device było niepełne (np. tryb metryczny) — ryzyko złego kadru.
        /// </summary>
        public bool ContentOutsideDeviceBounds { get; init; }
        /// <summary>Wyróżnione kolory wypełnień w SVG — diagnostyka „logo wyszło czarne" (fill=[#000000]).</summary>
        public IReadOnlyList<string> FillColors { get; init; } = Array.Empty<string>();
        /// <summary>Czy SVG zawiera osadzony raster (STRETCHDIBITS/ALPHABLEND) — może być właściwym obrazem logo.</summary>
        public bool HasEmbeddedImage { get; init; }
    }

    /// <summary>Czytelne nazwy najważniejszych rekordów EMF (diagnostyka „dlaczego blank").</summary>
    private static readonly Dictionary<uint, string> EmrNames = new()
    {
        [1] = "HEADER", [9] = "SETWINDOWEXTEX", [10] = "SETWINDOWORGEX", [11] = "SETVIEWPORTEXTEX",
        [12] = "SETVIEWPORTORGEX", [14] = "EOF", [17] = "SETMAPMODE", [19] = "SETPOLYFILLMODE",
        [27] = "MOVETOEX", [35] = "SETWORLDTRANSFORM", [36] = "MODIFYWORLDTRANSFORM", [37] = "SELECTOBJECT",
        [38] = "CREATEPEN", [39] = "CREATEBRUSHINDIRECT", [40] = "DELETEOBJECT", [42] = "ELLIPSE",
        [43] = "RECTANGLE", [44] = "ROUNDRECT", [54] = "LINETO", [59] = "BEGINPATH", [60] = "ENDPATH",
        [61] = "CLOSEFIGURE", [62] = "FILLPATH", [63] = "STROKEANDFILLPATH", [64] = "STROKEPATH",
        [70] = "GDICOMMENT", [81] = "STRETCHDIBITS", [84] = "EXTTEXTOUTA", [85] = "POLYBEZIER16",
        [86] = "POLYGON16", [87] = "POLYLINE16", [88] = "POLYBEZIERTO16", [89] = "POLYLINETO16",
        [91] = "POLYPOLYGON16", [95] = "EXTCREATEPEN", [98] = "SETICMMODE", [114] = "ALPHABLEND", [115] = "SETLAYOUT",
    };

    /// <summary>
    /// Zwięzły profil rekordów metafile do LOGU, gdy tłumaczenie nie dało widocznej treści.
    /// Read-only skan (bez rysowania) — np. „EMF rec=210 win/vp=yes world=no [MOVETOEX:74,LINETO:74,
    /// POLYBEZIERTO16:74,FILLPATH:3,BEGINPATH:3,#95:1,...]". Każdy wyjątek → krótka informacja.
    /// </summary>
    public static string Profile(GraphicKind kind, byte[] data)
    {
        try
        {
            return kind switch
            {
                GraphicKind.Emf => ProfileEmf(data),
                GraphicKind.Wmf => ProfileWmf(data),
                _ => $"{kind} (brak profilera)"
            };
        }
        catch (Exception ex)
        {
            return $"{kind} profil-niedostępny: {ex.GetType().Name}";
        }
    }

    private static string ProfileEmf(byte[] d)
    {
        if (d.Length < 88) return "EMF za krótki";
        uint headerSize = U32(d, 4);
        if (headerSize < 88 || headerSize > (uint)d.Length) return "EMF zły nagłówek";

        var hist = new Dictionary<uint, int>();
        bool win = false, vp = false, world = false;
        int o = (int)headerSize, guard = 0, total = 0;
        while (o + 8 <= d.Length && guard++ < MaxRecords)
        {
            uint iType = U32(d, o);
            uint nSize = U32(d, o + 4);
            if (nSize < 8 || nSize % 4 != 0 || (long)o + nSize > d.Length) break;
            hist[iType] = hist.GetValueOrDefault(iType) + 1;
            total++;
            if (iType == 9) win = true;
            else if (iType == 11) vp = true;
            else if (iType is 35 or 36) world = true;
            if (iType == 14) break; // EOF
            o += (int)nSize;
        }
        return $"EMF rec={total} win/vp={(win && vp ? "yes" : "no")} world={(world ? "yes" : "no")} [{FormatHist(hist)}]";
    }

    private static string ProfileWmf(byte[] d)
    {
        int o = d.Length >= 22 && U32(d, 0) == 0x9AC6CDD7 ? 22 : 0;
        if (o + 18 > d.Length) return "WMF za krótki";
        o += 18;
        var hist = new Dictionary<uint, int>();
        int guard = 0, total = 0;
        while (o + 6 <= d.Length && guard++ < MaxRecords)
        {
            uint sizeWords = U32(d, o);
            ushort func = U16(d, o + 4);
            long byteSize = (long)sizeWords * 2;
            if (sizeWords < 3 || o + byteSize > d.Length) break;
            hist[func] = hist.GetValueOrDefault(func) + 1;
            total++;
            if (func == 0x0000) break; // META_EOF
            o += (int)byteSize;
        }
        return $"WMF rec={total} [{FormatHist(hist)}]";
    }

    private static string FormatHist(Dictionary<uint, int> hist) =>
        string.Join(",", hist.OrderByDescending(kv => kv.Value).Take(16)
            .Select(kv => $"{(EmrNames.TryGetValue(kv.Key, out var n) ? n : $"#{kv.Key}")}:{kv.Value}"));

    public static MetafileSvg? Translate(GraphicKind kind, byte[] data, int widthPx, int heightPx)
    {
        try
        {
            return kind switch
            {
                GraphicKind.Emf => TranslateEmf(data, widthPx, heightPx),
                GraphicKind.Wmf => TranslateWmf(data, widthPx, heightPx),
                _ => null
            };
        }
        catch
        {
            // Uszkodzony/wrogi metafile nie może wywrócić importu — łańcuch przejdzie na blank.
            return null;
        }
    }

    // ---- wspólny stan GDI -----------------------------------------------------------------

    private readonly record struct PenDef(string Color, double Width, string? Dash, bool IsNull)
    {
        public static PenDef Default => new("#000000", 1, null, false);
        public static PenDef Null => new("#000000", 1, null, true);
    }

    private readonly record struct BrushDef(string Color, bool IsNull)
    {
        public static BrushDef White => new("#ffffff", false);
        public static BrushDef Null => new("#000000", true);
    }

    private sealed class GdiState
    {
        public PenDef Pen = PenDef.Default;     // GDI default: BLACK_PEN 1px solid
        public BrushDef Brush = BrushDef.White; // GDI default: WHITE_BRUSH
        public double X, Y;                     // bieżąca pozycja (jednostki logiczne)
        public string FillRule = "evenodd";     // GDI default: ALTERNATE

        // World transform (konwencja wektora wierszowego: p' = p·M).
        public double M11 = 1, M12, M21, M22 = 1, Dx, Dy;

        // Mapowanie page→device (SETWINDOW*/SETVIEWPORT*, tryb anizotropowy). Współrzędne
        // rysowania są LOGICZNE; bez tego mapowania lądują poza viewBox (= rclBounds z nagłówka,
        // w jednostkach URZĄDZENIA) i cała grafika wychodzi pusta — np. wektorowe logo w stopce.
        public double WinOrgX, WinOrgY, VpOrgX, VpOrgY;
        public double WinExtX = 1, WinExtY = 1, VpExtX = 1, VpExtY = 1;
        public bool HasWindowExt, HasViewportExt;

        public bool IsAxisAligned => M12 == 0 && M21 == 0;

        public (double X, double Y) Apply(double x, double y)
        {
            // 1) world transform: logical → page
            double px = x * M11 + y * M21 + Dx;
            double py = x * M12 + y * M22 + Dy;
            // 2) page → device (window/viewport). Tylko gdy metafile faktycznie zdefiniował OBA
            //    zakresy — inaczej tożsamość (zero zmian dla metafile bez tych rekordów).
            if (HasWindowExt && HasViewportExt && WinExtX != 0 && WinExtY != 0)
            {
                px = (px - WinOrgX) * (VpExtX / WinExtX) + VpOrgX;
                py = (py - WinOrgY) * (VpExtY / WinExtY) + VpOrgY;
            }
            return (px, py);
        }

        public void SetTransform(double m11, double m12, double m21, double m22, double dx, double dy)
        { M11 = m11; M12 = m12; M21 = m21; M22 = m22; Dx = dx; Dy = dy; }

        /// <summary>combined = a następnie b (row-vector: p·(A·B)).</summary>
        public void Combine(
            double a11, double a12, double a21, double a22, double adx, double ady, bool currentFirst)
        {
            double b11 = M11, b12 = M12, b21 = M21, b22 = M22, bdx = Dx, bdy = Dy;
            if (currentFirst)
            {
                // new = current × xform (MWT_RIGHTMULTIPLY)
                (b11, b12, b21, b22, bdx, bdy, a11, a12, a21, a22, adx, ady) =
                    (a11, a12, a21, a22, adx, ady, M11, M12, M21, M22, Dx, Dy);
            }
            SetTransform(
                a11 * b11 + a12 * b21,
                a11 * b12 + a12 * b22,
                a21 * b11 + a22 * b21,
                a21 * b12 + a22 * b22,
                adx * b11 + ady * b21 + bdx,
                adx * b12 + ady * b22 + bdy);
        }
    }

    private sealed class SvgCanvas
    {
        private readonly StringBuilder _sb = new();
        public int ElementCount;
        public double MinX = double.MaxValue, MinY = double.MaxValue,
                      MaxX = double.MinValue, MaxY = double.MinValue;

        public bool Full => ElementCount >= MaxSvgElements || _sb.Length >= MaxOutputChars;
        public bool HasContent => ElementCount > 0;
        public string Elements => _sb.ToString();

        public void Track(double x, double y)
        {
            if (x < MinX) MinX = x;
            if (y < MinY) MinY = y;
            if (x > MaxX) MaxX = x;
            if (y > MaxY) MaxY = y;
        }

        public void Add(string element)
        {
            if (Full) return;
            _sb.Append(element);
            ElementCount++;
        }
    }

    // ---- EMF -------------------------------------------------------------------------------

    private static MetafileSvg? TranslateEmf(byte[] d, int widthPx, int heightPx)
    {
        if (d.Length < 88) return null;
        uint headerSize = U32(d, 4);
        if (headerSize < 88 || headerSize > (uint)d.Length) return null;

        double bLeft = I32(d, 8), bTop = I32(d, 12), bRight = I32(d, 16), bBottom = I32(d, 20);
        bool haveVb = bRight > bLeft && bBottom > bTop;

        var st = new GdiState();
        var objects = new Dictionary<uint, object>();
        var canvas = new SvgCanvas();
        StringBuilder? path = null;
        int skipped = 0;

        // Rekordy wyłącznie stanowe/nieistotne dla etapu 1 — pomijane BEZ liczenia jako strata
        // (nie niosą treści graficznej): mapmode/viewport/bk/text-color/save-restore/komentarze.
        // SETWINDOW*/SETVIEWPORT* są teraz obsługiwane jawnie (mapowanie page→device) — NIE cichą.
        var silent = new HashSet<uint>
        {
            17, 18, 20, 21, 22,   // SETMAPMODE, SETBKMODE, SETROP2, SETSTRETCHBLTMODE, SETTEXTALIGN
            24, 25, 33, 34, 58,   // SETTEXTCOLOR, SETBKCOLOR, SAVEDC, RESTOREDC, SETMITERLIMIT
            70, 98, 115           // GDICOMMENT (w tym kontener EMF+), SETICMMODE, SETLAYOUT
        };

        int o = (int)headerSize;
        int guard = 0;
        while (o + 8 <= d.Length && guard++ < MaxRecords && !canvas.Full)
        {
            uint iType = U32(d, o);
            uint nSize = U32(d, o + 4);
            if (nSize < 8 || nSize % 4 != 0 || (long)o + nSize > d.Length) break;
            if (iType == 14) break; // EMR_EOF

            switch (iType)
            {
                case 9:  // SETWINDOWEXTEX (SizeL: cx, cy)
                    if (nSize >= 16) { st.WinExtX = I32(d, o + 8); st.WinExtY = I32(d, o + 12); st.HasWindowExt = true; }
                    break;
                case 10: // SETWINDOWORGEX (PointL: x, y)
                    if (nSize >= 16) { st.WinOrgX = I32(d, o + 8); st.WinOrgY = I32(d, o + 12); }
                    break;
                case 11: // SETVIEWPORTEXTEX (SizeL: cx, cy)
                    if (nSize >= 16) { st.VpExtX = I32(d, o + 8); st.VpExtY = I32(d, o + 12); st.HasViewportExt = true; }
                    break;
                case 12: // SETVIEWPORTORGEX (PointL: x, y)
                    if (nSize >= 16) { st.VpOrgX = I32(d, o + 8); st.VpOrgY = I32(d, o + 12); }
                    break;

                case 19: // SETPOLYFILLMODE
                    if (nSize >= 12) st.FillRule = U32(d, o + 8) == 2 ? "nonzero" : "evenodd";
                    break;

                case 27: // MOVETOEX
                    if (nSize >= 16)
                    {
                        double x = I32(d, o + 8), y = I32(d, o + 12);
                        if (path != null)
                        {
                            var (ax, ay) = st.Apply(x, y);
                            canvas.Track(ax, ay);
                            path.Append(CultureInfo.InvariantCulture, $"M {F(ax)} {F(ay)} ");
                        }
                        st.X = x; st.Y = y;
                    }
                    break;

                case 54: // LINETO
                    if (nSize >= 16)
                    {
                        double x = I32(d, o + 8), y = I32(d, o + 12);
                        var (x1, y1) = st.Apply(st.X, st.Y);
                        var (x2, y2) = st.Apply(x, y);
                        canvas.Track(x1, y1);
                        canvas.Track(x2, y2);
                        if (path != null)
                        {
                            EnsurePathStart(path, st, canvas);
                            path.Append(CultureInfo.InvariantCulture, $"L {F(x2)} {F(y2)} ");
                        }
                        else
                            canvas.Add($"<line x1='{F(x1)}' y1='{F(y1)}' x2='{F(x2)}' y2='{F(y2)}' fill='none'{StrokeAttrs(st.Pen)}/>");
                        st.X = x; st.Y = y;
                    }
                    break;

                case 42: // ELLIPSE
                case 43: // RECTANGLE
                    if (nSize >= 24)
                        EmitBox(canvas, st, iType == 42, I32(d, o + 8), I32(d, o + 12), I32(d, o + 16), I32(d, o + 20), 0, 0);
                    break;

                case 44: // ROUNDRECT
                    if (nSize >= 32)
                        EmitBox(canvas, st, false, I32(d, o + 8), I32(d, o + 12), I32(d, o + 16), I32(d, o + 20),
                            I32(d, o + 24), I32(d, o + 28));
                    break;

                case 2:  // POLYBEZIER
                case 3:  // POLYGON
                case 4:  // POLYLINE
                case 5:  // POLYBEZIERTO
                case 6:  // POLYLINETO
                case 85: // POLYBEZIER16
                case 86: // POLYGON16
                case 87: // POLYLINE16
                case 88: // POLYBEZIERTO16
                case 89: // POLYLINETO16
                {
                    bool wide = iType <= 6;
                    var pts = ReadEmfPoly(d, o, nSize, wide);
                    if (pts == null) { skipped++; break; }
                    uint baseType = wide ? iType : iType - 83; // 85→2, 86→3, ...
                    EmitPoly(canvas, st, ref path, baseType, pts);
                    break;
                }

                case 8:  // POLYPOLYGON
                case 91: // POLYPOLYGON16
                {
                    bool wide = iType == 8;
                    EmitPolyPolygon(canvas, st, d, o, nSize, wide, ref skipped);
                    break;
                }

                case 35: // SETWORLDTRANSFORM
                    if (nSize >= 32)
                        st.SetTransform(F32(d, o + 8), F32(d, o + 12), F32(d, o + 16), F32(d, o + 20), F32(d, o + 24), F32(d, o + 28));
                    break;

                case 36: // MODIFYWORLDTRANSFORM
                    if (nSize >= 36)
                    {
                        uint mode = U32(d, o + 32);
                        double m11 = F32(d, o + 8), m12 = F32(d, o + 12), m21 = F32(d, o + 16),
                               m22 = F32(d, o + 20), dx = F32(d, o + 24), dy = F32(d, o + 28);
                        switch (mode)
                        {
                            case 1: st.SetTransform(1, 0, 0, 1, 0, 0); break;          // MWT_IDENTITY
                            case 2: st.Combine(m11, m12, m21, m22, dx, dy, false); break; // MWT_LEFTMULTIPLY
                            case 3: st.Combine(m11, m12, m21, m22, dx, dy, true); break;  // MWT_RIGHTMULTIPLY
                            case 4: st.SetTransform(m11, m12, m21, m22, dx, dy); break;   // MWT_SET
                        }
                    }
                    break;

                case 37: // SELECTOBJECT
                    if (nSize >= 12)
                    {
                        uint ih = U32(d, o + 8);
                        if ((ih & 0x80000000) != 0) ApplyStockObject(st, ih & 0x7FFFFFFF);
                        else if (objects.TryGetValue(ih, out var obj))
                        {
                            if (obj is PenDef p) st.Pen = p;
                            else if (obj is BrushDef b) st.Brush = b;
                        }
                    }
                    break;

                case 38: // CREATEPEN
                    if (nSize >= 28)
                        objects[U32(d, o + 8)] = MakePen(U32(d, o + 12), I32(d, o + 16), U32(d, o + 24));
                    break;

                case 95: // EXTCREATEPEN
                    if (nSize >= 44)
                        objects[U32(d, o + 8)] = MakePen(U32(d, o + 28), U32(d, o + 32), U32(d, o + 40));
                    break;

                case 39: // CREATEBRUSHINDIRECT
                    if (nSize >= 24)
                    {
                        uint style = U32(d, o + 12);
                        // BS_NULL/HOLLOW = 1 → brak wypełnienia; inne style (hatch/pattern) →
                        // przybliżenie kolorem bazowym (etap 1).
                        objects[U32(d, o + 8)] = style == 1
                            ? BrushDef.Null
                            : new BrushDef(ColorRef(U32(d, o + 16)), false);
                    }
                    break;

                case 40: // DELETEOBJECT
                    if (nSize >= 12) objects.Remove(U32(d, o + 8));
                    break;

                case 59: path = new StringBuilder(); break; // BEGINPATH
                case 60: break;                             // ENDPATH — ścieżka czeka na Fill/Stroke
                case 61: path?.Append("Z "); break;         // CLOSEFIGURE
                case 68: path = null; break;                // ABORTPATH

                case 62: // FILLPATH
                case 63: // STROKEANDFILLPATH
                case 64: // STROKEPATH
                    if (path is { Length: > 0 })
                    {
                        var fill = iType == 64 ? " fill='none'" : FillAttrs(st.Brush, st.FillRule);
                        var stroke = iType == 62 ? " stroke='none'" : StrokeAttrs(st.Pen);
                        canvas.Add($"<path d='{path.ToString().TrimEnd()}'{fill}{stroke}/>");
                    }
                    path = null;
                    break;

                case 81: // STRETCHDIBITS → <image> (DIB dekodowany istniejącą ścieżką Skia)
                    if (nSize >= 80)
                    {
                        double xDest = I32(d, o + 24), yDest = I32(d, o + 28);
                        uint offBmi = U32(d, o + 48), cbBmi = U32(d, o + 52),
                             offBits = U32(d, o + 56), cbBits = U32(d, o + 60);
                        double cxDest = I32(d, o + 72), cyDest = I32(d, o + 76);
                        var dib = GraphicConversionService.SliceDib(d, o, offBmi, cbBmi, offBits, cbBits, nSize);
                        var png = dib != null ? GraphicConversionService.DibToPng(dib) : null;
                        if (png != null && cxDest > 0 && cyDest > 0)
                        {
                            var (ax, ay) = st.Apply(xDest, yDest);
                            canvas.Track(ax, ay);
                            canvas.Track(ax + cxDest, ay + cyDest);
                            canvas.Add($"<image x='{F(ax)}' y='{F(ay)}' width='{F(cxDest)}' height='{F(cyDest)}' " +
                                       $"href='data:image/png;base64,{Convert.ToBase64String(png)}'/>");
                        }
                        else skipped++;
                    }
                    break;

                case 114: // ALPHABLEND → <image> (źródłowy DIB z alfą; SrcConstantAlpha → opacity).
                    // Logo Office często trzyma KOLOROWY raster w tym rekordzie, a wektorowe ścieżki
                    // są tylko czarną maską/cieniem — bez tego cała grafika wychodzi CZARNA.
                    if (nSize >= 108)
                    {
                        double xDest = I32(d, o + 24), yDest = I32(d, o + 28);
                        double cxDest = I32(d, o + 32), cyDest = I32(d, o + 36);
                        byte srcAlpha = d[o + 42];
                        uint offBmi = U32(d, o + 84), cbBmi = U32(d, o + 88),
                             offBits = U32(d, o + 92), cbBits = U32(d, o + 96);
                        var dib = GraphicConversionService.SliceDib(d, o, offBmi, cbBmi, offBits, cbBits, nSize);
                        var png = dib != null ? GraphicConversionService.DibToPng(dib) : null;
                        if (png != null && cxDest > 0 && cyDest > 0)
                        {
                            var (ax, ay) = st.Apply(xDest, yDest);
                            var (bx, by) = st.Apply(xDest + cxDest, yDest + cyDest);
                            double ix = Math.Min(ax, bx), iy = Math.Min(ay, by);
                            double iw = Math.Abs(bx - ax), ih = Math.Abs(by - ay);
                            canvas.Track(ix, iy);
                            canvas.Track(ix + iw, iy + ih);
                            var op = srcAlpha < 255 ? $" opacity='{F(srcAlpha / 255.0)}'" : string.Empty;
                            canvas.Add($"<image x='{F(ix)}' y='{F(iy)}' width='{F(iw)}' height='{F(ih)}' " +
                                       $"href='data:image/png;base64,{Convert.ToBase64String(png)}'{op}/>");
                        }
                        else skipped++;
                    }
                    break;

                default:
                    if (!silent.Contains(iType)) skipped++;
                    break;
            }

            o += (int)nSize;
        }

        return BuildSvg(canvas, haveVb, bLeft, bTop, bRight - bLeft, bBottom - bTop, widthPx, heightPx, skipped,
            recordCount: guard, usedWindowViewport: st.HasWindowExt && st.HasViewportExt);
    }

    /// <summary>Punkty rekordu poly EMF: licznik na +24, punkty od +28 (POINTL s32 lub POINTS16 s16).</summary>
    private static List<(double X, double Y)>? ReadEmfPoly(byte[] d, int o, uint nSize, bool wide)
    {
        if (nSize < 28) return null;
        uint count = U32(d, o + 24);
        int ptSize = wide ? 8 : 4;
        if (count == 0 || count > MaxPointsPerPoly || 28 + (long)count * ptSize > nSize) return null;

        var pts = new List<(double, double)>((int)count);
        int p = o + 28;
        for (uint i = 0; i < count; i++, p += ptSize)
        {
            if (wide) pts.Add((I32(d, p), I32(d, p + 4)));
            else pts.Add((I16(d, p), I16(d, p + 2)));
        }
        return pts;
    }

    /// <summary>Emisja poly wg bazowego typu EMF (2=bezier,3=polygon,4=polyline,5=bezierTo,6=lineTo).</summary>
    private static void EmitPoly(SvgCanvas canvas, GdiState st, ref StringBuilder? path,
        uint baseType, List<(double X, double Y)> pts)
    {
        switch (baseType)
        {
            case 3: // POLYGON
            {
                var sb = new StringBuilder();
                foreach (var (x, y) in pts)
                {
                    var (ax, ay) = st.Apply(x, y);
                    canvas.Track(ax, ay);
                    sb.Append(CultureInfo.InvariantCulture, $"{F(ax)},{F(ay)} ");
                }
                canvas.Add($"<polygon points='{sb.ToString().TrimEnd()}'{FillAttrs(st.Brush, st.FillRule)}{StrokeAttrs(st.Pen)}/>");
                break;
            }
            case 4: // POLYLINE
            {
                var sb = new StringBuilder();
                foreach (var (x, y) in pts)
                {
                    var (ax, ay) = st.Apply(x, y);
                    canvas.Track(ax, ay);
                    sb.Append(CultureInfo.InvariantCulture, $"{F(ax)},{F(ay)} ");
                }
                canvas.Add($"<polyline points='{sb.ToString().TrimEnd()}' fill='none'{StrokeAttrs(st.Pen)}/>");
                break;
            }
            case 2: // POLYBEZIER: p0 = start, dalej trójki punktów kontrolnych
            {
                if (pts.Count < 4) break;
                var sb = new StringBuilder();
                var (sx, sy) = st.Apply(pts[0].X, pts[0].Y);
                canvas.Track(sx, sy);
                sb.Append(CultureInfo.InvariantCulture, $"M {F(sx)} {F(sy)} ");
                AppendBezierTriples(sb, canvas, st, pts, 1);
                canvas.Add($"<path d='{sb.ToString().TrimEnd()}' fill='none'{StrokeAttrs(st.Pen)}/>");
                break;
            }
            case 5: // POLYBEZIERTO: start = bieżąca pozycja
            {
                if (pts.Count < 3) break;
                var (sx, sy) = st.Apply(st.X, st.Y);
                if (path != null)
                {
                    EnsurePathStart(path, st, canvas);
                    AppendBezierTriples(path, canvas, st, pts, 0);
                }
                else
                {
                    var sb = new StringBuilder();
                    canvas.Track(sx, sy);
                    sb.Append(CultureInfo.InvariantCulture, $"M {F(sx)} {F(sy)} ");
                    AppendBezierTriples(sb, canvas, st, pts, 0);
                    canvas.Add($"<path d='{sb.ToString().TrimEnd()}' fill='none'{StrokeAttrs(st.Pen)}/>");
                }
                (st.X, st.Y) = (pts[^1].X, pts[^1].Y);
                break;
            }
            case 6: // POLYLINETO: start = bieżąca pozycja
            {
                if (path != null)
                {
                    EnsurePathStart(path, st, canvas);
                    foreach (var (x, y) in pts)
                    {
                        var (ax, ay) = st.Apply(x, y);
                        canvas.Track(ax, ay);
                        path.Append(CultureInfo.InvariantCulture, $"L {F(ax)} {F(ay)} ");
                    }
                }
                else
                {
                    var sb = new StringBuilder();
                    var (sx, sy) = st.Apply(st.X, st.Y);
                    canvas.Track(sx, sy);
                    sb.Append(CultureInfo.InvariantCulture, $"{F(sx)},{F(sy)} ");
                    foreach (var (x, y) in pts)
                    {
                        var (ax, ay) = st.Apply(x, y);
                        canvas.Track(ax, ay);
                        sb.Append(CultureInfo.InvariantCulture, $"{F(ax)},{F(ay)} ");
                    }
                    canvas.Add($"<polyline points='{sb.ToString().TrimEnd()}' fill='none'{StrokeAttrs(st.Pen)}/>");
                }
                (st.X, st.Y) = (pts[^1].X, pts[^1].Y);
                break;
            }
        }
    }

    /// <summary>
    /// Gwarantuje, że ścieżka zaczyna się od moveto. GDI+ FillPath często ustawia bieżący
    /// punkt przez MOVETOEX PRZED BEGINPATH, a wewnątrz nawiasu używa wyłącznie wariantów
    /// „To" (PolyLineTo/PolyBezierTo). Bez wiodącego „M" ścieżka SVG jest niepoprawna i nic
    /// nie rysuje (cały metafile wychodził wtedy jako pusty blank — np. wektorowe logo).
    /// </summary>
    private static void EnsurePathStart(StringBuilder path, GdiState st, SvgCanvas canvas)
    {
        if (path.Length != 0) return;
        var (sx, sy) = st.Apply(st.X, st.Y);
        canvas.Track(sx, sy);
        path.Append(CultureInfo.InvariantCulture, $"M {F(sx)} {F(sy)} ");
    }

    private static void AppendBezierTriples(StringBuilder sb, SvgCanvas canvas, GdiState st,
        List<(double X, double Y)> pts, int from)
    {
        for (int i = from; i + 2 < pts.Count; i += 3)
        {
            var (c1x, c1y) = st.Apply(pts[i].X, pts[i].Y);
            var (c2x, c2y) = st.Apply(pts[i + 1].X, pts[i + 1].Y);
            var (ex, ey) = st.Apply(pts[i + 2].X, pts[i + 2].Y);
            canvas.Track(ex, ey);
            sb.Append(CultureInfo.InvariantCulture,
                $"C {F(c1x)} {F(c1y)} {F(c2x)} {F(c2y)} {F(ex)} {F(ey)} ");
        }
    }

    /// <summary>POLYPOLYGON(16): wielokąty jako jedna ścieżka z podścieżkami (respektuje fill-rule).</summary>
    private static void EmitPolyPolygon(SvgCanvas canvas, GdiState st, byte[] d, int o, uint nSize,
        bool wide, ref int skipped)
    {
        if (nSize < 32) { skipped++; return; }
        uint nPolys = U32(d, o + 24);
        uint total = U32(d, o + 28);
        int ptSize = wide ? 8 : 4;
        if (nPolys == 0 || nPolys > 10_000 || total == 0 || total > MaxPointsPerPoly
            || 32 + (long)nPolys * 4 + (long)total * ptSize > nSize)
        {
            skipped++;
            return;
        }

        var sb = new StringBuilder();
        int countsAt = o + 32;
        int p = countsAt + (int)nPolys * 4;
        uint consumed = 0;
        for (uint poly = 0; poly < nPolys; poly++)
        {
            uint cnt = U32(d, countsAt + (int)poly * 4);
            if (cnt == 0 || consumed + cnt > total) { skipped++; return; }
            for (uint i = 0; i < cnt; i++, p += ptSize)
            {
                double x = wide ? I32(d, p) : I16(d, p);
                double y = wide ? I32(d, p + 4) : I16(d, p + 2);
                var (ax, ay) = st.Apply(x, y);
                canvas.Track(ax, ay);
                sb.Append(CultureInfo.InvariantCulture, $"{(i == 0 ? "M" : "L")} {F(ax)} {F(ay)} ");
            }
            sb.Append("Z ");
            consumed += cnt;
        }
        canvas.Add($"<path d='{sb.ToString().TrimEnd()}'{FillAttrs(st.Brush, st.FillRule)}{StrokeAttrs(st.Pen)}/>");
    }

    // ---- WMF -------------------------------------------------------------------------------

    private static MetafileSvg? TranslateWmf(byte[] d, int widthPx, int heightPx)
    {
        int o = 0;
        double vbL = 0, vbT = 0, vbW = 0, vbH = 0;
        bool haveVb = false;

        if (d.Length >= 22 && U32(d, 0) == 0x9AC6CDD7) // Aldus placeable
        {
            double l = I16(d, 6), t = I16(d, 8), r = I16(d, 10), b = I16(d, 12);
            if (r > l && b > t) { vbL = l; vbT = t; vbW = r - l; vbH = b - t; haveVb = true; }
            o = 22;
        }

        if (o + 18 > d.Length) return null;
        ushort type = U16(d, o);
        ushort hdrWords = U16(d, o + 2);
        if ((type != 1 && type != 2) || hdrWords != 9) return null;
        o += 18;

        var st = new GdiState(); // WMF16 nie ma world transform — identyczność
        var slots = new List<object?>();
        var canvas = new SvgCanvas();
        int skipped = 0;
        var placeholder = new object(); // fonty/palety/regiony zajmują sloty tabeli obiektów

        var silent = new HashSet<ushort>
        {
            0x0103, 0x0102, 0x0201, 0x0209, 0x0104, 0x0107, // SETMAPMODE/BKMODE/BKCOLOR/TEXTCOLOR/ROP2/STRETCHBLTMODE
            0x001E, 0x0127, 0x020D, 0x020E, 0x012E          // SAVEDC, RESTOREDC, SETVIEWPORTORG/EXT, SETTEXTALIGN
        };

        bool done = false;
        int guard = 0;
        while (!done && o + 6 <= d.Length && guard++ < MaxRecords && !canvas.Full)
        {
            uint sizeWords = U32(d, o);
            ushort func = U16(d, o + 4);
            long byteSize = (long)sizeWords * 2;
            if (sizeWords < 3 || o + byteSize > d.Length) break;
            int p = o + 6;
            long paramBytes = byteSize - 6;

            switch (func)
            {
                case 0x0000: done = true; break; // META_EOF

                case 0x020B: // SETWINDOWORG: [Y, X]
                    if (paramBytes >= 4) { vbT = I16(d, p); vbL = I16(d, p + 2); haveVb = vbW > 0 && vbH > 0; }
                    break;
                case 0x020C: // SETWINDOWEXT: [Y, X]
                    if (paramBytes >= 4)
                    {
                        vbH = Math.Abs((double)I16(d, p));
                        vbW = Math.Abs((double)I16(d, p + 2));
                        haveVb = vbW > 0 && vbH > 0;
                    }
                    break;

                case 0x0214: // MOVETO: [Y, X]
                    if (paramBytes >= 4) { st.Y = I16(d, p); st.X = I16(d, p + 2); }
                    break;

                case 0x0213: // LINETO: [Y, X]
                    if (paramBytes >= 4)
                    {
                        double y = I16(d, p), x = I16(d, p + 2);
                        canvas.Track(st.X, st.Y);
                        canvas.Track(x, y);
                        canvas.Add($"<line x1='{F(st.X)}' y1='{F(st.Y)}' x2='{F(x)}' y2='{F(y)}' fill='none'{StrokeAttrs(st.Pen)}/>");
                        st.X = x; st.Y = y;
                    }
                    break;

                case 0x041B: // RECTANGLE: [B, R, T, L]
                case 0x0418: // ELLIPSE: [B, R, T, L]
                    if (paramBytes >= 8)
                        EmitBox(canvas, st, func == 0x0418,
                            I16(d, p + 6), I16(d, p + 4), I16(d, p + 2), I16(d, p), 0, 0);
                    break;

                case 0x061C: // ROUNDRECT: [H, W, B, R, T, L]
                    if (paramBytes >= 12)
                        EmitBox(canvas, st, false,
                            I16(d, p + 10), I16(d, p + 8), I16(d, p + 6), I16(d, p + 4),
                            I16(d, p + 2), I16(d, p));
                    break;

                case 0x0324: // POLYGON: [count, x1,y1, ...]
                case 0x0325: // POLYLINE
                    if (paramBytes >= 2)
                    {
                        int count = U16(d, p);
                        if (count <= 0 || count > MaxPointsPerPoly || 2 + (long)count * 4 > paramBytes) { skipped++; break; }
                        var sb = new StringBuilder();
                        for (int i = 0; i < count; i++)
                        {
                            double x = I16(d, p + 2 + i * 4), y = I16(d, p + 4 + i * 4);
                            canvas.Track(x, y);
                            sb.Append(CultureInfo.InvariantCulture, $"{F(x)},{F(y)} ");
                        }
                        canvas.Add(func == 0x0324
                            ? $"<polygon points='{sb.ToString().TrimEnd()}'{FillAttrs(st.Brush, st.FillRule)}{StrokeAttrs(st.Pen)}/>"
                            : $"<polyline points='{sb.ToString().TrimEnd()}' fill='none'{StrokeAttrs(st.Pen)}/>");
                    }
                    break;

                case 0x0538: // POLYPOLYGON: [nPolys, counts..., points...]
                    if (paramBytes >= 2)
                    {
                        int nPolys = U16(d, p);
                        if (nPolys <= 0 || nPolys > 10_000 || 2 + (long)nPolys * 2 > paramBytes) { skipped++; break; }
                        var sb = new StringBuilder();
                        int pt = p + 2 + nPolys * 2;
                        bool ok = true;
                        for (int poly = 0; poly < nPolys && ok; poly++)
                        {
                            int cnt = U16(d, p + 2 + poly * 2);
                            if (cnt <= 0 || pt - o + (long)cnt * 4 > byteSize) { ok = false; break; }
                            for (int i = 0; i < cnt; i++, pt += 4)
                            {
                                double x = I16(d, pt), y = I16(d, pt + 2);
                                canvas.Track(x, y);
                                sb.Append(CultureInfo.InvariantCulture, $"{(i == 0 ? "M" : "L")} {F(x)} {F(y)} ");
                            }
                            sb.Append("Z ");
                        }
                        if (ok)
                            canvas.Add($"<path d='{sb.ToString().TrimEnd()}'{FillAttrs(st.Brush, st.FillRule)}{StrokeAttrs(st.Pen)}/>");
                        else skipped++;
                    }
                    break;

                case 0x0106: // SETPOLYFILLMODE
                    if (paramBytes >= 2) st.FillRule = U16(d, p) == 2 ? "nonzero" : "evenodd";
                    break;

                case 0x02FA: // CREATEPENINDIRECT: style, width.x, width.y, colorref
                    if (paramBytes >= 10)
                        AddSlot(slots, MakePen(U16(d, p), I16(d, p + 2), U32(d, p + 6)));
                    else AddSlot(slots, placeholder);
                    break;

                case 0x02FC: // CREATEBRUSHINDIRECT: style, colorref, hatch
                    if (paramBytes >= 8)
                    {
                        ushort style = U16(d, p);
                        AddSlot(slots, style == 1 ? BrushDef.Null : new BrushDef(ColorRef(U32(d, p + 2)), false));
                    }
                    else AddSlot(slots, placeholder);
                    break;

                // Te obiekty zajmują slot tabeli (inaczej indeksy SELECTOBJECT się rozjadą):
                case 0x02FB: // CREATEFONTINDIRECT
                case 0x00F7: // CREATEPALETTE
                case 0x01F9: // CREATEPATTERNBRUSH
                case 0x0142: // DIBCREATEPATTERNBRUSH
                case 0x06FF: // CREATEREGION
                    AddSlot(slots, placeholder);
                    break;

                case 0x012D: // SELECTOBJECT
                    if (paramBytes >= 2)
                    {
                        int idx = U16(d, p);
                        if (idx >= 0 && idx < slots.Count)
                        {
                            if (slots[idx] is PenDef pen) st.Pen = pen;
                            else if (slots[idx] is BrushDef brush) st.Brush = brush;
                        }
                    }
                    break;

                case 0x01F0: // DELETEOBJECT
                    if (paramBytes >= 2)
                    {
                        int idx = U16(d, p);
                        if (idx >= 0 && idx < slots.Count) slots[idx] = null;
                    }
                    break;

                case 0x0F43: // STRETCHDIB → <image>
                    if (paramBytes >= 22 + 40)
                    {
                        double destH = I16(d, p + 14), destW = I16(d, p + 16);
                        double yDst = I16(d, p + 18), xDst = I16(d, p + 20);
                        var dibLen = (int)(paramBytes - 22);
                        var dib = new byte[dibLen];
                        Array.Copy(d, p + 22, dib, 0, dibLen);
                        var png = GraphicConversionService.DibToPng(dib);
                        if (png != null && destW > 0 && destH > 0)
                        {
                            canvas.Track(xDst, yDst);
                            canvas.Track(xDst + destW, yDst + destH);
                            canvas.Add($"<image x='{F(xDst)}' y='{F(yDst)}' width='{F(destW)}' height='{F(destH)}' " +
                                       $"href='data:image/png;base64,{Convert.ToBase64String(png)}'/>");
                        }
                        else skipped++;
                    }
                    break;

                default:
                    if (!silent.Contains(func)) skipped++;
                    break;
            }

            o += (int)byteSize;
        }

        return BuildSvg(canvas, haveVb, vbL, vbT, vbW, vbH, widthPx, heightPx, skipped, recordCount: guard);
    }

    private static void AddSlot(List<object?> slots, object obj)
    {
        var i = slots.IndexOf(null);
        if (i >= 0) slots[i] = obj;
        else slots.Add(obj);
    }

    // ---- wspólna emisja ---------------------------------------------------------------------

    /// <summary>Rectangle/Ellipse/RoundRect z uwzględnieniem transformacji (rotacja → polygon).</summary>
    private static void EmitBox(SvgCanvas canvas, GdiState st, bool ellipse,
        double l, double t, double r, double b, double cornerW, double cornerH)
    {
        if (r < l) (l, r) = (r, l);
        if (b < t) (t, b) = (b, t);
        if (r - l <= 0 || b - t <= 0) return;

        var fill = FillAttrs(st.Brush, st.FillRule);
        var stroke = StrokeAttrs(st.Pen);

        if (st.IsAxisAligned)
        {
            var (x1, y1) = st.Apply(l, t);
            var (x2, y2) = st.Apply(r, b);
            if (x2 < x1) (x1, x2) = (x2, x1);
            if (y2 < y1) (y1, y2) = (y2, y1);
            canvas.Track(x1, y1);
            canvas.Track(x2, y2);
            if (ellipse)
            {
                canvas.Add($"<ellipse cx='{F((x1 + x2) / 2)}' cy='{F((y1 + y2) / 2)}' " +
                           $"rx='{F((x2 - x1) / 2)}' ry='{F((y2 - y1) / 2)}'{fill}{stroke}/>");
            }
            else
            {
                var rAttrs = cornerW > 0 || cornerH > 0
                    ? $" rx='{F(Math.Abs(cornerW) / 2)}' ry='{F(Math.Abs(cornerH) / 2)}'"
                    : string.Empty;
                canvas.Add($"<rect x='{F(x1)}' y='{F(y1)}' width='{F(x2 - x1)}' height='{F(y2 - y1)}'{rAttrs}{fill}{stroke}/>");
            }
        }
        else
        {
            // Rotacja/skos: prostokąt jako polygon 4 narożników; elipsa — przybliżenie bez rotacji
            // (etap 1; dokładna rotacja elipsy wymaga transform-attrybutu — do rozważenia w etapie 2).
            var corners = new[] { st.Apply(l, t), st.Apply(r, t), st.Apply(r, b), st.Apply(l, b) };
            foreach (var (x, y) in corners) canvas.Track(x, y);
            if (ellipse)
            {
                double cx = corners.Average(c => c.X), cy = corners.Average(c => c.Y);
                double rx = (corners.Max(c => c.X) - corners.Min(c => c.X)) / 2;
                double ry = (corners.Max(c => c.Y) - corners.Min(c => c.Y)) / 2;
                canvas.Add($"<ellipse cx='{F(cx)}' cy='{F(cy)}' rx='{F(rx)}' ry='{F(ry)}'{fill}{stroke}/>");
            }
            else
            {
                var sb = new StringBuilder();
                foreach (var (x, y) in corners)
                    sb.Append(CultureInfo.InvariantCulture, $"{F(x)},{F(y)} ");
                canvas.Add($"<polygon points='{sb.ToString().TrimEnd()}'{fill}{stroke}/>");
            }
        }
    }

    private static MetafileSvg? BuildSvg(SvgCanvas canvas, bool haveVb,
        double vbL, double vbT, double vbW, double vbH, int widthPx, int heightPx, int skipped,
        int recordCount = 0, bool usedWindowViewport = false)
    {
        if (!canvas.HasContent) return null;

        // rclBounds z nagłówka jest w jednostkach URZĄDZENIA. Jeśli mimo mapowania window/viewport
        // treść i tak nie pokrywa się z tym prostokątem (np. nieobsłużony tryb mapowania metrycznego),
        // użyj bounding boxa realnie narysowanej treści — inaczej grafika wyszłaby pusta.
        bool contentIntersectsBounds = haveVb && vbW > 0 && vbH > 0
            && canvas.MaxX >= vbL && canvas.MinX <= vbL + vbW
            && canvas.MaxY >= vbT && canvas.MinY <= vbT + vbH;
        bool contentOutsideBounds = haveVb && !contentIntersectsBounds;

        if (!contentIntersectsBounds)
        {
            // viewBox z bounding boxa realnie narysowanej treści (padding 1 j.).
            vbL = canvas.MinX - 1;
            vbT = canvas.MinY - 1;
            vbW = Math.Max(1, canvas.MaxX - canvas.MinX + 2);
            vbH = Math.Max(1, canvas.MaxY - canvas.MinY + 2);
        }
        if (widthPx <= 0) widthPx = (int)Math.Clamp(vbW, 1, 2000);
        if (heightPx <= 0) heightPx = (int)Math.Clamp(vbH, 1, 2000);

        var svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='{widthPx}' height='{heightPx}' " +
                  $"viewBox='{F(vbL)} {F(vbT)} {F(vbW)} {F(vbH)}' preserveAspectRatio='xMidYMid meet'>" +
                  canvas.Elements + "</svg>";
        if (svg.Length > MaxOutputChars + 512) return null;
        var elements = canvas.Elements;
        var fillColors = System.Text.RegularExpressions.Regex.Matches(elements, "fill='(#[0-9a-fA-F]{6})'")
            .Select(m => m.Groups[1].Value.ToLowerInvariant()).Distinct().ToList();
        return new MetafileSvg
        {
            Svg = svg,
            SkippedRecords = skipped,
            RecordCount = recordCount,
            UsedWindowViewport = usedWindowViewport,
            ContentOutsideDeviceBounds = contentOutsideBounds,
            FillColors = fillColors,
            HasEmbeddedImage = elements.Contains("<image ", StringComparison.Ordinal)
        };
    }

    private static void ApplyStockObject(GdiState st, uint index)
    {
        switch (index)
        {
            case 0: st.Brush = new BrushDef("#ffffff", false); break; // WHITE_BRUSH
            case 1: st.Brush = new BrushDef("#c0c0c0", false); break; // LTGRAY_BRUSH
            case 2: st.Brush = new BrushDef("#808080", false); break; // GRAY_BRUSH
            case 3: st.Brush = new BrushDef("#404040", false); break; // DKGRAY_BRUSH
            case 4: st.Brush = new BrushDef("#000000", false); break; // BLACK_BRUSH
            case 5: st.Brush = BrushDef.Null; break;                  // NULL_BRUSH
            case 6: st.Pen = new PenDef("#ffffff", 1, null, false); break; // WHITE_PEN
            case 7: st.Pen = PenDef.Default; break;                   // BLACK_PEN
            case 8: st.Pen = PenDef.Null; break;                      // NULL_PEN
            // fonty/palety systemowe — bez znaczenia w etapie 1
        }
    }

    private static PenDef MakePen(uint style, double width, uint colorRef)
    {
        var s = style & 0xF; // PS_STYLE_MASK
        if (s == 5) return PenDef.Null; // PS_NULL
        string? dash = s switch
        {
            1 => "6 3",         // PS_DASH
            2 => "2 2",         // PS_DOT
            3 => "6 3 2 3",     // PS_DASHDOT
            4 => "6 3 2 3 2 3", // PS_DASHDOTDOT
            _ => null
        };
        return new PenDef(ColorRef(colorRef), Math.Max(1, Math.Abs(width)), dash, false);
    }

    private static string StrokeAttrs(in PenDef p) => p.IsNull
        ? " stroke='none'"
        : $" stroke='{p.Color}' stroke-width='{F(p.Width)}'" +
          (p.Dash != null ? $" stroke-dasharray='{p.Dash}'" : string.Empty);

    private static string FillAttrs(in BrushDef b, string fillRule) => b.IsNull
        ? " fill='none'"
        : $" fill='{b.Color}' fill-rule='{fillRule}'";

    /// <summary>COLORREF 0x00BBGGRR → #rrggbb.</summary>
    private static string ColorRef(uint c) =>
        $"#{(byte)c:x2}{(byte)(c >> 8):x2}{(byte)(c >> 16):x2}";

    private static string F(double v) =>
        Math.Round(v, 2).ToString("0.##", CultureInfo.InvariantCulture);

    private static uint U32(byte[] d, int o) => BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o));
    private static int I32(byte[] d, int o) => BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(o));
    private static ushort U16(byte[] d, int o) => BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(o));
    private static short I16(byte[] d, int o) => BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(o));
    private static float F32(byte[] d, int o) => BinaryPrimitives.ReadSingleLittleEndian(d.AsSpan(o));
}

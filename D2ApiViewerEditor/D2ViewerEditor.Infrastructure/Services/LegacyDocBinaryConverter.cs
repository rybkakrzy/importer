using System.Buffers.Binary;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Konwerter starszego, binarnego formatu Worda (.doc / OLE-CFBF, MS-DOC) do .docx — czysto
/// zarządzany, bez LibreOffice/GDI/System.Drawing (Linux/GCP-safe). Parsuje FIB (File Information
/// Block) i piece table (CLX/PlcPcd) ze strumieni WordDocument + (0Table|1Table), wyciąga tekst
/// dokumentu oraz bezpośrednie formatowanie znaków (CHPX) i buduje poprawny DOCX przez OpenXML.
///
/// Zakres formatowania znaków (świadomy): odzyskiwane jest bezpośrednie formatowanie run-ów —
/// pogrubienie, kursywa, podkreślenie, przekreślenie i kolor tekstu — przez parsowanie warstwy
/// CHPX (PlcfBteChpx → CHPX FKP → SPRM-y) formatu MS-DOC. Nie odtwarzamy stylów akapitowych/znakowych,
/// tabel, obrazów ani nagłówków/stopek — binarny .doc to format sprzed OOXML i pełna wierność
/// wymagałaby dużego parsera. Gdy warstwa CHPX jest niespójna, degradujemy się do samego tekstu
/// (formatowanie domyślne) zamiast wywracać import. Gdy sam FIB/CLX jest niespójny, metoda zwraca
/// null, a wołający stosuje kontrolowane odrzucenie (bez śmieci).
/// </summary>
public static class LegacyDocBinaryConverter
{
    private const ushort WordDocumentMagic = 0xA5EC;     // wIdent w FIB
    private const int FibFlagsOffset = 0x000A;           // bit 0x0200 = fWhichTblStm (0Table/1Table)
    private const int FcClxOffset = 0x01A2;              // fcClx w FibRgFcLcb97
    private const int LcbClxOffset = 0x01A6;             // lcbClx
    private const int FcPlcfBteChpxOffset = 0x00FA;      // fcPlcfBteChpx w FibRgFcLcb97 (para 12)
    private const int LcbPlcfBteChpxOffset = 0x00FE;     // lcbPlcfBteChpx
    private const int FcSttbfFfnOffset = 0x0112;         // fcSttbfFfn w FibRgFcLcb97 (para 15) — tablica fontów
    private const int LcbSttbfFfnOffset = 0x0116;        // lcbSttbfFfn
    private const int FfnNameOffset = 40;                // xszFfn: 1(cbFfnM1)+1(flags)+2(wWeight)+1(chs)+1(ixchSzAlt)+10(panose)+24(fs)
    private const int MaxFonts = 4096;                   // sanity limit wpisów SttbfFfn
    private const ushort FWhichTblStmBit = 0x0200;
    private const uint FcCompressedFlag = 0x40000000;    // PCD.fc: 1 bajt/znak (CP1252) zamiast UTF-16
    private const uint FcValueMask = 0x3FFFFFFF;
    private const int FkpSize = 512;                     // rozmiar strony FKP (Formatted disk PaGe)
    private const uint PnFkpMask = 0x003FFFFF;           // dolne 22 bity PnFkpChpx = numer strony FKP
    private const int MaxChars = 8_000_000;              // twardy limit (anty-abuse)

    /// <summary>
    /// Próbuje przekonwertować binarny .doc do DOCX. Zwraca bajty DOCX albo null, gdy nie da się
    /// bezpiecznie odczytać (uszkodzony FIB/CLX, brak tekstu, nieobsługiwany wariant).
    /// </summary>
    public static byte[]? TryConvert(byte[] wordDocument, byte[]? table0Stream, byte[]? table1Stream)
    {
        var paragraphs = ExtractParagraphs(wordDocument, table0Stream, table1Stream);
        if (paragraphs == null) return null;
        // Wymagamy realnej treści — pusty dokument lub same znaki sterujące → kontrolowane odrzucenie.
        bool hasText = paragraphs.Any(p => p.Any(r => !string.IsNullOrWhiteSpace(r.Text)));
        if (!hasText) return null;
        return BuildDocx(paragraphs);
    }

    private static List<List<TextRun>>? ExtractParagraphs(byte[] wd, byte[]? table0, byte[]? table1)
    {
        if (wd == null || wd.Length < LcbClxOffset + 4) return null;
        if (BinaryPrimitives.ReadUInt16LittleEndian(wd) != WordDocumentMagic) return null;

        ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(wd.AsSpan(FibFlagsOffset));
        bool useTable1 = (flags & FWhichTblStmBit) != 0;
        var table = useTable1 ? table1 : table0;
        if (table == null || table.Length == 0) return null;

        int fcClx = BinaryPrimitives.ReadInt32LittleEndian(wd.AsSpan(FcClxOffset));
        uint lcbClx = BinaryPrimitives.ReadUInt32LittleEndian(wd.AsSpan(LcbClxOffset));
        if (fcClx < 0 || lcbClx == 0 || (long)fcClx + lcbClx > table.Length) return null;

        var clx = table.AsSpan(fcClx, (int)lcbClx);
        var plcPcd = FindPlcPcd(clx);
        if (plcPcd == null) return null;

        // Indeks formatowania znaków (CHPX). Gdy niespójny → null i degradacja do samego tekstu.
        var chpx = BuildChpxIndex(wd, table);

        return ReadPieces(wd, plcPcd, chpx);
    }

    /// <summary>Pomija ewentualne Prc (clxt=1) i zwraca blok PlcPcd z rekordu Pcdt (clxt=2).</summary>
    private static byte[]? FindPlcPcd(ReadOnlySpan<byte> clx)
    {
        int pos = 0;
        while (pos < clx.Length)
        {
            byte clxt = clx[pos];
            if (clxt == 0x01) // Prc — grpprl z prefiksem długości (UInt16)
            {
                if (pos + 3 > clx.Length) return null;
                ushort cb = BinaryPrimitives.ReadUInt16LittleEndian(clx.Slice(pos + 1));
                pos += 3 + cb;
            }
            else if (clxt == 0x02) // Pcdt — lcb (UInt32) + PlcPcd
            {
                if (pos + 5 > clx.Length) return null;
                uint lcb = BinaryPrimitives.ReadUInt32LittleEndian(clx.Slice(pos + 1));
                int start = pos + 5;
                if (lcb == 0 || (long)start + lcb > clx.Length) return null;
                return clx.Slice(start, (int)lcb).ToArray();
            }
            else return null; // nieznany wpis CLX
        }
        return null;
    }

    /// <summary>
    /// Odczytuje tekst wg piece table: PlcPcd = (n+1) pozycji znakowych (CP, UInt32) + n PCD (8 B).
    /// PCD niesie fc (FcCompressed) — bit 0x40000000 ⇒ 1 bajt/znak (CP1252/Latin1), wpp 16-bit Unicode.
    /// Dla każdego znaku liczymy jego offset bajtowy w strumieniu WordDocument i pytamy indeks CHPX
    /// o bezpośrednie formatowanie; kolejne znaki o identycznym formacie łączą się w jeden run.
    /// </summary>
    private static List<List<TextRun>>? ReadPieces(byte[] wd, byte[] plcPcd, ChpxIndex? chpx)
    {
        // lcb = 4*(n+1) + 8*n  ⇒  n = (lcb - 4) / 12
        if (plcPcd.Length < 4 + 8) return null;
        int n = (plcPcd.Length - 4) / 12;
        if (n <= 0) return null;

        int cpBase = 0;
        int pcdBase = 4 * (n + 1);

        var builder = new ParagraphBuilder();

        for (int i = 0; i < n; i++)
        {
            uint cpStart = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(cpBase + i * 4));
            uint cpEnd = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(cpBase + (i + 1) * 4));
            if (cpEnd <= cpStart) continue;
            long cch = cpEnd - cpStart;
            if (cch > MaxChars || builder.TotalChars + cch > MaxChars) return null;

            uint fc = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(pcdBase + i * 8 + 2));
            bool compressed = (fc & FcCompressedFlag) != 0;
            uint fcValue = fc & FcValueMask;

            string chunk;
            long byteBase;
            int byteStride;
            if (compressed)
            {
                byteBase = fcValue / 2;
                byteStride = 1;
                if (byteBase < 0 || byteBase + cch > wd.Length) return null;
                chunk = Encoding.Latin1.GetString(wd, (int)byteBase, (int)cch);
            }
            else
            {
                byteBase = fcValue;
                byteStride = 2;
                if (byteBase < 0 || byteBase + cch * 2 > wd.Length) return null;
                chunk = Encoding.Unicode.GetString(wd, (int)byteBase, (int)(cch * 2));
            }

            for (int j = 0; j < chunk.Length; j++)
            {
                long off = byteBase + (long)j * byteStride;
                builder.Append(chunk[j], chpx?.GetFormat(off) ?? default);
            }
        }

        return builder.Finish();
    }

    private static byte[] BuildDocx(List<List<TextRun>> paragraphs)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            var body = main.Document.AppendChild(new Body());

            foreach (var runs in paragraphs)
            {
                var p = new Paragraph();
                foreach (var run in runs)
                {
                    if (run.Text.Length == 0) continue;
                    var r = new Run();
                    var rpr = BuildRunProperties(run.Format);
                    if (rpr != null) r.AppendChild(rpr);
                    r.AppendChild(new Text(run.Text) { Space = SpaceProcessingModeValues.Preserve });
                    p.AppendChild(r);
                }
                body.AppendChild(p);
            }

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>Buduje w:rPr w kolejności schematu CT_RPr (rFonts, b, i, strike, color, sz, u); null gdy brak formatowania.</summary>
    private static RunProperties? BuildRunProperties(CharFormat f)
    {
        if (!f.HasAny) return null;
        var rpr = new RunProperties();
        if (f.FontName != null)
            rpr.AppendChild(new RunFonts { Ascii = f.FontName, HighAnsi = f.FontName });
        if (f.Bold) rpr.AppendChild(new Bold());
        if (f.Italic) rpr.AppendChild(new Italic());
        if (f.Strike) rpr.AppendChild(new Strike());
        if (f.ColorHex != null) rpr.AppendChild(new Color { Val = f.ColorHex });
        if (f.SizeHalfPoints is > 0)
            rpr.AppendChild(new FontSize { Val = f.SizeHalfPoints.Value.ToString() });
        if (f.Underline) rpr.AppendChild(new Underline { Val = UnderlineValues.Single });
        return rpr;
    }

    // ---- warstwa formatowania znaków (CHPX) --------------------------------------

    /// <summary>Bezpośrednie formatowanie run-a wyciągnięte z CHPX (podzbiór wspierany w podglądzie).</summary>
    private readonly record struct CharFormat(
        bool Bold, bool Italic, bool Underline, bool Strike, string? ColorHex,
        ushort? SizeHalfPoints, string? FontName)
    {
        public bool HasAny =>
            Bold || Italic || Underline || Strike || ColorHex != null
            || SizeHalfPoints != null || FontName != null;
    }

    private sealed class TextRun
    {
        public required string Text { get; init; }
        public required CharFormat Format { get; init; }
    }

    /// <summary>Akumulator akapitów: łączy sąsiednie znaki o tym samym formacie w run-y, dzieli na akapity po 0x0D.</summary>
    private sealed class ParagraphBuilder
    {
        private readonly List<List<TextRun>> _paragraphs = new();
        private List<TextRun> _current = new();
        private readonly StringBuilder _runText = new();
        private CharFormat _runFormat;
        private bool _runOpen;

        public long TotalChars { get; private set; }

        public void Append(char c, CharFormat format)
        {
            if (c == '\r' || c == '\n') { EndParagraph(); return; }

            char outChar;
            if (c == '\t') outChar = '\t';
            else if (c >= ' ') outChar = c;
            else return; // cell mark 0x07, pola 0x13/0x14/0x15, obraz 0x01, NUL … — pomijamy

            if (!_runOpen || !format.Equals(_runFormat))
            {
                FlushRun();
                _runFormat = format;
                _runOpen = true;
            }
            _runText.Append(outChar);
            TotalChars++;
        }

        private void FlushRun()
        {
            if (_runOpen && _runText.Length > 0)
                _current.Add(new TextRun { Text = _runText.ToString(), Format = _runFormat });
            _runText.Clear();
            _runOpen = false;
        }

        private void EndParagraph()
        {
            FlushRun();
            _paragraphs.Add(_current);
            _current = new List<TextRun>();
        }

        public List<List<TextRun>> Finish()
        {
            EndParagraph(); // domyka ostatni (Word kończy akapit CR, ale zabezpieczamy resztkę bez CR)
            return _paragraphs;
        }
    }

    /// <summary>
    /// Buduje indeks CHPX z PlcfBteChpx (w strumieniu tablicy) → PnFkpChpx (numery stron FKP w
    /// strumieniu WordDocument). Zwraca null, gdy struktura jest niespójna (degradacja do samego tekstu).
    /// </summary>
    private static ChpxIndex? BuildChpxIndex(byte[] wd, byte[] table)
    {
        try
        {
            if (wd.Length < LcbPlcfBteChpxOffset + 4) return null;
            int fcPlcf = BinaryPrimitives.ReadInt32LittleEndian(wd.AsSpan(FcPlcfBteChpxOffset));
            uint lcbPlcf = BinaryPrimitives.ReadUInt32LittleEndian(wd.AsSpan(LcbPlcfBteChpxOffset));
            if (fcPlcf < 0 || lcbPlcf < 4 + 4 || (long)fcPlcf + lcbPlcf > table.Length) return null;

            // PlcfBteChpx = (n+1) FC (UInt32) + n PnFkpChpx (UInt32);  lcb = 4*(n+1) + 4*n = 8n + 4
            int n = ((int)lcbPlcf - 4) / 8;
            if (n <= 0) return null;

            var plcf = table.AsSpan(fcPlcf, (int)lcbPlcf);
            var boundaries = new long[n + 1];
            for (int i = 0; i <= n; i++)
                boundaries[i] = BinaryPrimitives.ReadUInt32LittleEndian(plcf.Slice(i * 4));

            var pages = new int[n];
            int pnBase = 4 * (n + 1);
            for (int i = 0; i < n; i++)
            {
                uint pnRaw = BinaryPrimitives.ReadUInt32LittleEndian(plcf.Slice(pnBase + i * 4));
                pages[i] = (int)(pnRaw & PnFkpMask);
            }

            return new ChpxIndex(wd, boundaries, pages, ParseSttbfFfn(wd, table));
        }
        catch
        {
            return null; // dowolna niespójność → samodegradacja (tekst bez formatowania)
        }
    }

    /// <summary>
    /// Parsuje tablicę fontów SttbfFfn (strumień tablicy): cData, cbExtra, potem wpisy FFN —
    /// nazwa fontu to UTF-16 xszFfn od bajtu 40 wpisu. Indeks listy = ftc z sprmCRgFtc0.
    /// Dowolna niespójność → null (run-y zostają bez w:rFonts, reszta formatowania przeżywa).
    /// </summary>
    private static IReadOnlyList<string>? ParseSttbfFfn(byte[] wd, byte[] table)
    {
        try
        {
            if (wd.Length < LcbSttbfFfnOffset + 4) return null;
            int fc = BinaryPrimitives.ReadInt32LittleEndian(wd.AsSpan(FcSttbfFfnOffset));
            uint lcb = BinaryPrimitives.ReadUInt32LittleEndian(wd.AsSpan(LcbSttbfFfnOffset));
            if (fc < 0 || lcb < 4 || (long)fc + lcb > table.Length) return null;

            var span = table.AsSpan(fc, (int)lcb);
            int cData = BinaryPrimitives.ReadUInt16LittleEndian(span);
            int cbExtra = BinaryPrimitives.ReadUInt16LittleEndian(span.Slice(2));
            if (cData <= 0 || cData > MaxFonts) return null;

            var fonts = new List<string>(cData);
            int pos = 4;
            for (int i = 0; i < cData && pos < span.Length; i++)
            {
                int entryLen = span[pos] + 1; // cbFfnM1 + 1
                if (entryLen <= 1 || pos + entryLen > span.Length) break;
                string name = "";
                int nameStart = pos + FfnNameOffset;
                int nameBytes = pos + entryLen - nameStart;
                if (nameBytes >= 2)
                {
                    var raw = span.Slice(nameStart, nameBytes);
                    int end = 0;
                    while (end + 1 < raw.Length
                           && BinaryPrimitives.ReadUInt16LittleEndian(raw.Slice(end)) != 0)
                        end += 2;
                    name = Encoding.Unicode.GetString(raw.Slice(0, end)).Trim();
                }
                fonts.Add(name);
                pos += entryLen + cbExtra;
            }
            return fonts.Count > 0 ? fonts : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Odwzorowuje offset bajtowy znaku (FC w strumieniu WordDocument) na jego bezpośrednie
    /// formatowanie. FKP-y parsowane leniwie i cache'owane; każdy błąd struktury → format domyślny.
    /// </summary>
    private sealed class ChpxIndex
    {
        private readonly byte[] _wd;
        private readonly long[] _boundaries;       // n+1 granic FC (rosnące)
        private readonly int[] _pages;             // n numerów stron FKP
        private readonly IReadOnlyList<string>? _fonts; // SttbfFfn: ftc → nazwa fontu
        private readonly Dictionary<int, Fkp> _fkpCache = new();

        public ChpxIndex(byte[] wd, long[] boundaries, int[] pages, IReadOnlyList<string>? fonts)
        {
            _wd = wd;
            _boundaries = boundaries;
            _pages = pages;
            _fonts = fonts;
        }

        public CharFormat GetFormat(long fc)
        {
            int bucket = FindBucket(_boundaries, fc, _pages.Length);
            if (bucket < 0) return default;
            var fkp = GetFkp(_pages[bucket]);
            return fkp.GetFormat(fc);
        }

        private Fkp GetFkp(int page)
        {
            if (_fkpCache.TryGetValue(page, out var cached)) return cached;
            var fkp = Fkp.Parse(_wd, page, _fonts);
            _fkpCache[page] = fkp;
            return fkp;
        }
    }

    /// <summary>Sparsowana strona CHPX FKP: granice FC run-ów + gotowe formaty.</summary>
    private sealed class Fkp
    {
        private static readonly Fkp Empty = new(Array.Empty<long>(), Array.Empty<CharFormat>());

        private readonly long[] _boundaries; // crun+1
        private readonly CharFormat[] _formats; // crun

        private Fkp(long[] boundaries, CharFormat[] formats)
        {
            _boundaries = boundaries;
            _formats = formats;
        }

        public CharFormat GetFormat(long fc)
        {
            int bucket = FindBucket(_boundaries, fc, _formats.Length);
            return bucket < 0 ? default : _formats[bucket];
        }

        public static Fkp Parse(byte[] wd, int page, IReadOnlyList<string>? fonts)
        {
            try
            {
                long baseOff = (long)page * FkpSize;
                if (baseOff < 0 || baseOff + FkpSize > wd.Length) return Empty;

                int crun = wd[baseOff + FkpSize - 1];
                if (crun <= 0) return Empty;

                int rgfcBytes = (crun + 1) * 4;
                int rgbBase = (int)baseOff + rgfcBytes;
                if (rgbBase + crun > baseOff + FkpSize - 1) return Empty;

                var boundaries = new long[crun + 1];
                for (int i = 0; i <= crun; i++)
                    boundaries[i] = BinaryPrimitives.ReadUInt32LittleEndian(wd.AsSpan((int)baseOff + i * 4));

                var formats = new CharFormat[crun];
                for (int i = 0; i < crun; i++)
                {
                    byte wordOffset = wd[rgbBase + i];
                    if (wordOffset == 0) { formats[i] = default; continue; } // brak wyjątku → domyślne
                    int chpxOff = (int)baseOff + wordOffset * 2;
                    if (chpxOff < baseOff || chpxOff >= baseOff + FkpSize) { formats[i] = default; continue; }
                    int cb = wd[chpxOff];
                    int grpprlStart = chpxOff + 1;
                    if (cb <= 0 || grpprlStart + cb > baseOff + FkpSize) { formats[i] = default; continue; }
                    formats[i] = ParseChpxGrpprl(wd.AsSpan(grpprlStart, cb), fonts);
                }

                return new Fkp(boundaries, formats);
            }
            catch
            {
                return Empty;
            }
        }
    }

    /// <summary>Binary-search kubełka: indeks i taki, że boundaries[i] &lt;= fc &lt; boundaries[i+1]; -1 poza zakresem.</summary>
    private static int FindBucket(long[] boundaries, long fc, int count)
    {
        if (count <= 0) return -1;
        if (fc < boundaries[0] || fc >= boundaries[count]) return -1;
        int lo = 0, hi = count - 1;
        while (lo <= hi)
        {
            int mid = (lo + hi) >> 1;
            if (fc < boundaries[mid]) hi = mid - 1;
            else if (fc >= boundaries[mid + 1]) lo = mid + 1;
            else return mid;
        }
        return -1;
    }

    // ---- SPRM-y (single properties) --------------------------------------------

    // Character SPRM-y istotne dla podglądu formatowania (sgc=2). Wartości opcode za [MS-DOC] 2.6.1.
    private const ushort SprmCFBold = 0x0835;    // ToggleOperand (1 B)
    private const ushort SprmCFItalic = 0x0836;  // ToggleOperand (1 B)
    private const ushort SprmCFStrike = 0x0837;  // ToggleOperand (1 B)
    private const ushort SprmCKul = 0x2A3E;      // Kul (1 B): 0=brak, ≠0=podkreślenie
    private const ushort SprmCIco = 0x2A42;      // Ico (1 B): indeks w 16-kolorowej palecie
    private const ushort SprmCCv = 0x6870;       // COLORREF (4 B): R,G,B,fAuto
    private const ushort SprmCHps = 0x4A43;      // rozmiar czcionki (2 B, half-points)
    private const ushort SprmCRgFtc0 = 0x4A4F;   // ftc ASCII (2 B): indeks w SttbfFfn

    /// <summary>Interpretuje grpprl (ciąg SPRM-ów) CHPX i składa z niego wspierane formatowanie znaku.</summary>
    private static CharFormat ParseChpxGrpprl(ReadOnlySpan<byte> grpprl, IReadOnlyList<string>? fonts)
    {
        bool bold = false, italic = false, underline = false, strike = false;
        string? color = null;
        ushort? sizeHalfPoints = null;
        string? fontName = null;

        int p = 0;
        while (p + 2 <= grpprl.Length)
        {
            ushort sprm = BinaryPrimitives.ReadUInt16LittleEndian(grpprl.Slice(p));
            p += 2;

            int spra = (sprm >> 13) & 0x7;
            int opLen;
            if (spra == 6)
            {
                if (p >= grpprl.Length) break;
                opLen = grpprl[p];
                p += 1;
            }
            else
            {
                opLen = spra switch { 0 => 1, 1 => 1, 2 => 2, 3 => 4, 4 => 2, 5 => 2, 7 => 3, _ => 0 };
            }
            if (opLen < 0 || p + opLen > grpprl.Length) break;
            var operand = grpprl.Slice(p, opLen);
            p += opLen;

            switch (sprm)
            {
                case SprmCFBold when operand.Length >= 1: bold = ToggleOn(operand[0]); break;
                case SprmCFItalic when operand.Length >= 1: italic = ToggleOn(operand[0]); break;
                case SprmCFStrike when operand.Length >= 1: strike = ToggleOn(operand[0]); break;
                case SprmCKul when operand.Length >= 1: underline = operand[0] != 0; break;
                case SprmCIco when operand.Length >= 1:
                    var mapped = IcoToHex(operand[0]);
                    if (mapped != null) color = mapped;
                    break;
                case SprmCCv when operand.Length >= 3:
                    color = $"{operand[0]:X2}{operand[1]:X2}{operand[2]:X2}";
                    break;
                case SprmCHps when operand.Length >= 2:
                    var hps = BinaryPrimitives.ReadUInt16LittleEndian(operand);
                    if (hps is > 0 and <= 3276) sizeHalfPoints = hps; // limit Worda: 1638 pt
                    break;
                case SprmCRgFtc0 when operand.Length >= 2:
                    int ftc = BinaryPrimitives.ReadUInt16LittleEndian(operand);
                    if (fonts != null && ftc < fonts.Count && !string.IsNullOrEmpty(fonts[ftc]))
                        fontName = fonts[ftc];
                    break;
            }
        }

        return new CharFormat(bold, italic, underline, strike, color, sizeHalfPoints, fontName);
    }

    /// <summary>
    /// ToggleOperand: 0=wył., 1=wł., 0x80=dziedzicz ze stylu (brak stylu → wył.), 0x81=odwróć styl
    /// (odwrócenie domyślnego „wył." → wł.). Budujemy płaski DOCX bez stylów, więc bazą jest „wył.".
    /// </summary>
    private static bool ToggleOn(byte value) => value == 0x01 || value == 0x81;

    /// <summary>Mapuje indeks Ico (paleta 16 kolorów MS-DOC) na hex; null dla auto (0) i poza zakresem.</summary>
    private static string? IcoToHex(byte ico) => ico switch
    {
        1 => "000000",  // black
        2 => "0000FF",  // blue
        3 => "00FFFF",  // cyan
        4 => "00FF00",  // green
        5 => "FF00FF",  // magenta
        6 => "FF0000",  // red
        7 => "FFFF00",  // yellow
        8 => "FFFFFF",  // white
        9 => "000080",  // dark blue
        10 => "008080", // dark cyan
        11 => "008000", // dark green
        12 => "800080", // dark magenta
        13 => "800000", // dark red
        14 => "808000", // dark yellow
        15 => "808080", // dark gray
        16 => "C0C0C0", // light gray
        _ => null,      // 0 = auto (dziedziczy) lub nieznany indeks
    };
}

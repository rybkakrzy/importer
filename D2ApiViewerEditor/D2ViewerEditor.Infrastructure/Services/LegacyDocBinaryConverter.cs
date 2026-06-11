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
/// dokumentu i buduje minimalny, poprawny DOCX przez OpenXML.
///
/// Zakres (świadomy): odzyskuje TEKST i podział akapitów. Bogate formatowanie (czcionki, style,
/// tabele, obrazy, nagłówki/stopki) nie jest odtwarzane — binarny .doc to format sprzed OOXML i
/// pełna wierność wymagałaby dużego parsera lub konwertera zewnętrznego. Gdy cokolwiek w strukturze
/// jest niespójne, metoda zwraca null, a wołający stosuje kontrolowane odrzucenie (bez śmieci).
/// </summary>
public static class LegacyDocBinaryConverter
{
    private const ushort WordDocumentMagic = 0xA5EC;     // wIdent w FIB
    private const int FibFlagsOffset = 0x000A;           // bit 0x0200 = fWhichTblStm (0Table/1Table)
    private const int FcClxOffset = 0x01A2;              // fcClx w FibRgFcLcb97
    private const int LcbClxOffset = 0x01A6;             // lcbClx
    private const ushort FWhichTblStmBit = 0x0200;
    private const uint FcCompressedFlag = 0x40000000;    // PCD.fc: 1 bajt/znak (CP1252) zamiast UTF-16
    private const uint FcValueMask = 0x3FFFFFFF;
    private const int MaxChars = 8_000_000;              // twardy limit (anty-abuse)

    /// <summary>
    /// Próbuje przekonwertować binarny .doc do DOCX. Zwraca bajty DOCX albo null, gdy nie da się
    /// bezpiecznie odczytać (uszkodzony FIB/CLX, brak tekstu, nieobsługiwany wariant).
    /// </summary>
    public static byte[]? TryConvert(byte[] wordDocument, byte[]? table0Stream, byte[]? table1Stream)
    {
        var text = ExtractText(wordDocument, table0Stream, table1Stream);
        if (string.IsNullOrWhiteSpace(text)) return null;
        return BuildDocx(text);
    }

    private static string? ExtractText(byte[] wd, byte[]? table0, byte[]? table1)
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

        return ReadPieces(wd, plcPcd);
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
    /// </summary>
    private static string? ReadPieces(byte[] wd, byte[] plcPcd)
    {
        // lcb = 4*(n+1) + 8*n  ⇒  n = (lcb - 4) / 12
        if (plcPcd.Length < 4 + 8) return null;
        int n = (plcPcd.Length - 4) / 12;
        if (n <= 0) return null;

        int cpBase = 0;
        int pcdBase = 4 * (n + 1);
        var sb = new StringBuilder();

        for (int i = 0; i < n; i++)
        {
            uint cpStart = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(cpBase + i * 4));
            uint cpEnd = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(cpBase + (i + 1) * 4));
            if (cpEnd <= cpStart) continue;
            long cch = cpEnd - cpStart;
            if (cch > MaxChars || sb.Length + cch > MaxChars) return null;

            uint fc = BinaryPrimitives.ReadUInt32LittleEndian(plcPcd.AsSpan(pcdBase + i * 8 + 2));
            bool compressed = (fc & FcCompressedFlag) != 0;
            uint fcValue = fc & FcValueMask;

            if (compressed)
            {
                long off = fcValue / 2;
                if (off < 0 || off + cch > wd.Length) return null;
                AppendChars(sb, Encoding.Latin1.GetString(wd, (int)off, (int)cch));
            }
            else
            {
                long off = fcValue;
                if (off < 0 || off + cch * 2 > wd.Length) return null;
                AppendChars(sb, Encoding.Unicode.GetString(wd, (int)off, (int)(cch * 2)));
            }
        }

        return sb.ToString();
    }

    /// <summary>Normalizuje znaki sterujące Worda: 0x0D = koniec akapitu (\n); pozostałe kontrolki pomijane.</summary>
    private static void AppendChars(StringBuilder sb, string chunk)
    {
        foreach (char c in chunk)
        {
            if (c == '\r' || c == '\n') sb.Append('\n');     // znak akapitu
            else if (c == '\t') sb.Append('\t');
            else if (c >= ' ') sb.Append(c);                  // drukowalne (w tym wynik pól)
            // pozostałe (cell mark 0x07, field 0x13/0x14/0x15, picture 0x01, NUL, …) — pomijamy
        }
    }

    private static byte[] BuildDocx(string text)
    {
        // Word używa CR jako znaku akapitu — dzielimy na akapity i budujemy minimalny DOCX.
        var paragraphs = text.Replace("\r\n", "\n").Split('\n');

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            var body = main.Document.AppendChild(new Body());

            foreach (var line in paragraphs)
            {
                var p = new Paragraph();
                if (line.Length > 0)
                {
                    var run = new Run();
                    run.AppendChild(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                    p.AppendChild(run);
                }
                body.AppendChild(p);
            }

            main.Document.Save();
        }
        return ms.ToArray();
    }
}

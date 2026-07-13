using System.Buffers.Binary;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Konwerter binarnego .doc (MS-DOC) → .docx: parsowanie FIB + piece table, ekstrakcja tekstu i
/// podziału akapitów, bezpieczne odrzucenie uszkodzonego wejścia. Czysto zarządzany (bez GDI/LO).
/// </summary>
[TestFixture]
public class LegacyDocBinaryConverterTests
{
    [Test]
    public void TryConvert_CompressedPiece_ExtractsTextAndSplitsParagraphs()
    {
        // Tekst "Hello\rWorld" (0x0D = znak akapitu) w piece "compressed" (1 bajt/znak, Latin1).
        const string raw = "Hello\rWorld";
        var (wd, table0) = BuildMinimalDoc(raw, textByteOffset: 512, compressed: true, useTable1: false);

        var docx = LegacyDocBinaryConverter.TryConvert(wd, table0, null);

        docx.Should().NotBeNull();
        var paragraphs = ReadParagraphs(docx!);
        paragraphs.Should().Contain("Hello");
        paragraphs.Should().Contain("World");
    }

    [Test]
    public void TryConvert_UnicodePiece_ExtractsText()
    {
        const string raw = "Zażółć"; // znaki spoza ASCII → wariant 16-bit Unicode
        var (wd, table0) = BuildMinimalDoc(raw, textByteOffset: 512, compressed: false, useTable1: false);

        var docx = LegacyDocBinaryConverter.TryConvert(wd, table0, null);

        docx.Should().NotBeNull();
        ReadParagraphs(docx!).Should().Contain("Zażółć");
    }

    [Test]
    public void TryConvert_WithPrcBeforePcdt_SkipsPrcAndExtractsText()
    {
        var (wd, table0) = BuildMinimalDoc("Hello", textByteOffset: 512, compressed: true, useTable1: false, withPrc: true);

        var docx = LegacyDocBinaryConverter.TryConvert(wd, table0, null);

        docx.Should().NotBeNull();
        ReadParagraphs(docx!).Should().Contain("Hello");
    }

    [Test]
    public void TryConvert_BadMagic_ReturnsNull()
    {
        var (wd, table0) = BuildMinimalDoc("Hello", 512, compressed: true, useTable1: false);
        BinaryPrimitives.WriteUInt16LittleEndian(wd, 0x1234); // zepsuj wIdent

        LegacyDocBinaryConverter.TryConvert(wd, table0, null).Should().BeNull();
    }

    [Test]
    public void TryConvert_NoTableStream_ReturnsNull()
    {
        var (wd, _) = BuildMinimalDoc("Hello", 512, compressed: true, useTable1: false);

        LegacyDocBinaryConverter.TryConvert(wd, null, null).Should().BeNull();
    }

    // ---- formatowanie znaków (CHPX) ---------------------------------------------

    /// <summary>
    /// Regresja zgłoszenia: .doc z pogrubieniem/kursywą/podkreśleniem/przekreśleniem/kolorami
    /// (osobno i łącznie) — formatowanie musi round-tripować z warstwy CHPX do run properties DOCX,
    /// nie spłaszczać się do zwykłego tekstu.
    /// </summary>
    [Test]
    public void TryConvert_WithCharacterFormatting_PreservesBoldItalicUnderlineStrikeAndColors()
    {
        var segments = new (string text, byte[]? grpprl)[]
        {
            ("Bold", Sprm(0x0835, 0x01)),                         // sprmCFBold
            ("Ital", Sprm(0x0836, 0x01)),                         // sprmCFItalic
            ("Undr", Sprm(0x2A3E, 0x01)),                         // sprmCKul (single)
            ("Strk", Sprm(0x0837, 0x01)),                         // sprmCFStrike
            ("Redd", Sprm(0x2A42, 0x06)),                         // sprmCIco = 6 (red)
            ("Cvbl", Sprm(0x6870, 0x00, 0x00, 0xFF, 0x00)),       // sprmCCv COLORREF = blue
            ("Comb", Combine(Sprm(0x0835, 0x01), Sprm(0x0836, 0x01),
                             Sprm(0x0837, 0x01), Sprm(0x2A42, 0x06))), // bold+italic+strike+red
            ("Plain", null),                                      // brak formatowania
        };

        var (wd, table) = BuildFormattedDoc(segments);

        var docx = LegacyDocBinaryConverter.TryConvert(wd, table, null);
        docx.Should().NotBeNull();

        var runs = ReadRuns(docx!);

        RunProps(runs, "Bold").Bold.Should().NotBeNull();
        RunProps(runs, "Ital").Italic.Should().NotBeNull();

        var undr = RunProps(runs, "Undr").Underline;
        undr.Should().NotBeNull();
        undr!.Val!.Value.Should().Be(UnderlineValues.Single);

        RunProps(runs, "Strk").Strike.Should().NotBeNull();

        RunProps(runs, "Redd").GetFirstChild<Color>()!.Val!.Value.Should().Be("FF0000");
        RunProps(runs, "Cvbl").GetFirstChild<Color>()!.Val!.Value.Should().Be("0000FF");

        var comb = RunProps(runs, "Comb");
        comb.Bold.Should().NotBeNull();
        comb.Italic.Should().NotBeNull();
        comb.Strike.Should().NotBeNull();
        comb.GetFirstChild<Color>()!.Val!.Value.Should().Be("FF0000");

        // Zwykły tekst nie dostaje żadnych właściwości runa (nie „udajemy” formatowania).
        runs.Single(r => r.InnerText == "Plain").RunProperties.Should().BeNull();
    }

    /// <summary>Uszkodzona warstwa CHPX nie może wywalać importu — degradacja do samego tekstu.</summary>
    [Test]
    public void TryConvert_CorruptChpxLayer_StillExtractsTextWithoutFormatting()
    {
        var (wd, table) = BuildFormattedDoc(new (string, byte[]?)[] { ("Hello", Sprm(0x0835, 0x01)) });
        // Wskaż PnFkpChpx na stronę FKP daleko poza strumieniem — parser FKP musi to złapać.
        BinaryPrimitives.WriteInt32LittleEndian(wd.AsSpan(0x00FA), table.Length); // fcPlcfBteChpx w tablicy
        var badTable = table.ToArray();
        // Ostatnie 4 bajty tablicy to numer strony FKP — ustaw absurdalnie duży.
        BinaryPrimitives.WriteUInt32LittleEndian(badTable.AsSpan(badTable.Length - 4), 0x000FFFFF);

        var docx = LegacyDocBinaryConverter.TryConvert(wd, badTable, null);

        docx.Should().NotBeNull();
        var runs = ReadRuns(docx!);
        runs.Should().ContainSingle(r => r.InnerText == "Hello");
        runs.Single(r => r.InnerText == "Hello").RunProperties.Should().BeNull();
    }

    // ---- synthetic MS-DOC builder ------------------------------------------------

    /// <summary>Składa SPRM: 2-bajtowy opcode (little-endian) + operand.</summary>
    private static byte[] Sprm(ushort opcode, params byte[] operand)
    {
        var b = new byte[2 + operand.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(b, opcode);
        operand.CopyTo(b, 2);
        return b;
    }

    private static byte[] Combine(params byte[][] grpprls) => grpprls.SelectMany(g => g).ToArray();

    /// <summary>
    /// Buduje .doc (compressed, tekst @512) z warstwą CHPX: jedna strona FKP (@ fkpPage*512) z run-ami
    /// wg segmentów i PlcfBteChpx w strumieniu tablicy pokrywającym całą treść. Pozwala testować
    /// odzyskiwanie bezpośredniego formatowania znaków dokładnie tą samą ścieżką co produkcja.
    /// </summary>
    private static (byte[] wd, byte[] table) BuildFormattedDoc(
        IReadOnlyList<(string text, byte[]? grpprl)> segments, int fkpPage = 2)
    {
        const int textByteOffset = 512;
        string text = string.Concat(segments.Select(s => s.text));
        var textBytes = Encoding.Latin1.GetBytes(text);
        int cch = text.Length;

        int wdLen = Math.Max(0x01AA + 4, Math.Max(textByteOffset + textBytes.Length + 16, fkpPage * 512 + 512));
        var wd = new byte[wdLen];
        BinaryPrimitives.WriteUInt16LittleEndian(wd, 0xA5EC);                 // wIdent
        BinaryPrimitives.WriteUInt16LittleEndian(wd.AsSpan(0x000A), 0);       // fWhichTblStm = 0Table
        textBytes.CopyTo(wd, textByteOffset);

        // PlcPcd: jeden compressed piece na całą treść.
        var plcPcd = new byte[4 * 2 + 8];
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(0), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(4), (uint)cch);
        uint fc = 0x40000000u | (uint)(textByteOffset * 2);
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(8 + 2), fc);

        var pcdt = new byte[1 + 4 + plcPcd.Length];
        pcdt[0] = 0x02;
        BinaryPrimitives.WriteUInt32LittleEndian(pcdt.AsSpan(1), (uint)plcPcd.Length);
        plcPcd.CopyTo(pcdt, 5);
        byte[] clx = pcdt;

        // Granice FC run-ów (offsety bajtowe w WordDocument): compressed ⇒ znak k @ 512 + k.
        int crun = segments.Count;
        var boundaries = new int[crun + 1];
        int cursor = textByteOffset;
        boundaries[0] = cursor;
        for (int i = 0; i < crun; i++) { cursor += segments[i].text.Length; boundaries[i + 1] = cursor; }

        // CHPX FKP (512 B): rgfc, rgb (offset słowa CHPX), dane CHPX (cb + grpprl), crun @511.
        var fkp = new byte[512];
        for (int i = 0; i <= crun; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(fkp.AsSpan(i * 4), (uint)boundaries[i]);
        int rgbBase = (crun + 1) * 4;
        int dataOff = rgbBase + crun;
        for (int i = 0; i < crun; i++)
        {
            var g = segments[i].grpprl;
            if (g == null || g.Length == 0) { fkp[rgbBase + i] = 0; continue; }
            if (dataOff % 2 != 0) dataOff++;            // rgb wskazuje słowa (offset parzysty)
            fkp[rgbBase + i] = (byte)(dataOff / 2);
            fkp[dataOff] = (byte)g.Length;               // cb
            g.CopyTo(fkp, dataOff + 1);
            dataOff += 1 + g.Length;
        }
        fkp[511] = (byte)crun;
        fkp.CopyTo(wd, fkpPage * 512);

        // PlcfBteChpx (w tablicy): 2 FC (początek/koniec) + 1 PnFkpChpx (numer strony FKP).
        var plcfBte = new byte[8 + 4];
        BinaryPrimitives.WriteUInt32LittleEndian(plcfBte.AsSpan(0), (uint)textByteOffset);
        BinaryPrimitives.WriteUInt32LittleEndian(plcfBte.AsSpan(4), (uint)cursor);
        BinaryPrimitives.WriteUInt32LittleEndian(plcfBte.AsSpan(8), (uint)fkpPage);
        var table = clx.Concat(plcfBte).ToArray();

        // Wskaźniki FIB: CLX na początku tablicy, PlcfBteChpx tuż za nim.
        BinaryPrimitives.WriteInt32LittleEndian(wd.AsSpan(0x01A2), 0);                 // fcClx
        BinaryPrimitives.WriteUInt32LittleEndian(wd.AsSpan(0x01A6), (uint)clx.Length); // lcbClx
        BinaryPrimitives.WriteInt32LittleEndian(wd.AsSpan(0x00FA), clx.Length);        // fcPlcfBteChpx
        BinaryPrimitives.WriteUInt32LittleEndian(wd.AsSpan(0x00FE), (uint)plcfBte.Length); // lcbPlcfBteChpx

        return (wd, table);
    }

    private static List<Run> ReadRuns(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Body!.Descendants<Run>().ToList();
    }

    private static RunProperties RunProps(List<Run> runs, string text) =>
        runs.Single(r => r.InnerText == text).RunProperties!;

    /// <summary>
    /// Buduje minimalny strumień WordDocument (FIB) + tablicę (CLX/PlcPcd) z jednym piece.
    /// </summary>
    private static (byte[] wd, byte[] table) BuildMinimalDoc(
        string text, int textByteOffset, bool compressed, bool useTable1, bool withPrc = false)
    {
        var textBytes = compressed ? Encoding.Latin1.GetBytes(text) : Encoding.Unicode.GetBytes(text);
        int cch = text.Length;

        var wd = new byte[Math.Max(0x01AA + 4, textByteOffset + textBytes.Length + 16)];
        BinaryPrimitives.WriteUInt16LittleEndian(wd, 0xA5EC);                                  // wIdent
        BinaryPrimitives.WriteUInt16LittleEndian(wd.AsSpan(0x000A), (ushort)(useTable1 ? 0x0200 : 0)); // fWhichTblStm
        textBytes.CopyTo(wd, textByteOffset);

        // PlcPcd: CP[0..1] (2× UInt32) + 1× PCD (8 B)
        var plcPcd = new byte[4 * 2 + 8];
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(0), 0);          // CP start
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(4), (uint)cch);  // CP end
        uint fc = compressed
            ? (0x40000000u | (uint)(textByteOffset * 2))   // compressed: fc = offset*2 + flag
            : (uint)textByteOffset;                         // unicode: fc = offset
        BinaryPrimitives.WriteUInt32LittleEndian(plcPcd.AsSpan(8 + 2), fc);     // PCD.fc (po 2 B flag)

        // Pcdt: clxt=2 + lcb (UInt32) + PlcPcd
        var pcdt = new byte[1 + 4 + plcPcd.Length];
        pcdt[0] = 0x02;
        BinaryPrimitives.WriteUInt32LittleEndian(pcdt.AsSpan(1), (uint)plcPcd.Length);
        plcPcd.CopyTo(pcdt, 5);

        // Opcjonalny Prc (clxt=1, cbGrpprl=4 + 4 bajty) przed Pcdt — parser musi go pominąć.
        byte[] clx = withPrc
            ? new byte[] { 0x01, 0x04, 0x00, 0xAA, 0xBB, 0xCC, 0xDD }.Concat(pcdt).ToArray()
            : pcdt;

        // fcClx = 0, lcbClx = clx.Length (CLX na początku strumienia tablicy)
        BinaryPrimitives.WriteInt32LittleEndian(wd.AsSpan(0x01A2), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(wd.AsSpan(0x01A6), (uint)clx.Length);

        return (wd, clx);
    }

    private static List<string> ReadParagraphs(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();
    }
}

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

    // ---- synthetic MS-DOC builder ------------------------------------------------

    /// <summary>
    /// Buduje minimalny strumień WordDocument (FIB) + tablicę (CLX/PlcPcd) z jednym piece.
    /// </summary>
    private static (byte[] wd, byte[] table) BuildMinimalDoc(string text, int textByteOffset, bool compressed, bool useTable1)
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

        // CLX: Pcdt (clxt=2) + lcb (UInt32) + PlcPcd
        var clx = new byte[1 + 4 + plcPcd.Length];
        clx[0] = 0x02;
        BinaryPrimitives.WriteUInt32LittleEndian(clx.AsSpan(1), (uint)plcPcd.Length);
        plcPcd.CopyTo(clx, 5);

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

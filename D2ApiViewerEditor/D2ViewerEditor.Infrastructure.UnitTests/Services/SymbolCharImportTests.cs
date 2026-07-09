using System.IO;
using System.Linq;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Regresja zgłoszenia „znaki specjalne (strzałki i inne) nie wyświetlają się w edytorze i nie
/// przenoszą się do finalnego dokumentu". Przyczyna: znaki wstawione przez `w:sym`
/// (Insert → Symbol z fontu Symbol/Wingdings) reader emitował jako surowy punkt Private Use Area
/// (`&amp;#xF0E0;`), który bez fontu symbolicznego jest niewidoczny, a po round-tripie wracał jako
/// niewidoczny znak. Teraz: znane glify → prawdziwy Unicode, reszta → znak w oryginalnym foncie.
/// </summary>
[TestFixture]
public class SymbolCharImportTests
{
    // ── Import (w:sym → HTML) ─────────────────────────────────────────────

    [Test]
    public void SymbolFontRightArrow_MapsToUnicodeArrow()
    {
        // Word: strzałka w prawo z fontu Symbol = w:sym w:font="Symbol" w:char="F0AE" (PUA-shifted 0xAE).
        var html = ConvertSym("Symbol", "F0AE");

        html.Should().Contain("→");           // →
        html.Should().NotContain("F0AE");          // nie surowy PUA
        html.Should().NotContain("&#x");           // nie encja PUA
    }

    [Test]
    public void SymbolFontArrows_MapToUnicode_BothPuaAndRawByte()
    {
        // Ta sama strzałka bywa zapisana jako 0x00AE (raw) albo 0xF0AE (PUA) — obie muszą dać →.
        ConvertSym("Symbol", "00AE").Should().Contain("→");
        ConvertSym("Symbol", "F0AE").Should().Contain("→");
    }

    [Test]
    public void WingdingsCheckbox_MapsToUnicodeCheckedBox()
    {
        // Wingdings 0xFE = zaznaczony checkbox ☑.
        ConvertSym("Wingdings", "F0FE").Should().Contain("☑"); // ☑
    }

    [Test]
    public void UnmappedWingdingsGlyph_KeepsCharInOriginalFont_NotInvisiblePua()
    {
        // Wingdings 0xE0 nie ma pewnego mapowania Unicode → zachowujemy znak w foncie Wingdings
        // (niski bajt 0xE0 = 'à'), żeby renderował się z Wingdings i przetrwał round-trip.
        var html = ConvertSym("Wingdings", "F0E0");

        html.Should().Contain("font-family:'Wingdings'");
        // WebUtility.HtmlEncode koduje 'à' (U+00E0) jako &#224; — dekoduje się i renderuje w Wingdings.
        System.Net.WebUtility.HtmlDecode(html).Should().Contain("à"); // à — niski bajt, glif Wingdings
        html.Should().NotContain("F0E0");  // nie surowy PUA
        html.Should().NotContain("&#x");   // nie surowa encja PUA (heks)
    }

    // ── Round-trip (w:sym → HTML → DOCX) ──────────────────────────────────

    [Test]
    public void MappedArrow_SurvivesRoundTripAsText()
    {
        var html = ConvertSym("Symbol", "F0AE");
        var docx = new HtmlToDocxConverter().Convert(html);

        BodyText(docx).Should().Contain("→"); // → wraca jako zwykły tekst w:t
    }

    [Test]
    public void UnmappedGlyph_SurvivesRoundTripWithFont()
    {
        var html = ConvertSym("Wingdings", "F0E0");
        var docx = new HtmlToDocxConverter().Convert(html);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var run = doc.MainDocumentPart!.Document.Body!
            .Descendants<Run>()
            .FirstOrDefault(r => r.InnerText.Contains('à'));

        run.Should().NotBeNull("znak z Wingdings nie może zniknąć przy zapisie");
        run!.RunProperties?.RunFonts?.Ascii?.Value.Should().Be("Wingdings");
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static string ConvertSym(string font, string hexChar)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(
                        new RunProperties(),
                        new SymbolChar { Font = font, Char = hexChar }))));
            main.Document.Save();
        }
        return new DocxToHtmlConverter().Convert(new MemoryStream(ms.ToArray())).Html;
    }

    private static string BodyText(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Body!.InnerText;
    }
}

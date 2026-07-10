using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Znaki specjalne z fontów symbolicznych (strzałki, checkboxy itp.) muszą renderować się
/// w edytorze jak w MS Word. Zgłoszenie: „Utwórz KYC → Know Your Customer" — w Wordzie
/// strzałka, w DOC2 kwadrat (tofu). Root cause: (1) <c>w:sym</c> emitowany jako goła encja
/// Private Use Area (U+F0xx) bez mapowania na Unicode i bez font-family; (2) znaki PUA /
/// znaki bajtowe fontów symbolicznych w zwykłym <c>w:t</c> (autokorekta Worda „--&gt;"
/// wstawia <c>è</c> w foncie Wingdings) przechodziły literalnie.
/// </summary>
[TestFixture]
public class SymbolCharFidelityTests
{
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup() => _reader = new DocxToHtmlConverter();

    // --- w:sym ------------------------------------------------------------------

    [Test]
    public void Sym_SymbolFontArrow_MapsToUnicodeArrow()
    {
        // Symbol 0xAE = → (strzałka w prawo); Word zapisuje kod z przesunięciem PUA F0AE.
        var docx = BuildBody(ParagraphWithSym("Symbol", "F0AE"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("→");
        html.Should().NotContain("&#xF0AE;");
    }

    [Test]
    public void Sym_WingdingsArrow_MapsToUnicodeArrow()
    {
        // Wingdings 0xE8 = ➔ — strzałka, którą autokorekta Worda wstawia za „-->".
        var docx = BuildBody(ParagraphWithSym("Wingdings", "F0E8"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("➔");
    }

    [Test]
    public void Sym_CodeWithoutPuaShift_IsAlsoMapped()
    {
        // Niektóre dokumenty zapisują kod bajtowy bez przesunięcia PUA (w:char="FC").
        var docx = BuildBody(ParagraphWithSym("Wingdings", "FC"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("✔");
    }

    [Test]
    public void Sym_NormalFontWithPlainUnicode_EmitsCharacter()
    {
        // w:sym bywa używany też ze zwykłym fontem i normalnym code-pointem.
        var docx = BuildBody(ParagraphWithSym("Arial", "2192"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("→");
    }

    [Test]
    public void Sym_UnknownSymbolCode_FallsBackToEntityWithSymbolFont()
    {
        // Kod spoza tabeli mapowań NIE może stać się gołą encją PUA bez fontu (tofu).
        // Najlepsze przybliżenie: encja + font-family fontu symbolicznego (jest na Windows).
        var docx = BuildBody(ParagraphWithSym("Wingdings 3", "F067"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("font-family:'Wingdings 3'");
        html.Should().Contain("&#xF067;");
    }

    // --- w:t (zwykły tekst) -------------------------------------------------------

    [Test]
    public void Text_PuaCharInWingdingsRun_MapsToUnicode()
    {
        var run = new Run(
            new RunProperties(new RunFonts { Ascii = "Wingdings" }),
            new Text("\uF0FC") { Space = SpaceProcessingModeValues.Preserve });
        var docx = BuildBody(new Paragraph(run));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("✔");
        html.Should().NotContain("\uF0FC");
    }

    [Test]
    public void Text_ByteCharInWingdingsRun_MapsToUnicode()
    {
        // Autokorekta Worda „-->" = literalne „è" (0xE8) w runie z fontem Wingdings.
        var run = new Run(
            new RunProperties(new RunFonts { Ascii = "Wingdings" }),
            new Text("è") { Space = SpaceProcessingModeValues.Preserve });
        var docx = BuildBody(new Paragraph(run));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("➔");
    }

    [Test]
    public void Text_PuaCharWithoutSymbolicFont_IsLeftUntouched()
    {
        // PUA bez fontu symbolicznego nie ma publicznej semantyki — nie zgadujemy glifu.
        var run = new Run(new Text("\uF0E0") { Space = SpaceProcessingModeValues.Preserve });
        var docx = BuildBody(new Paragraph(run));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("\uF0E0");
    }

    [Test]
    public void Text_SegoeUiSymbol_IsNotTreatedAsSymbolicFont()
    {
        // „Segoe UI Symbol" to normalny font Unicode — znaków NIE wolno przemapowywać.
        var run = new Run(
            new RunProperties(new RunFonts { Ascii = "Segoe UI Symbol" }),
            new Text("è") { Space = SpaceProcessingModeValues.Preserve });
        var docx = BuildBody(new Paragraph(run));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        // EscapeHtml koduje Latin-1 numerycznie: è = &#232; — ważne, że NIE ma zamiany na strzałkę.
        html.Should().Contain("&#232;");
        html.Should().NotContain("➔");
    }

    [Test]
    public void Text_PlainUnicodeArrow_PassesThroughUnchanged()
    {
        var run = new Run(new Text("Utwórz KYC → Know Your Customer") { Space = SpaceProcessingModeValues.Preserve });
        var docx = BuildBody(new Paragraph(run));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        // „ó" wychodzi jako encja (&#243;) — pilnujemy strzałki i reszty frazy.
        html.Should().Contain("→ Know Your Customer");
    }

    // --- builders -------------------------------------------------------------

    private static Paragraph ParagraphWithSym(string font, string charHex)
    {
        var sym = new SymbolChar { Font = font, Char = charHex };
        return new Paragraph(
            new Run(new Text("przed ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(sym),
            new Run(new Text(" po") { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static byte[] BuildBody(params Paragraph[] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var p in paragraphs) body.Append(p);
            body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }));
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }
}

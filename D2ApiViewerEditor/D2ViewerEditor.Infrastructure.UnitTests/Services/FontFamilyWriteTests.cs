using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// A font-family chosen in the toolbar must reach DOCX as a clean w:rFonts. The editor sets
/// element.style.fontFamily, which the browser serialises with DOUBLE quotes for multi-word
/// names ("Times New Roman"); the reader emits SINGLE quotes plus a generic fallback
/// ('Times New Roman',serif). The writer must strip quotes and the fallback so Word gets the
/// real face — otherwise the quote chars leak in and Word silently reverts to its default font
/// (the reported "zmiana czcionki nie działa / zapisuje się domyślną czcionką").
/// </summary>
[TestFixture]
public class FontFamilyWriteTests
{
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup() => _writer = new HtmlToDocxConverter();

    private static string FirstRunAscii(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var run = body.Descendants<Run>()
            .FirstOrDefault(r => r.InnerText.Contains("Treść")
                && r.GetFirstChild<RunProperties>()?.RunFonts?.Ascii != null);
        if (run == null)
            throw new AssertionException("No run with text 'Treść' and RunFonts/Ascii. Body XML:\n" + body.OuterXml);
        return run.RunProperties!.RunFonts!.Ascii!.Value!;
    }

    [Test]
    public void DoubleQuotedMultiWordFont_IsWrittenWithoutQuotes()
    {
        // Browser-serialised form after the user picks "Times New Roman" in the toolbar.
        var html = "<p><span style=\"font-family: &quot;Times New Roman&quot;\">Treść</span></p>";

        var docx = _writer.Convert(html);

        FirstRunAscii(docx).Should().Be("Times New Roman");
    }

    [Test]
    public void SingleQuotedFontWithGenericFallback_KeepsOnlyTheFace()
    {
        // Reader-emitted form (FontFamilyCss): 'Face',serif.
        var html = "<p><span style=\"font-family:'Calibri',sans-serif\">Treść</span></p>";

        var docx = _writer.Convert(html);

        FirstRunAscii(docx).Should().Be("Calibri");
    }

    [Test]
    public void UnquotedSingleWordFont_StillWorks()
    {
        var html = "<p><span style=\"font-family:Arial\">Treść</span></p>";

        var docx = _writer.Convert(html);

        FirstRunAscii(docx).Should().Be("Arial");
    }

    private static (string Font, string Text)[] RunFonts(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Body!
            .Descendants<Run>()
            .Where(r => r.Elements<Text>().Any())
            .Select(r => (r.RunProperties?.RunFonts?.Ascii?.Value ?? "(inherit)", r.InnerText))
            .ToArray();
    }

    [Test]
    public void NestedSpanFont_OverridesInheritedAncestorFont()
    {
        // Exact DOM the editor produces when a single-run paragraph (reader wraps run text in a
        // font span) has per-word fonts applied: extractContents+insertNode NESTS the new font
        // spans inside the original run's font span. The child font must win (CSS cascade),
        // otherwise the whole sentence collapses to the ancestor font (reported bug: "one word
        // Arial, next few Tahoma, rest nothing → whole sentence ends up Arial").
        var html =
            "<p><span style=\"font-family:'Calibri',sans-serif\">" +
            "<span style=\"font-family:Arial\">word1</span> " +
            "<span style=\"font-family:Tahoma\">word2 word3 word4</span> word5 word6" +
            "</span></p>";

        var runs = RunFonts(_writer.Convert(html));

        runs.Should().ContainSingle(r => r.Text == "word1").Which.Font.Should().Be("Arial");
        runs.Should().ContainSingle(r => r.Text == "word2 word3 word4").Which.Font.Should().Be("Tahoma");
        // Words with no explicit font inherit the ancestor span font (as in the editor).
        runs.Where(r => r.Text.Contains("word5")).Should().OnlyContain(r => r.Font == "Calibri");
    }

    [Test]
    public void NestedSpanFontSize_OverridesInheritedAncestorSize()
    {
        var html =
            "<p><span style=\"font-size:11pt\">" +
            "<span style=\"font-size:20pt\">big</span> small" +
            "</span></p>";

        using var ms = new MemoryStream(_writer.Convert(html));
        using var doc = WordprocessingDocument.Open(ms, false);
        var runs = doc.MainDocumentPart!.Document.Body!.Descendants<Run>()
            .Where(r => r.Elements<Text>().Any()).ToList();

        runs.Single(r => r.InnerText == "big").RunProperties!.FontSize!.Val!.Value.Should().Be("40");
        runs.Single(r => r.InnerText.Contains("small")).RunProperties!.FontSize!.Val!.Value.Should().Be("22");
    }

    [Test]
    public void NestedSpanColor_OverridesInheritedAncestorColor()
    {
        var html =
            "<p><span style=\"color:#000000\">" +
            "<span style=\"color:#FF0000\">red</span> black" +
            "</span></p>";

        using var ms = new MemoryStream(_writer.Convert(html));
        using var doc = WordprocessingDocument.Open(ms, false);
        var runs = doc.MainDocumentPart!.Document.Body!.Descendants<Run>()
            .Where(r => r.Elements<Text>().Any()).ToList();

        runs.Single(r => r.InnerText == "red").RunProperties!.Color!.Val!.Value.Should().Be("FF0000");
        runs.Single(r => r.InnerText.Contains("black")).RunProperties!.Color!.Val!.Value.Should().Be("000000");
    }
}

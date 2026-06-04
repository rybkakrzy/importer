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
}

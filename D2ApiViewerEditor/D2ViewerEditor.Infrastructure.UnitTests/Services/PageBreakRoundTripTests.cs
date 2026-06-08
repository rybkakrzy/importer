using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// R-15: a manual page break (w:br type=page) must survive DOCX → HTML → DOCX. The reader
/// emits it as &lt;div class="page-break"&gt; nested inside a paragraph; the writer must turn
/// any page-break representation back into a single w:br type=page — without inventing breaks
/// for documents that have none.
/// </summary>
[TestFixture]
public class PageBreakRoundTripTests
{
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup() => _writer = new HtmlToDocxConverter();

    private static int CountPageBreaks(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Body!
            .Descendants<Break>().Count(b => b.Type?.Value == BreakValues.Page);
    }

    [Test]
    public void ParagraphWithPageBreakBefore_EmitsPageBreakMarker_OnImport()
    {
        // Word „podział strony przed" (w:pageBreakBefore) — wcześniej ignorowane na imporcie,
        // przez co wielostronicowe dokumenty zwijały się do jednej strony (Issue „3 strony → 1").
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Strona 1"))),
                new Paragraph(
                    new ParagraphProperties(new PageBreakBefore()),
                    new Run(new Text("Strona 2"))),
                new Paragraph(
                    new ParagraphProperties(new PageBreakBefore()),
                    new Run(new Text("Strona 3")))));
            main.Document.Save();
        }

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(ms.ToArray())).Html;

        // Dwa markery → trzy strony po paginacji we froncie.
        System.Text.RegularExpressions.Regex.Matches(html, "class=\"page-break\"").Count.Should().Be(2);
    }

    [Test]
    public void ParagraphWithPageBreakBeforeFalse_DoesNotEmitMarker()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(new PageBreakBefore { Val = false }),
                    new Run(new Text("Bez podziału")))));
            main.Document.Save();
        }

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(ms.ToArray())).Html;

        html.Should().NotContain("page-break");
    }

    [Test]
    public void NestedPageBreakDiv_BecomesWBrTypePage()
    {
        // Exactly how the reader renders a Word page break (inside a paragraph's span).
        var html = "<p>Przed</p><p><span><div class=\"page-break\"></div></span></p><p>Po</p>";

        CountPageBreaks(_writer.Convert(html)).Should().Be(1);
    }

    [Test]
    public void TopLevelPageBreakDiv_BecomesWBrTypePage()
    {
        var html = "<p>A</p><div class=\"page-break\"></div><p>B</p>";

        CountPageBreaks(_writer.Convert(html)).Should().Be(1);
    }

    [Test]
    public void CssBreakBefore_IsRecognised()
    {
        var html = "<p>A</p><div style=\"break-before: page;\"></div><p>B</p>";

        CountPageBreaks(_writer.Convert(html)).Should().Be(1);
    }

    [Test]
    public void NoPageBreak_ProducesNone()
    {
        var html = "<p>Akapit jeden</p><p>Akapit dwa</p>";

        CountPageBreaks(_writer.Convert(html)).Should().Be(0);
    }

    [Test]
    public void ReaderEmitsTopLevelBlock_ForStandalonePageBreakParagraph()
    {
        // A dedicated page-break paragraph (only w:br type=page) must become a TOP-LEVEL
        // <div class="page-break"> — not nested inside <p><span> (which broke editor pagination).
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Przed"))),
                new Paragraph(new Run(new Break { Type = BreakValues.Page })),
                new Paragraph(new Run(new Text("Po")))));
            main.Document.Save();
        }
        ms.Position = 0;

        var content = new DocxToHtmlConverter().Convert(ms);

        content.Html.Should().Contain("<div class=\"page-break\"></div>");
        content.Html.Should().NotContain("<p style=\"\"><span style=\"\"><div class=\"page-break\">");
        // ...and it still round-trips back to exactly one w:br type=page.
        CountPageBreaks(_writer.Convert(content.Html)).Should().Be(1);
    }

    [Test]
    public void NotDuplicated_OnRepeatedConversion()
    {
        var html = "<p>A</p><p><span><div class=\"page-break\"></div></span></p><p>B</p>";

        // First save → DOCX. Re-importing would yield one page-break div again; saving again must
        // keep exactly one break (no accumulation).
        var first = _writer.Convert(html);
        CountPageBreaks(first).Should().Be(1);

        // Simulate a second round on equivalent HTML — still exactly one.
        CountPageBreaks(_writer.Convert(html)).Should().Be(1);
    }
}

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

    private static MemoryStream DocxWithPageBreakBeforeParagraphs()
    {
        var ms = new MemoryStream();
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
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void ParagraphWithPageBreakBefore_EmitsCssProperty_NotMarker()
    {
        // Word „podział strony przed" (w:pageBreakBefore) to WŁAŚCIWOŚĆ akapitu — reader
        // emituje `page-break-before:always` w stylu inline akapitu (paginacja GUI łamie
        // stronę przed blokiem), a NIE marker div.page-break (ten reprezentuje ręczny w:br).
        using var ms = DocxWithPageBreakBeforeParagraphs();

        var html = new DocxToHtmlConverter().Convert(ms).Html;

        html.Should().NotContain("class=\"page-break\"");
        System.Text.RegularExpressions.Regex
            .Matches(html, "page-break-before:always").Count.Should().Be(2);
    }

    [Test]
    public void PageBreakBeforeProperty_RoundTripsAsProperty_NotManualBreak()
    {
        // Pełny cykl: właściwość wraca jako w:pageBreakBefore w pPr TEGO SAMEGO akapitu —
        // bez zamiany na twardy w:br i bez dodatkowego pustego akapitu (checkbox w Wordzie
        // pozostaje zaznaczony po zapisie z edytora).
        using var ms = DocxWithPageBreakBeforeParagraphs();
        var html = new DocxToHtmlConverter().Convert(ms).Html;

        var bytes = _writer.Convert(html);

        using var result = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var body = result.MainDocumentPart!.Document.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();
        paragraphs.Should().HaveCount(3, "właściwość nie może dokładać akapitów z w:br");
        CountPageBreaks(bytes).Should().Be(0, "właściwość ≠ ręczny podział");
        paragraphs
            .Count(p => p.ParagraphProperties?.GetFirstChild<PageBreakBefore>() is { } pbb
                        && (pbb.Val == null || pbb.Val.Value))
            .Should().Be(2);
    }

    [Test]
    public void ParagraphWithPageBreakBeforeFalse_EmitsExplicitAuto_AndRoundTripsAsFalse()
    {
        // Jawne w:pageBreakBefore val=false (wyłączenie podziału ze STYLU) → CSS `auto`,
        // a writer odtwarza val=false — bez tego styl Worda przywróciłby podział.
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

        html.Should().NotContain("class=\"page-break\"");
        html.Should().Contain("page-break-before:auto");

        using var result = WordprocessingDocument.Open(new MemoryStream(_writer.Convert(html)), false);
        var pbb = result.MainDocumentPart!.Document.Body!
            .Descendants<PageBreakBefore>().Single();
        pbb.Val!.Value.Should().BeFalse();
    }

    [Test]
    public void PageBreakBeforeFromParagraphStyle_EmitsCssProperty()
    {
        // Styl akapitowy z pageBreakBefore (typowe „Nagłówek 1" rozdziałów) — właściwość
        // z definicji stylu trafia do CSS akapitu przez łańcuch stylów.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(
                    new StyleParagraphProperties(new PageBreakBefore()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Rozdzial",
                    StyleName = new StyleName { Val = "Rozdzial" }
                });
            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Rozdzial" }),
                    new Run(new Text("Nowy rozdział")))));
            main.Document.Save();
        }

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(ms.ToArray())).Html;

        html.Should().Contain("page-break-before:always");
        html.Should().NotContain("class=\"page-break\"");
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

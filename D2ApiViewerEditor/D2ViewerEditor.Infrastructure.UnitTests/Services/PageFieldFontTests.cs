using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// A page-number field (PAGE) in the footer must keep the font-size of its own run, not fall
/// back to the editor/container default. Reader emits the run's clean CSS on the field span;
/// writer round-trips the span's font into the PAGE field run's properties.
/// </summary>
[TestFixture]
public class PageFieldFontTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    [Test]
    public void ReaderEmitsPageFieldSpan_WithRunFontSize()
    {
        // Footer: PAGE field whose run is 8pt (sz=16 half-points).
        var docx = BuildFooterWithPageField(halfPoints: "16");

        var content = _reader.Convert(new MemoryStream(docx));

        content.Footer!.Html.Should().Contain("class=\"field-page\"");
        content.Footer.Html.Should().Contain("font-size:8pt");
    }

    [Test]
    public void WriterKeepsPageFieldRunFontSize_OnRoundTrip()
    {
        var footer = new HeaderFooterContent
        {
            Html = "<span class=\"field-page\" style=\"font-size:8pt\">1</span>"
        };

        var bytes = _writer.Convert("<p>Body</p>", footer: footer);

        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var footerPart = doc.MainDocumentPart!.FooterParts.First();
        var field = footerPart.Footer!.Descendants<SimpleField>()
            .First(f => f.Instruction?.Value?.Contains("PAGE") == true);
        var run = field.Descendants<Run>().First();
        run.RunProperties!.GetFirstChild<FontSize>()!.Val!.Value.Should().Be("16");
    }

    // --- builder --------------------------------------------------------------

    private static byte[] BuildFooterWithPageField(string halfPoints)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            var pageRun = new Run(
                new RunProperties(new FontSize { Val = halfPoints }),
                new Text("1"));
            footerPart.Footer = new Footer(
                new Paragraph(new SimpleField(pageRun) { Instruction = " PAGE " }));
            footerPart.Footer.Save();

            var sectPr = new SectionProperties();
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) });
            body.Append(new Paragraph(new Run(new Text("Body"))));
            body.Append(sectPr);

            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }
}

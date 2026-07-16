using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;
using V = DocumentFormat.OpenXml.Vml;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Pozioma linia Worda („Wstaw → linia pozioma"): w:pict → v:rect z o:hr="t" (VML).
/// Poprzednio reader gubił ją całkowicie (ConvertPictureToHtml znał tylko v:imagedata
/// i v:textbox) — linia znikała z podglądu ORAZ z pliku po pierwszym zapisie
/// (repro: „Tekst artykułu.docx", 41 linii-separatorów sekcji).
/// Kontrakt: reader → blokowy span.docx-hr + data-hr-* (hr w &lt;p&gt; jest niedozwolony —
/// parser rozbiłby akapit); writer odtwarza VML 1:1.
/// </summary>
[TestFixture]
public class HorizontalRuleFidelityTests
{
    private const string OfficeVmlNs = "urn:schemas-microsoft-com:office:office";

    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static V.Rectangle BuildHrRect()
    {
        var rect = new V.Rectangle { Style = "width:0;height:0", FillColor = "#a0a0a0", Stroked = false };
        rect.SetAttribute(new OpenXmlAttribute("o", "hralign", OfficeVmlNs, "center"));
        rect.SetAttribute(new OpenXmlAttribute("o", "hrstd", OfficeVmlNs, "t"));
        rect.SetAttribute(new OpenXmlAttribute("o", "hr", OfficeVmlNs, "t"));
        return rect;
    }

    private static MemoryStream BuildDocxWithHr()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Przed linią"))),
                new Paragraph(new Run(new Picture(BuildHrRect()))),
                new Paragraph(new Run(new Text("Po linii")))));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_VmlHorizontalRule_EmitsBlockSpanWithContract()
    {
        var html = _reader.Convert(BuildDocxWithHr()).Html;

        html.Should().Contain("class=\"docx-hr\"");
        html.Should().Contain("data-docx-hr=\"1\"");
        html.Should().Contain("data-hr-align=\"center\"");
        html.Should().Contain("data-hr-fill=\"#a0a0a0\"");
        html.Should().Contain("display:block");
        html.Should().Contain("background:#a0a0a0");
        html.Should().NotContain("<hr", "hr wewnątrz <p> rozbiłby akapit w przeglądarce");
        // Linia siedzi między akapitami treści.
        html.IndexOf("docx-hr", StringComparison.Ordinal).Should().BeGreaterThan(
            html.IndexOf("Przed linią", StringComparison.Ordinal));
        html.IndexOf("Po linii", StringComparison.Ordinal).Should().BeGreaterThan(
            html.IndexOf("docx-hr", StringComparison.Ordinal));
    }

    [Test]
    public void Write_DocxHrSpan_RecreatesVmlRectangle()
    {
        var html =
            "<p>Przed</p>" +
            "<p><span class=\"docx-hr\" data-docx-hr=\"1\" data-hr-align=\"center\" data-hr-std=\"t\" " +
            "data-hr-fill=\"#a0a0a0\" style=\"display:block;height:1px;background:#a0a0a0;\"></span></p>" +
            "<p>Po</p>";

        var docx = _writer.Convert(html);
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var rect = doc.MainDocumentPart!.Document!.Body!.Descendants<V.Rectangle>().Single();

        rect.Ancestors<Picture>().Should().NotBeEmpty("linia musi siedzieć w w:pict");
        rect.GetAttribute("hr", OfficeVmlNs).Value.Should().Be("t");
        rect.GetAttribute("hralign", OfficeVmlNs).Value.Should().Be("center");
        rect.GetAttribute("hrstd", OfficeVmlNs).Value.Should().Be("t");
        rect.FillColor!.Value.Should().Be("#a0a0a0");
    }

    [Test]
    public void RoundTrip_HorizontalRule_SurvivesSave()
    {
        var content = _reader.Convert(BuildDocxWithHr());
        var docx = _writer.Convert(content.Html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var rects = body.Descendants<V.Rectangle>().ToList();
        rects.Should().HaveCount(1, "linia nie może ani zniknąć, ani się zduplikować");
        rects[0].GetAttribute("hr", OfficeVmlNs).Value.Should().Be("t");
        body.InnerText.Should().Contain("Przed linią").And.Contain("Po linii");
    }
}

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Pola numeru strony (PAGE/NUMPAGES) muszą wracać do DOCX jako pola, nie jako literalny
/// tekst placeholdera „{page}"/„{pages}". Wcześniej ścieżka TREŚCI (body) emitowała
/// span.field-page/page-number jako zwykły run tekstowy → w Wordzie widać było „{page}".
/// </summary>
[TestFixture]
public class PageNumberFieldExportTests
{
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup() => _writer = new HtmlToDocxConverter();

    [Test]
    public void BodyPageField_ExportsAsSimpleField_NotLiteralPlaceholder()
    {
        var html = "<p>Strona <span class=\"field-page\">{page}</span> z "
                   + "<span class=\"field-numpages\">{pages}</span></p>";

        var body = ExportBody(html);

        body.Descendants<SimpleField>().Select(f => f.Instruction!.Value!.Trim())
            .Should().Contain("PAGE").And.Contain("NUMPAGES");
        body.InnerText.Should().NotContain("{page}");
        body.InnerText.Should().NotContain("{pages}");
    }

    [Test]
    public void BodyPageNumberSpan_ExportsAsPageField()
    {
        // Wstawka z toolbara: <span class="page-number">{page}</span>.
        var body = ExportBody("<p><span class=\"page-number\">{page}</span></p>");

        body.Descendants<SimpleField>().Should().ContainSingle()
            .Which.Instruction!.Value!.Trim().Should().Be("PAGE");
        body.InnerText.Should().NotContain("{page}");
    }

    [Test]
    public void FieldCachedValue_IsNeverThePlaceholderText()
    {
        var body = ExportBody("<p><span class=\"field-page\">{page}</span></p>");

        var field = body.Descendants<SimpleField>().Single();
        field.InnerText.Should().Be("1", "wartość zbuforowana pola musi być liczbą, nie placeholderem");
    }

    [Test]
    public void FooterPageField_StillExportsAsField()
    {
        var footer = new HeaderFooterContent
        {
            Html = "<span class=\"field-page\" style=\"font-size:8pt\">{page}</span> z "
                   + "<span class=\"field-numpages\" style=\"font-size:8pt\">{pages}</span>"
        };

        var bytes = _writer.Convert("<p>Body</p>", footer: footer);
        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var footerPart = doc.MainDocumentPart!.FooterParts.First();

        footerPart.Footer!.Descendants<SimpleField>().Select(f => f.Instruction!.Value!.Trim())
            .Should().Contain("PAGE").And.Contain("NUMPAGES");
        footerPart.Footer.InnerText.Should().NotContain("{page}");
    }

    [Test]
    public void HeaderPageField_ExportsAsField()
    {
        var header = new HeaderFooterContent
        {
            Html = "Strona <span class=\"field-page\">{page}</span> z "
                   + "<span class=\"field-numpages\">{pages}</span>"
        };

        var bytes = _writer.Convert("<p>Body</p>", header: header);
        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.First();

        headerPart.Header!.Descendants<SimpleField>().Select(f => f.Instruction!.Value!.Trim())
            .Should().Contain("PAGE").And.Contain("NUMPAGES");
        headerPart.Header.InnerText.Should().NotContain("{page}");
    }

    private Body ExportBody(string html)
    {
        var bytes = _writer.Convert(html);
        var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return doc.MainDocumentPart!.Document!.Body!;
    }
}

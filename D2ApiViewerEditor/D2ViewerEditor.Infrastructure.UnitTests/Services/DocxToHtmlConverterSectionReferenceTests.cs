using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Header/footer selection must follow the section's references (sectPr), not the
/// physical part order. A document with even/default/first parts (as produced by
/// Word) must render the DEFAULT one on ordinary pages — earlier code used
/// HeaderParts.FirstOrDefault() and could surface an empty even/first part instead,
/// dropping the real header/footer content (e.g. a logo or footer text).
/// </summary>
[TestFixture]
public class DocxToHtmlConverterSectionReferenceTests
{
    private DocxToHtmlConverter _converter = null!;

    [SetUp]
    public void Setup() => _converter = new DocxToHtmlConverter();

    private static MemoryStream BuildDocxWithReferences(bool titlePage, bool evenAndOdd = false)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            if (evenAndOdd)
            {
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(new EvenAndOddHeaders());
                settingsPart.Settings.Save();
            }

            // Add the EVEN part FIRST on purpose: HeaderParts.FirstOrDefault() would
            // return it, so this guards the regression where order (not type) decided.
            var evenHeader = AddHeader(mainPart, "EVEN-HEADER");
            var defaultHeader = AddHeader(mainPart, "DEFAULT-HEADER");
            var firstHeader = AddHeader(mainPart, "FIRST-HEADER");

            var evenFooter = AddFooter(mainPart, "EVEN-FOOTER");
            var defaultFooter = AddFooter(mainPart, "DEFAULT-FOOTER");
            var firstFooter = AddFooter(mainPart, "FIRST-FOOTER");

            var sectPr = new SectionProperties();
            sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(evenHeader) });
            sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultHeader) });
            sectPr.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstHeader) });
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(evenFooter) });
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultFooter) });
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstFooter) });
            if (titlePage) sectPr.Append(new TitlePage());
            sectPr.Append(new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static HeaderPart AddHeader(MainDocumentPart mainPart, string text)
    {
        var part = mainPart.AddNewPart<HeaderPart>();
        part.Header = new Header(new Paragraph(new Run(new Text(text))));
        part.Header.Save();
        return part;
    }

    private static FooterPart AddFooter(MainDocumentPart mainPart, string text)
    {
        var part = mainPart.AddNewPart<FooterPart>();
        part.Footer = new Footer(new Paragraph(new Run(new Text(text))));
        part.Footer.Save();
        return part;
    }

    [Test]
    public void Header_RendersDefaultReference_NotFirstPartByOrder()
    {
        using var stream = BuildDocxWithReferences(titlePage: false);

        var content = _converter.Convert(stream);

        content.Header!.Html.Should().Contain("DEFAULT-HEADER");
        content.Header.Html.Should().NotContain("EVEN-HEADER");
    }

    [Test]
    public void Footer_RendersDefaultReference_NotFirstPartByOrder()
    {
        using var stream = BuildDocxWithReferences(titlePage: false);

        var content = _converter.Convert(stream);

        content.Footer!.Html.Should().Contain("DEFAULT-FOOTER");
        content.Footer.Html.Should().NotContain("EVEN-FOOTER");
    }

    [Test]
    public void Header_WithoutTitlePage_DoesNotExposeFirstPageContent()
    {
        using var stream = BuildDocxWithReferences(titlePage: false);

        var content = _converter.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeFalse();
        content.Header.FirstPageHtml.Should().BeNull();
    }

    [Test]
    public void Header_WithTitlePage_ExposesFirstPageContent_KeepsDefaultForOrdinaryPages()
    {
        using var stream = BuildDocxWithReferences(titlePage: true);

        var content = _converter.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeTrue();
        content.Header.FirstPageHtml.Should().Contain("FIRST-HEADER");
        content.Header.Html.Should().Contain("DEFAULT-HEADER");
    }

    [Test]
    public void Footer_WithTitlePage_ExposesFirstPageContent()
    {
        using var stream = BuildDocxWithReferences(titlePage: true);

        var content = _converter.Convert(stream);

        content.Footer!.DifferentFirstPage.Should().BeTrue();
        content.Footer.FirstPageHtml.Should().Contain("FIRST-FOOTER");
    }

    [Test]
    public void Header_WithoutEvenAndOddHeaders_DoesNotExposeEvenContent()
    {
        using var stream = BuildDocxWithReferences(titlePage: false, evenAndOdd: false);

        var content = _converter.Convert(stream);

        content.Header!.DifferentOddEven.Should().BeFalse();
        content.Header.EvenHtml.Should().BeNull();
    }

    [Test]
    public void Header_WithEvenAndOddHeaders_ExposesEvenContent_KeepsDefaultForOddPages()
    {
        using var stream = BuildDocxWithReferences(titlePage: false, evenAndOdd: true);

        var content = _converter.Convert(stream);

        content.Header!.DifferentOddEven.Should().BeTrue();
        content.Header.EvenHtml.Should().Contain("EVEN-HEADER");
        content.Header.Html.Should().Contain("DEFAULT-HEADER");
    }

    [Test]
    public void Footer_WithEvenAndOddHeaders_ExposesEvenContent()
    {
        using var stream = BuildDocxWithReferences(titlePage: false, evenAndOdd: true);

        var content = _converter.Convert(stream);

        content.Footer!.DifferentOddEven.Should().BeTrue();
        content.Footer.EvenHtml.Should().Contain("EVEN-FOOTER");
    }
}

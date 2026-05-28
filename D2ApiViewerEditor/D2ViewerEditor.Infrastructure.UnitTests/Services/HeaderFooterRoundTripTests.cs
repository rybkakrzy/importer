using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Round-trip of header/footer variants: HtmlToDocx must emit separate parts
/// with type=First/Even references and set titlePg / evenAndOddHeaders, so that
/// a subsequent DocxToHtml pass surfaces the same DifferentFirstPage / FirstPageHtml
/// and DifferentOddEven / EvenHtml. Without this, editing first/even variants in
/// the UI would not persist (rule 10: no apparent editing).
/// </summary>
[TestFixture]
public class HeaderFooterRoundTripTests
{
    private HtmlToDocxConverter _writer = null!;
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup()
    {
        _writer = new HtmlToDocxConverter();
        _reader = new DocxToHtmlConverter();
    }

    private static HeaderFooterContent Hf(string html,
        bool differentFirstPage = false, string? firstPageHtml = null,
        bool differentOddEven = false, string? evenHtml = null) =>
        new()
        {
            Html = html,
            Height = 1.25,
            DifferentFirstPage = differentFirstPage,
            FirstPageHtml = firstPageHtml,
            DifferentOddEven = differentOddEven,
            EvenHtml = evenHtml
        };

    [Test]
    public void DefaultOnly_RoundTrips_WithoutFirstOrEvenVariants()
    {
        var docx = _writer.Convert(
            "<p>body</p>",
            header: Hf("<p>HDR</p>"),
            footer: Hf("<p>FTR</p>"));

        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Header!.Html.Should().Contain("HDR");
        content.Header.DifferentFirstPage.Should().BeFalse();
        content.Header.DifferentOddEven.Should().BeFalse();
        content.Footer!.Html.Should().Contain("FTR");
    }

    [Test]
    public void FirstPageHeader_RoundTrips_AndSetsTitlePage()
    {
        var docx = _writer.Convert(
            "<p>body</p>",
            header: Hf("<p>DEFAULT-H</p>", differentFirstPage: true, firstPageHtml: "<p>FIRST-H</p>"));

        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeTrue();
        content.Header.FirstPageHtml.Should().Contain("FIRST-H");
        content.Header.Html.Should().Contain("DEFAULT-H");
    }

    [Test]
    public void FirstPageFooter_RoundTrips()
    {
        var docx = _writer.Convert(
            "<p>body</p>",
            footer: Hf("<p>DEFAULT-F</p>", differentFirstPage: true, firstPageHtml: "<p>FIRST-F</p>"));

        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Footer!.DifferentFirstPage.Should().BeTrue();
        content.Footer.FirstPageHtml.Should().Contain("FIRST-F");
    }

    [Test]
    public void EvenHeader_RoundTrips_AndSetsEvenAndOddHeaders()
    {
        var docx = _writer.Convert(
            "<p>body</p>",
            header: Hf("<p>DEFAULT-H</p>", differentOddEven: true, evenHtml: "<p>EVEN-H</p>"));

        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Header!.DifferentOddEven.Should().BeTrue();
        content.Header.EvenHtml.Should().Contain("EVEN-H");
        content.Header.Html.Should().Contain("DEFAULT-H");
    }

    [Test]
    public void EvenFooter_RoundTrips()
    {
        var docx = _writer.Convert(
            "<p>body</p>",
            footer: Hf("<p>DEFAULT-F</p>", differentOddEven: true, evenHtml: "<p>EVEN-F</p>"));

        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Footer!.DifferentOddEven.Should().BeTrue();
        content.Footer.EvenHtml.Should().Contain("EVEN-F");
    }

    [Test]
    public void AllVariantsCombined_RoundTrip_PreservesEachOne()
    {
        var header = Hf(
            "<p>DEFAULT-H</p>",
            differentFirstPage: true, firstPageHtml: "<p>FIRST-H</p>",
            differentOddEven: true, evenHtml: "<p>EVEN-H</p>");
        var footer = Hf(
            "<p>DEFAULT-F</p>",
            differentFirstPage: true, firstPageHtml: "<p>FIRST-F</p>",
            differentOddEven: true, evenHtml: "<p>EVEN-F</p>");

        var docx = _writer.Convert("<p>body</p>", header: header, footer: footer);
        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);

        content.Header!.Html.Should().Contain("DEFAULT-H");
        content.Header.FirstPageHtml.Should().Contain("FIRST-H");
        content.Header.EvenHtml.Should().Contain("EVEN-H");
        content.Footer!.Html.Should().Contain("DEFAULT-F");
        content.Footer.FirstPageHtml.Should().Contain("FIRST-F");
        content.Footer.EvenHtml.Should().Contain("EVEN-F");
    }
}

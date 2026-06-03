using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Etap 6: image alt text round-trip. The editor keeps alt on the &lt;img&gt;; the writer
/// maps it to wp:docPr/@descr, the reader mirrors it back (@descr preferred, else @title).
/// Images without alt stay unchanged (no empty alt attribute is emitted).
/// </summary>
[TestFixture]
public class ImageAltTextRoundTripTests
{
    private HtmlToDocxConverter _writer = null!;
    private DocxToHtmlConverter _reader = null!;

    private const string PngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    [SetUp]
    public void Setup()
    {
        _writer = new HtmlToDocxConverter();
        _reader = new DocxToHtmlConverter();
    }

    private string ConvertAndReadBackImg(string html)
    {
        var docx = _writer.Convert(html);
        using var stream = new MemoryStream(docx);
        return _reader.Convert(stream).Html;
    }

    [Test]
    public void AltText_RoundTrips_ToDocPrDescriptionAndBack()
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " data-width-emu=\"952500\" data-height-emu=\"476250\""
            + " alt=\"Logo firmy\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain("alt=\"Logo firmy\"");
    }

    [Test]
    public void AltText_WithSpecialCharacters_IsEscaped()
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " alt=\"Faktura &amp; rachunek &lt;2024&gt;\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain("Faktura &amp; rachunek &lt;2024&gt;");
        roundTripped.Should().NotContain("<2024>");
    }

    [Test]
    public void NoAltText_EmitsNoAltAttribute()
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain("<img");
        roundTripped.Should().NotContain(" alt=");
    }
}

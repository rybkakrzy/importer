using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Border (a:ln) + crop (a:srcRect) round-trip. Editor persists border and crop on the
/// &lt;img&gt; via data-border-* and data-crop-* attributes; the writer maps them to
/// pic:spPr/a:ln and pic:blipFill/a:srcRect, the reader mirrors them back.
/// </summary>
[TestFixture]
public class ImageBorderCropRoundTripTests
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

    [TestCase(2, "FF0000", "solid")]
    [TestCase(5, "00AAFF", "dashed")]
    [TestCase(1, "808080", "dotted")]
    public void Handle_Border_RoundTrip_PreservesWidthColorAndStyle(int widthPx, string colorHex, string style)
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " data-width-emu=\"952500\" data-height-emu=\"476250\""
            + $" data-border-width=\"{widthPx}\""
            + $" data-border-color=\"#{colorHex}\""
            + $" data-border-style=\"{style}\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain($"data-border-width=\"{widthPx}\"");
        roundTripped.Should().Contain($"data-border-color=\"#{colorHex}\"");
        roundTripped.Should().Contain($"data-border-style=\"{style}\"");
    }

    [Test]
    public void Handle_BorderWithBadColor_IsDropped()
    {
        // Invalid hex must not crash, must not corrupt the file — border is silently dropped.
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " data-width-emu=\"952500\" data-height-emu=\"476250\""
            + " data-border-width=\"3\""
            + " data-border-color=\"not-a-hex\""
            + " data-border-style=\"solid\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().NotContain("data-border-width");
        roundTripped.Should().NotContain("data-border-color");
    }

    [TestCase(10, 5, 8, 12)]
    [TestCase(25, 25, 0, 0)]
    public void Handle_Crop_RoundTrip_PreservesPercentages(int l, int r, int t, int b)
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " data-width-emu=\"952500\" data-height-emu=\"476250\""
            + $" data-crop-l=\"{l}\" data-crop-r=\"{r}\""
            + $" data-crop-t=\"{t}\" data-crop-b=\"{b}\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        if (l == 0 && r == 0 && t == 0 && b == 0)
        {
            roundTripped.Should().NotContain("data-crop");
            return;
        }
        roundTripped.Should().Contain($"data-crop-l=\"{l}\"");
        roundTripped.Should().Contain($"data-crop-r=\"{r}\"");
        roundTripped.Should().Contain($"data-crop-t=\"{t}\"");
        roundTripped.Should().Contain($"data-crop-b=\"{b}\"");
    }

    [Test]
    public void Handle_PlainImage_HasNoBorderOrCropAttrsAfterRoundTrip()
    {
        var html =
            $"<p><img src=\"data:image/png;base64,{PngBase64}\""
            + " style=\"width:100px;height:50px;\""
            + " data-width-emu=\"952500\" data-height-emu=\"476250\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().NotContain("data-border");
        roundTripped.Should().NotContain("data-crop");
    }
}

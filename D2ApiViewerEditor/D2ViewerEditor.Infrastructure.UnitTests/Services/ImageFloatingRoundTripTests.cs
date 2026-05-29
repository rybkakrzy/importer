using System.Text.RegularExpressions;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Word-like image positioning round-trip. The editor marks an image as floating by
/// setting data-pos-mode="front"|"behind" on the &lt;img&gt; (with data-x-emu / data-y-emu
/// offsets). On save HtmlToDocxConverter must emit wp:anchor with the matching behindDoc
/// flag; DocxToHtmlConverter must read it back into the same data attributes.
/// </summary>
[TestFixture]
public class ImageFloatingRoundTripTests
{
    private HtmlToDocxConverter _writer = null!;
    private DocxToHtmlConverter _reader = null!;

    // 1×1 PNG, base64 — smallest legal image so the converters produce a real ImagePart.
    private const string PngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    [SetUp]
    public void Setup()
    {
        _writer = new HtmlToDocxConverter();
        _reader = new DocxToHtmlConverter();
    }

    private static string BuildHtml(string posMode, long xEmu, long yEmu, int widthPx = 100, int heightPx = 50) =>
        "<p>before</p>"
        + $"<p><img src=\"data:image/png;base64,{PngBase64}\""
        + $" style=\"max-width:100%;width:{widthPx}px;height:{heightPx}px;\""
        + $" data-width-emu=\"{widthPx * 9525}\" data-height-emu=\"{heightPx * 9525}\""
        + $" data-pos-mode=\"{posMode}\" data-x-emu=\"{xEmu}\" data-y-emu=\"{yEmu}\" /></p>"
        + "<p>after</p>";

    private string ConvertAndReadBackImg(string html)
    {
        var docx = _writer.Convert(html);
        using var stream = new MemoryStream(docx);
        var content = _reader.Convert(stream);
        return content.Html;
    }

    [TestCase("front", 762000L, 381000L, false)]
    [TestCase("behind", 1524000L, 762000L, true)]
    public void Handle_FloatingImage_RoundTrip_PreservesPositionAndMode(
        string posMode, long xEmu, long yEmu, bool expectBehind)
    {
        var html = BuildHtml(posMode, xEmu, yEmu);

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain($"data-pos-mode=\"{posMode}\"");
        roundTripped.Should().Contain($"data-x-emu=\"{xEmu}\"");
        roundTripped.Should().Contain($"data-y-emu=\"{yEmu}\"");
        // Sanity: behindDoc determines layering, captured in posMode round-trip.
        if (expectBehind)
            roundTripped.Should().Contain("data-pos-mode=\"behind\"");
        else
            roundTripped.Should().Contain("data-pos-mode=\"front\"");
    }

    [Test]
    public void Handle_InlineImageWithoutPosMode_StillRoundTripsWithoutFloatingAttributes()
    {
        var html =
            "<p><img src=\"data:image/png;base64," + PngBase64 + "\""
            + " style=\"max-width:100%;width:120px;height:60px;\""
            + " data-width-emu=\"1143000\" data-height-emu=\"571500\" /></p>";

        var roundTripped = ConvertAndReadBackImg(html);

        roundTripped.Should().Contain("<img");
        roundTripped.Should().NotContain("data-pos-mode");
    }

    [Test]
    public void Handle_FloatingImage_SizeRoundTripsAlongsidePosition()
    {
        var html = BuildHtml("front", 200000, 100000, widthPx: 240, heightPx: 80);

        var roundTripped = ConvertAndReadBackImg(html);

        // Width/height are emitted as px (after EmuToPx round-trip).
        var widthMatch = Regex.Match(roundTripped, @"width:(\d+)px");
        var heightMatch = Regex.Match(roundTripped, @"height:(\d+)px");
        widthMatch.Success.Should().BeTrue();
        heightMatch.Success.Should().BeTrue();
        int.Parse(widthMatch.Groups[1].Value).Should().BeInRange(238, 242);
        int.Parse(heightMatch.Groups[1].Value).Should().BeInRange(78, 82);
    }
}

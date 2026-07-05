using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Pins the OOXML line-spacing → CSS mapping (Issue 5). Word's w:spacing maps as:
///  - line + lineRule=auto  → unitless line-height (value/240: 240=single, 360=1.5, 480=double)
///  - line + lineRule=exact/atLeast → line-height in pt (twips/20)
///  - before/after (twips)  → margin-top/bottom in pt
/// The mapping is correct; remaining Word differences (font line-gap in "single", CSS adjacent-
/// margin collapsing) are inherent HTML/CSS limits documented in .ai/DOCX_CONVERSION.md.
/// </summary>
[TestFixture]
public class LineSpacingMappingTests
{
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup() => _reader = new DocxToHtmlConverter();

    private string ParagraphCss(SpacingBetweenLines spacing)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var p = new Paragraph(
                new ParagraphProperties(spacing),
                new Run(new Text("Tekst")));
            main.Document = new Document(new Body(p));
            main.Document.Save();
        }
        var html = _reader.Convert(new MemoryStream(ms.ToArray())).Html;
        var m = System.Text.RegularExpressions.Regex.Match(html, "<p[^>]*style=\"([^\"]*)\"");
        return m.Success ? m.Groups[1].Value : html;
    }

    [Test]
    public void SingleSpacing_MapsToUnitlessLineHeight1()
        => ParagraphCss(new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto })
            .Should().Contain("line-height:1;");

    [Test]
    public void OneAndHalfSpacing_MapsToLineHeight1_5()
        => ParagraphCss(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto })
            .Should().Contain("line-height:1.5;");

    [Test]
    public void DoubleSpacing_MapsToLineHeight2()
        => ParagraphCss(new SpacingBetweenLines { Line = "480", LineRule = LineSpacingRuleValues.Auto })
            .Should().Contain("line-height:2;");

    [Test]
    public void ExactSpacing_MapsToPointLineHeight()
        => ParagraphCss(new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Exact })
            .Should().Contain("line-height:12pt;");

    [Test]
    public void AtLeastSpacing_MapsToPointLineHeight()
        => ParagraphCss(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.AtLeast })
            .Should().Contain("line-height:18pt;");

    [Test]
    public void AtLeastSpacing_IsMarkedForRoundTrip_ExactIsNot()
    {
        // Bez markera writer mapował KAŻDE pt-owe line-height na w:lineRule=exact;
        // atLeast→exact przycina w Wordzie tekst wyższy niż linia.
        ParagraphCss(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.AtLeast })
            .Should().Contain("--w-line-rule:atLeast;");
        ParagraphCss(new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Exact })
            .Should().NotContain("--w-line-rule");
    }

    [Test]
    public void AtLeastSpacing_RoundTripsBackToAtLeastRule()
    {
        var writer = new HtmlToDocxConverter();

        var bytes = writer.Convert("<p style=\"line-height:18pt;--w-line-rule:atLeast;\">Tekst</p>");

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var spacing = doc.MainDocumentPart!.Document!.Body!
            .Descendants<SpacingBetweenLines>().Single();
        spacing.LineRule!.Value.Should().Be(LineSpacingRuleValues.AtLeast);
        spacing.Line!.Value.Should().Be("360");
    }

    [Test]
    public void ExactSpacing_RoundTripsBackToExactRule()
    {
        var writer = new HtmlToDocxConverter();

        var bytes = writer.Convert("<p style=\"line-height:18pt;\">Tekst</p>");

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var spacing = doc.MainDocumentPart!.Document!.Body!
            .Descendants<SpacingBetweenLines>().Single();
        spacing.LineRule!.Value.Should().Be(LineSpacingRuleValues.Exact);
    }

    [Test]
    public void SpaceBeforeAfter_MapToMarginsInPoints()
    {
        var css = ParagraphCss(new SpacingBetweenLines { Before = "240", After = "200" });
        css.Should().Contain("margin-top:12pt;");
        css.Should().Contain("margin-bottom:10pt;");
    }
}

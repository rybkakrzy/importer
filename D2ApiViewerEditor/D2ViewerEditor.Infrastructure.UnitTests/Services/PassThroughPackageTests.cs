using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// R-16: pass-through writer must stop discarding the original DOCX parts it doesn't edit.
/// A from-scratch save regenerated a minimal styles.xml (~16 styles vs 164), dropped table
/// styles and swapped the body font (Cambria→Calibri). <c>ConvertPreservingPackage</c> keeps
/// the original styles.xml, theme and fontTable while still writing the edited body.
/// </summary>
[TestFixture]
public class PassThroughPackageTests
{
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup() => _writer = new HtmlToDocxConverter();

    /// <summary>Builds an original package with many styles (incl. table styles), a Cambria theme
    /// and docDefaults that resolve the body font through the theme (asciiTheme=minorHAnsi).</summary>
    private static MemoryStream BuildRichOriginal(int paragraphStyles, int tableStyles)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Oryginał")))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi },
                new FontSize { Val = "22" }))));
            styles.Append(new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            for (var i = 0; i < paragraphStyles; i++)
                styles.Append(new Style(new StyleName { Val = $"P{i}" }) { Type = StyleValues.Paragraph, StyleId = $"P{i}" });
            for (var i = 0; i < tableStyles; i++)
                styles.Append(new Style(new StyleName { Val = $"T{i}" }) { Type = StyleValues.Table, StyleId = $"T{i}" });
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();

            var themePart = mainPart.AddNewPart<ThemePart>();
            using (var w = new StreamWriter(themePart.GetStream(FileMode.Create)))
                w.Write(ThemeXml("Cambria"));

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static string ThemeXml(string minorLatin) =>
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\">" +
        "<a:themeElements><a:clrScheme name=\"C\">" +
        string.Concat(new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" }
            .Select(c => $"<a:{c}><a:srgbClr val=\"000000\"/></a:{c}>")) +
        "</a:clrScheme>" +
        "<a:fontScheme name=\"F\">" +
        "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        $"<a:minorFont><a:latin typeface=\"{minorLatin}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>" +
        "</a:fontScheme>" +
        "<a:fmtScheme name=\"S\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>" +
        "<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>" +
        "</a:themeElements></a:theme>";

    private static (int styleCount, string? minorFont, bool docDefaultsUsesTheme, string bodyText) Inspect(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var styleCount = main.StyleDefinitionsPart?.Styles?.Elements<Style>().Count() ?? 0;
        var minor = main.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
        var rpr = main.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts;
        var usesTheme = rpr?.AsciiTheme != null;
        var bodyText = string.Concat(main.Document.Body!.Descendants<Text>().Select(t => t.Text));
        return (styleCount, minor, usesTheme, bodyText);
    }

    [Test]
    public void PreservesFullStyleSet_NotReducedToMinimal()
    {
        using var original = BuildRichOriginal(paragraphStyles: 40, tableStyles: 100); // 141 incl. Normal
        var result = _writer.ConvertPreservingPackage("<p>Edytowana treść</p>", original);

        var (styleCount, _, _, bodyText) = Inspect(result);
        styleCount.Should().BeGreaterThanOrEqualTo(141); // not collapsed to ~16
        bodyText.Should().Contain("Edytowana treść"); // edited body present
    }

    [Test]
    public void PreservesTableStyles()
    {
        using var original = BuildRichOriginal(paragraphStyles: 5, tableStyles: 30);
        var result = _writer.ConvertPreservingPackage("<p>x</p>", original);

        using var ms = new MemoryStream(result);
        using var doc = WordprocessingDocument.Open(ms, false);
        var tableStyles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Count(s => s.Type?.Value == StyleValues.Table);
        tableStyles.Should().Be(30);
    }

    [Test]
    public void PreservesThemeFont_BodyStaysCambriaNotCalibri()
    {
        using var original = BuildRichOriginal(paragraphStyles: 5, tableStyles: 5);
        var result = _writer.ConvertPreservingPackage("<p>x</p>", original);

        var (_, minorFont, usesTheme, _) = Inspect(result);
        minorFont.Should().Be("Cambria");      // theme preserved
        usesTheme.Should().BeTrue();            // docDefaults still resolve via theme → Cambria
    }

    [Test]
    public void NullOriginal_FallsBackToSelfContainedConvert()
    {
        var result = _writer.ConvertPreservingPackage("<p>Bez oryginału</p>", null);

        var (styleCount, _, _, bodyText) = Inspect(result);
        styleCount.Should().BeGreaterThan(0);          // generated minimal styles
        bodyText.Should().Contain("Bez oryginału");
    }

    [Test]
    public void EditedBodyReplacesOriginalBody()
    {
        using var original = BuildRichOriginal(paragraphStyles: 3, tableStyles: 3);
        var result = _writer.ConvertPreservingPackage("<p>Nowa treść edytora</p>", original);

        var (_, _, _, bodyText) = Inspect(result);
        bodyText.Should().Contain("Nowa treść edytora");
        bodyText.Should().NotContain("Oryginał"); // original body content is gone
    }
}

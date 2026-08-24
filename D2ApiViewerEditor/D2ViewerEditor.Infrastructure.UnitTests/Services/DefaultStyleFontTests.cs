using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Reader must resolve the document's real body font. In Word the default paragraph style
/// (w:default="1", e.g. "Normalny" with rFonts ascii="Times New Roman") overrides docDefaults
/// (asciiTheme=minorHAnsi → theme minor). The reader previously read only docDefaults, so body
/// paragraphs without an explicit font fell back to the editor default (Calibri). The document
/// container must carry the default-style font. No font is hardcoded — it comes from the document.
/// </summary>
[TestFixture]
public class DefaultStyleFontTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static string Container(string html) =>
        System.Text.RegularExpressions.Regex.Match(html, "<div class=\"document-content\"[^>]*>").Value;

    [Test]
    public void DefaultParagraphStyleFont_OverridesDocDefaultsTheme()
    {
        // docDefaults → theme minor (Cambria); default style "Normalny" → Times New Roman.
        var docx = Build(
            docDefaultsTheme: true,
            defaultStyleAscii: "Times New Roman",
            themeMinor: "Cambria");

        var content = _reader.Convert(new MemoryStream(docx));

        Container(content.Html).Should().Contain("'Times New Roman'");
        Container(content.Html).Should().NotContain("Cambria");
    }

    [Test]
    public void SerifFont_FallsBackToSerif_NotSansSerif()
    {
        // A missing Times New Roman must render as a serif, not sans-serif (Calibri-like).
        var docx = Build(docDefaultsTheme: true, defaultStyleAscii: "Times New Roman", themeMinor: "Cambria");

        var content = _reader.Convert(new MemoryStream(docx));

        Container(content.Html).Should().Contain("'Times New Roman',serif");
    }

    [Test]
    public void SansSerifFont_KeepsSansSerifFallback()
    {
        var docx = Build(docDefaultsTheme: false, docDefaultsAscii: "Calibri", defaultStyleAscii: null, themeMinor: "Calibri");

        var content = _reader.Convert(new MemoryStream(docx));

        Container(content.Html).Should().Contain("'Calibri',sans-serif");
    }

    [Test]
    public void DocDefaultsFont_UsedWhenDefaultStyleHasNoFont()
    {
        var docx = Build(docDefaultsTheme: true, defaultStyleAscii: null, themeMinor: "Cambria");

        var content = _reader.Convert(new MemoryStream(docx));

        Container(content.Html).Should().Contain("Cambria"); // falls back to docDefaults/theme
    }

    [Test]
    public void DirectDocDefaultsFont_IsUsed()
    {
        var docx = Build(docDefaultsTheme: false, docDefaultsAscii: "Times New Roman", defaultStyleAscii: null, themeMinor: "Cambria");

        var content = _reader.Convert(new MemoryStream(docx));

        Container(content.Html).Should().Contain("'Times New Roman'");
    }

    [Test]
    public void PassThrough_KeepsDefaultStyleFont_NotCalibri()
    {
        var docx = Build(docDefaultsTheme: true, defaultStyleAscii: "Times New Roman", themeMinor: "Cambria");
        var content = _reader.Convert(new MemoryStream(docx));

        var result = _writer.ConvertPreservingPackage(content.Html, new MemoryStream(docx));

        using var ms = new MemoryStream(result);
        using var doc = WordprocessingDocument.Open(ms, false);
        var normal = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().First(s => s.Default?.Value == true && s.Type?.Value == StyleValues.Paragraph);
        normal.StyleRunProperties!.GetFirstChild<RunFonts>()!.Ascii!.Value.Should().Be("Times New Roman");
    }

    // --- builder --------------------------------------------------------------

    private static byte[] Build(bool docDefaultsTheme, string? defaultStyleAscii, string themeMinor,
        string? docDefaultsAscii = null)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Treść akapitu bez własnego fontu.")))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var ddFonts = docDefaultsTheme
                ? new RunFonts { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi }
                : new RunFonts { Ascii = docDefaultsAscii, HighAnsi = docDefaultsAscii };
            var styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(ddFonts, new FontSize { Val = "22" }))));

            var normalRunProps = new StyleRunProperties();
            if (defaultStyleAscii != null)
                normalRunProps.Append(new RunFonts { Ascii = defaultStyleAscii, HighAnsi = defaultStyleAscii });
            normalRunProps.Append(new FontSize { Val = "21" });
            styles.Append(new Style(new StyleName { Val = "Normalny" }, normalRunProps)
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normalny",
                Default = true
            });
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();

            mainPart.AddNewPart<ThemePart>().FeedXml(ThemeXml(themeMinor));

            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    private static string ThemeXml(string minorLatin) =>
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\">" +
        "<a:themeElements><a:clrScheme name=\"C\">" +
        string.Concat(new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" }
            .Select(c => $"<a:{c}><a:srgbClr val=\"000000\"/></a:{c}>")) +
        "</a:clrScheme><a:fontScheme name=\"F\">" +
        "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        $"<a:minorFont><a:latin typeface=\"{minorLatin}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
        "<a:fmtScheme name=\"S\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>" +
        "<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>" +
        "</a:themeElements></a:theme>";
}

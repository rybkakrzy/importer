using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Conversion;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Wierność renderu ZARAZ PO WCZYTANIU (reader → GUI), bez cyklu edycji:
///
/// 1. Dokument BEZ domyślnych odstępów (starsze szablony: Normal/docDefaults bez w:spacing)
///    musi emitować JAWNE data-default-before-tw="0"/after-tw="0". Brak atrybutu znaczył
///    „nie wiadomo": GUI dodawało fallback 10px po każdym akapicie (dokument „rozstrzelony"
///    względem Worda już na wejściu), a writer przy zapisie wstrzykiwał hardkod
///    after=160/line=259 do docDefaults — dokument dostawał odstępy, których nigdy nie miał.
///
/// 2. Dokument bez domyślnej interlinii = pojedynczy odstęp Worda; kontener musi nieść
///    skalibrowane line-height (np. Calibri 1.221), bo inaczej GUI spada na 1.15 edytora
///    (~6% ciaśniej niż Word, rozjazd rośnie z każdą linią).
///
/// 3. Interlinia auto kalibruje się po foncie AKAPITU/STYLU, nie foncie domyślnym dokumentu
///    (PG-09): akapit w Times (single 1.149) w dokumencie Calibri (1.221) dostawał zły
///    współczynnik.
/// </summary>
[TestFixture]
public class DefaultSpacingAndLineFontFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    /// <summary>DOCX à la stary szablon: docDefaults tylko font, ZERO domyślnych odstępów.</summary>
    private static MemoryStream BuildLegacyNoDefaultsDocx(params OpenXmlElement[] bodyChildren)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                        new FontSize { Val = "22" }))),
                new Style(new StyleName { Val = "Normal" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    Default = true
                },
                new Style(
                    new StyleName { Val = "TimesHeading" },
                    new StyleParagraphProperties(
                        new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto }),
                    new StyleRunProperties(new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "TimesHeading"
                });
            stylesPart.Styles.Save();

            foreach (var child in bodyChildren)
                body.Append(child);
            body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static string ContainerTag(string html)
    {
        var start = html.IndexOf("<div class=\"document-content\"", StringComparison.Ordinal);
        start.Should().BeGreaterThanOrEqualTo(0, "reader powinien emitować kontener document-content");
        var end = html.IndexOf('>', start);
        return html.Substring(start, end - start + 1);
    }

    // ---------- reader: jawne zera domyślnych odstępów ----------

    [Test]
    public void Read_NoDefaultSpacing_EmitsExplicitZeroDataAttributes()
    {
        using var docx = BuildLegacyNoDefaultsDocx(
            new Paragraph(new Run(new Text("Lorem ipsum"))));
        var container = ContainerTag(_reader.Convert(docx).Html);

        container.Should().Contain("data-default-before-tw=\"0\"",
            "brak atrybutu uruchamiał fallbacki GUI/writera niezgodne z dokumentem");
        container.Should().Contain("data-default-after-tw=\"0\"");
    }

    [Test]
    public void Read_NoDefaultLine_ContainerCarriesCalibratedSingleLineHeight()
    {
        using var docx = BuildLegacyNoDefaultsDocx(
            new Paragraph(new Run(new Text("Lorem ipsum"))));
        var container = ContainerTag(_reader.Convert(docx).Html);

        // Pojedynczy odstęp Calibri wg tabeli metryk PG-09, nie 1.15 edytora.
        var expected = ExtractLineHeight(WordLineSpacing.AutoCss(240, "Calibri"));
        container.Should().Contain($"line-height:{expected};");
    }

    // ---------- reader: kalibracja interlinii po foncie akapitu/stylu ----------

    [Test]
    public void Read_AutoLine_CalibratedByParagraphRunFont_NotDocumentDefault()
    {
        // Akapit 1.5 wiersza w Times (runy z jawnym rFonts) w dokumencie domyślnie Calibri.
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto }),
            new Run(
                new RunProperties(new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" }),
                new Text("Times akapit")));
        using var docx = BuildLegacyNoDefaultsDocx(paragraph);
        var html = _reader.Convert(docx).Html;

        var timesCss = WordLineSpacing.AutoCss(360, "Times New Roman");
        var calibriCss = WordLineSpacing.AutoCss(360, "Calibri");
        html.Should().Contain($"line-height:{ExtractLineHeight(timesCss)};");
        html.Should().NotContain($"line-height:{ExtractLineHeight(calibriCss)};",
            "kalibracja po foncie domyślnym dokumentu zawyżała interlinię akapitu w Times");
        html.Should().Contain("--w-line-tw:360;", "marker round-trip musi zostać");
    }

    [Test]
    public void Read_AutoLineInStyle_CalibratedByStyleFont()
    {
        // Interlinia zdefiniowana w STYLU z własnym fontem — kalibracja po foncie stylu.
        var paragraph = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "TimesHeading" }),
            new Run(new Text("Styl w Times")));
        using var docx = BuildLegacyNoDefaultsDocx(paragraph);
        var html = _reader.Convert(docx).Html;

        html.Should().Contain($"line-height:{ExtractLineHeight(WordLineSpacing.AutoCss(360, "Times New Roman"))};");
    }

    // ---------- writer: round-trip jawnych zer ----------

    [Test]
    public void Write_ExplicitZeroDefaults_DoNotBecomeHardcodedWordDefaults()
    {
        var html = "<div class=\"document-content\" data-default-before-tw=\"0\" data-default-after-tw=\"0\"" +
                   " style=\"font-family:'Calibri',sans-serif;font-size:11pt;\">" +
                   "<p style=\"\">Tekst</p></div>";
        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var spacing = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!
            .GetFirstChild<SpacingBetweenLines>()!;
        spacing.Before!.Value.Should().Be("0");
        spacing.After!.Value.Should().Be("0");
        spacing.Line.Should().BeNull("dokument bez domyślnej interlinii nie może dostać hardkodu 259");
    }

    [Test]
    public void RoundTrip_LegacyNoDefaults_KeepsZeroSpacingThroughSave()
    {
        using var docx = BuildLegacyNoDefaultsDocx(
            new Paragraph(new Run(new Text("Lorem ipsum"))));
        var pass1 = _reader.Convert(docx);
        var export = _writer.Convert(pass1.Html, pass1.Metadata, pass1.Header, pass1.Footer, pass1.Margins, pass1.PageSize);

        using var exported = WordprocessingDocument.Open(new MemoryStream(export), false);
        var spacing = exported.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!
            .GetFirstChild<SpacingBetweenLines>()!;
        spacing.Before!.Value.Should().Be("0",
            "dokument bez odstępów nie może po zapisie dostać odstępów, których nigdy nie miał");
        spacing.After!.Value.Should().Be("0");
    }

    private static string ExtractLineHeight(string autoCss)
    {
        var match = System.Text.RegularExpressions.Regex.Match(autoCss, @"line-height:([\d.]+);");
        match.Success.Should().BeTrue();
        return match.Groups[1].Value;
    }
}

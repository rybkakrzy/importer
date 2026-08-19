using System.IO;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Regresja realnego buga „rozmiar czcionki z Worda (np. 14) ładuje się jako domyślny (np. 10)".
/// Sprawdzamy rozwiązanie rozmiaru na każdym poziomie dziedziczenia OOXML — który zawiedzie,
/// ten jest źródłem buga. Word zapisuje 14pt jako `w:sz w:val="28"` (half-points).
/// </summary>
[TestFixture]
public class FontSizeImportTests
{
    [Test]
    public void DirectRunFontSize_IsPreserved()
    {
        var docx = BuildDocx(stylesXml: null, body: new Body(
            new Paragraph(
                new Run(
                    new RunProperties(new FontSize { Val = "28" }),
                    new Text("Tekst 14pt")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("font-size:14pt");
    }

    [Test]
    public void ParagraphStyleFontSize_IsAppliedToParagraph()
    {
        // Styl akapitowy „Body14" z rozmiarem 14pt; run BEZ własnego sz dziedziczy ze stylu.
        var styles = new Styles(
            NormalStyle(),
            new Style(
                new StyleName { Val = "Body14" },
                new StyleRunProperties(new FontSize { Val = "28" }))
            { Type = StyleValues.Paragraph, StyleId = "Body14" });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Body14" }),
                new Run(new Text("Tekst stylowy")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("font-size:14pt");
    }

    [Test]
    public void DocDefaultsFontSize_IsAppliedToContainer()
    {
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "28" }))),
            NormalStyle());

        var docx = BuildDocx(styles, new Body(
            new Paragraph(new Run(new Text("Tekst domyślny")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("document-content");
        html.Should().Contain("font-size:14pt");
    }

    [Test]
    public void DefaultParagraphStyleFontSize_BeatsDocDefaults()
    {
        // docDefaults = 11pt, ale styl Normal (default) = 14pt → Normal wygrywa (bardziej szczegółowy).
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))),
            new Style(
                new StyleName { Val = "Normal" },
                new StyleRunProperties(new FontSize { Val = "28" }))
            { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(new Run(new Text("Tekst")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("font-size:14pt");
        html.Should().NotContain("font-size:11pt");
    }

    [Test]
    public void HeadingWithoutStyleSize_CarriesExplicitDocumentDefaultSize()
    {
        // Firmowy wzorzec: Heading1 basedOn Normal, tylko bold, BEZ w:sz — Word renderuje go
        // rozmiarem dokumentu (11pt), a prezentacyjne h1 z SCSS edytora pokazywało 24pt.
        // Reader musi emitować jawny inline font-size, żeby wygrać z klasą.
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))),
            NormalStyle(),
            new Style(
                new StyleName { Val = "heading 1" },
                new BasedOn { Val = "Normal" },
                new StyleRunProperties(new Bold()))
            { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text("Nagłówek bez rozmiaru")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        var h1 = System.Text.RegularExpressions.Regex.Match(html, "<h1[^>]*>");
        h1.Success.Should().BeTrue();
        h1.Value.Should().Contain("font-size:11pt");
    }

    [Test]
    public void HeadingWithStyleSize_KeepsStyleSize()
    {
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))),
            NormalStyle(),
            new Style(
                new StyleName { Val = "heading 1" },
                new BasedOn { Val = "Normal" },
                new StyleRunProperties(new FontSize { Val = "32" }))
            { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text("Nagłówek 16pt")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        var h1 = System.Text.RegularExpressions.Regex.Match(html, "<h1[^>]*>");
        h1.Success.Should().BeTrue();
        h1.Value.Should().Contain("font-size:16pt");
        h1.Value.Should().NotContain("font-size:11pt");
    }

    [Test]
    public void DefaultStyleFontSize_ResolvedThroughBasedOnChain()
    {
        // Normal (default) bez w:sz, basedOn „Base" z 14pt — Word dziedziczy do docDefaults,
        // bezpośredni odczyt rPr Normal gubił rozmiar.
        var styles = new Styles(
            new Style(
                new StyleName { Val = "Base" },
                new StyleRunProperties(new FontSize { Val = "28" }))
            { Type = StyleValues.Paragraph, StyleId = "Base" },
            new Style(
                new StyleName { Val = "Normal" },
                new BasedOn { Val = "Base" })
            { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(new Run(new Text("Tekst")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("font-size:14pt");
    }

    [Test]
    public void TableStyleRunFontSize_AppliesToCellParagraphs()
    {
        // w:rPr stylu TABELI (np. 9pt) obowiązuje tekst komórek — dotąd ignorowany,
        // komórki spadały na rozmiar dokumentu.
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))),
            NormalStyle(),
            new Style(
                new StyleName { Val = "TabelaMala" },
                new StyleRunProperties(new FontSize { Val = "18" }))
            { Type = StyleValues.Table, StyleId = "TabelaMala" });

        var docx = BuildDocx(styles, new Body(
            new Table(
                new TableProperties(new TableStyle { Val = "TabelaMala" }),
                new TableGrid(new GridColumn()),
                new TableRow(
                    new TableCell(
                        new Paragraph(new Run(new Text("Tekst komórki"))))))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        var cellP = System.Text.RegularExpressions.Regex.Match(html, "<td[^>]*>.*?<p[^>]*>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        cellP.Success.Should().BeTrue();
        cellP.Value.Should().Contain("font-size:9pt");
    }

    [Test]
    public void CharacterStyleVertAlign_WithDirectVertAlign_DoesNotDoubleShrink()
    {
        // Styl znakowy z vertAlign daje CSS „vertical-align:super;font-size:smaller";
        // run z DODATKOWO bezpośrednim w:vertAlign dostaje semantyczny <sup>. Bez strip
        // CSS-a indeks kurczył się podwójnie (0.83 × 0.83).
        var styles = new Styles(
            NormalStyle(),
            new Style(
                new StyleName { Val = "IndeksGorny" },
                new StyleRunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }))
            { Type = StyleValues.Character, StyleId = "IndeksGorny" });

        var docx = BuildDocx(styles, new Body(
            new Paragraph(
                new Run(
                    new RunProperties(
                        new RunStyle { Val = "IndeksGorny" },
                        new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                    new Text("2")))));

        var html = new DocxToHtmlConverter().Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("<sup>");
        html.Should().NotContain("font-size:smaller", "semantyczny <sup> sam niesie zmniejszenie");
    }

    private static Style NormalStyle() => new Style(
        new StyleName { Val = "Normal" })
    { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };

    private static byte[] BuildDocx(Styles? stylesXml, Body body)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(body);
            if (stylesXml != null)
            {
                var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = stylesXml;
                stylesPart.Styles.Save();
            }
            main.Document.Save();
        }
        return ms.ToArray();
    }
}

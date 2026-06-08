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

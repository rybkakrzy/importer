using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Testy odwzorowania nagłówka/stopki DOCX → HTML. Skupiają się na regresji
/// „za duży tekst" (brak dziedziczenia domyślnego rozmiaru z docDefaults) oraz
/// na zachowaniu jawnego rozmiaru/koloru runu.
/// </summary>
[TestFixture]
public class DocxToHtmlConverterHeaderFooterTests
{
    private DocxToHtmlConverter _converter = null!;

    [SetUp]
    public void Setup() => _converter = new DocxToHtmlConverter();

    /// <summary>
    /// Buduje minimalny DOCX w pamięci. `docDefaultHalfPoints` = w:sz w docDefaults
    /// (half-points; null = brak docDefaults). Nagłówek/stopka zawierają run bez
    /// własnego rPr oraz (opcjonalnie) run z jawnym kolorem i rozmiarem.
    /// </summary>
    private static MemoryStream BuildDocx(int? docDefaultHalfPoints,
        string? explicitColorHex = null, int? explicitRunHalfPoints = null)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body")))));

            // styles.xml z docDefaults/rPrDefault
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var rPrDefault = new RunPropertiesDefault();
            if (docDefaultHalfPoints.HasValue)
            {
                rPrDefault.RunPropertiesBaseStyle = new RunPropertiesBaseStyle(
                    new FontSize { Val = docDefaultHalfPoints.Value.ToString() });
            }
            stylesPart.Styles = new Styles(new DocDefaults(rPrDefault));
            stylesPart.Styles.Save();

            Run BuildExplicitRun()
            {
                var rp = new RunProperties();
                if (explicitColorHex != null) rp.Append(new Color { Val = explicitColorHex });
                if (explicitRunHalfPoints.HasValue) rp.Append(new FontSize { Val = explicitRunHalfPoints.Value.ToString() });
                return new Run(rp, new Text("Styled"));
            }

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text("Plain")), BuildExplicitRun()));
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text("Plain")), BuildExplicitRun()));
            footerPart.Footer.Save();

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Header_WrapsContentInDefaultContainer_WithDocDefaultFontSize()
    {
        // docDefault 20 half-points = 10pt
        using var stream = BuildDocx(docDefaultHalfPoints: 20);

        var content = _converter.Convert(stream);

        content.Header.Should().NotBeNull();
        content.Header!.Html.Should().Contain("header-footer-content");
        // Kluczowa naprawa „za dużego tekstu": kontener niesie domyślny rozmiar 10pt,
        // więc run bez własnego w:sz dziedziczy go zamiast domyślnego rozmiaru edytora.
        content.Header.Html.Should().Contain("font-size:10pt");
    }

    [Test]
    public void Footer_WrapsContentInDefaultContainer_WithDocDefaultFontSize()
    {
        using var stream = BuildDocx(docDefaultHalfPoints: 20);

        var content = _converter.Convert(stream);

        content.Footer.Should().NotBeNull();
        content.Footer!.Html.Should().Contain("header-footer-content");
        content.Footer.Html.Should().Contain("font-size:10pt");
    }

    [Test]
    public void Header_PreservesExplicitRunColorAndFontSize()
    {
        // Jawny run: kolor 1F3864, rozmiar 16 half-points = 8pt.
        using var stream = BuildDocx(docDefaultHalfPoints: 20,
            explicitColorHex: "1F3864", explicitRunHalfPoints: 16);

        var content = _converter.Convert(stream);

        content.Header!.Html.Should().Contain("color:#1F3864");
        content.Header.Html.Should().Contain("font-size:8pt");
    }

    [Test]
    public void Footer_PreservesExplicitRunColor()
    {
        using var stream = BuildDocx(docDefaultHalfPoints: 22, explicitColorHex: "C00000");

        var content = _converter.Convert(stream);

        content.Footer!.Html.Should().Contain("color:#C00000");
    }

    [Test]
    public void Header_WithoutDocDefaults_StillWrapsInContainer_NoForcedFontSize()
    {
        // Brak docDefaults → kontener bez font-size; rozmiar zapewnia fallback CSS frontu.
        using var stream = BuildDocx(docDefaultHalfPoints: null);

        var content = _converter.Convert(stream);

        content.Header!.Html.Should().Contain("header-footer-content");
        content.Header.Html.Should().NotContain("font-size:0");
    }
}

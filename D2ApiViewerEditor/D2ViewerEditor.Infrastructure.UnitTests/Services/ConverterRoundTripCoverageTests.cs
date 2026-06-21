using System.Net;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Szeroki round-trip HTML → DOCX → HTML na bogatej treści (nagłówki, formatowanie znakowe,
/// listy, link, tabela, akapit z wyrównaniem/kolorem/rozmiarem). Sprawdza, że treść przeżywa
/// obie konwersje — przy okazji uruchamia dużą część gałęzi obu konwerterów.
/// </summary>
[TestFixture]
public class ConverterRoundTripCoverageTests
{
    private static string RoundTrip(string html)
    {
        var docx = new HtmlToDocxConverter().Convert(html);
        using var ms = new MemoryStream(docx);
        // Decode entities (table cells encode Polish diacritics, e.g. ó → &#243;) so text asserts are robust.
        return WebUtility.HtmlDecode(new DocxToHtmlConverter().Convert(ms).Html);
    }

    [Test]
    public void RoundTrip_RichDocument_PreservesTextContent()
    {
        const string html = """
            <h1>Tytuł dokumentu</h1>
            <h2>Podtytuł</h2>
            <p style="text-align:center;color:#ff0000;font-size:14pt">
              <b>Pogrubiony</b> <i>kursywa</i> <u>podkreślony</u> tekst akapitu.
            </p>
            <ul><li>Punkt pierwszy</li><li>Punkt drugi</li></ul>
            <ol><li>Krok jeden</li><li>Krok dwa</li></ol>
            <p><a href="https://example.com/link">Odnośnik zewnętrzny</a></p>
            <table>
              <tr><td>Komórka A1</td><td>Komórka B1</td></tr>
              <tr><td>Komórka A2</td><td>Komórka B2</td></tr>
            </table>
            """;

        var result = RoundTrip(html);

        result.Should().NotBeNullOrWhiteSpace();
        foreach (var text in new[]
        {
            "Tytuł dokumentu", "Podtytuł", "Pogrubiony", "kursywa", "podkreślony",
            "Punkt pierwszy", "Punkt drugi", "Krok jeden", "Krok dwa",
            "Odnośnik zewnętrzny", "Komórka A1", "Komórka B2"
        })
        {
            result.Should().Contain(text);
        }
    }

    [Test]
    public void RoundTrip_PlainParagraphs_AreStable()
    {
        var first = RoundTrip("<p>Pierwszy akapit.</p><p>Drugi akapit.</p>");
        var second = RoundTrip(first);

        first.Should().Contain("Pierwszy akapit.").And.Contain("Drugi akapit.");
        second.Should().Contain("Pierwszy akapit.").And.Contain("Drugi akapit.");
    }

    [Test]
    public void HtmlToDocx_EmptyHtml_ProducesValidDocxThatReadsBack()
    {
        var docx = new HtmlToDocxConverter().Convert("<p></p>");
        docx.Should().NotBeNullOrEmpty();

        using var ms = new MemoryStream(docx);
        var html = new DocxToHtmlConverter().Convert(ms).Html;
        html.Should().NotBeNull();
    }
}

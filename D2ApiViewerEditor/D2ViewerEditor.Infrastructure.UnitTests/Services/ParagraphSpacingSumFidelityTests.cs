using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// ADR-0053: Word SUMUJE odstęp „po" akapitu z odstępem „przed" następnego, a marginesy CSS
/// rodzeństwa kolapsują do max — reader emituje w:after jako padding-bottom (nie kolapsuje),
/// z wyjątkiem akapitów z tłem/obramowaniem (padding malowałby odstęp tłem → margin-bottom).
/// w:contextualSpacing znosi before/after między bezpośrednimi sąsiadami tego samego stylu.
/// </summary>
[TestFixture]
public class ParagraphSpacingSumFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream Docx(params OpenXmlElement[] bodyChildren)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var c in bodyChildren) body.Append(c);
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Paragraph P(string text, int? beforeTw = null, int? afterTw = null,
        bool contextual = false, string? shadingFill = null, string? styleId = null)
    {
        var pPr = new ParagraphProperties();
        if (styleId != null) pPr.Append(new ParagraphStyleId { Val = styleId });
        if (shadingFill != null)
            pPr.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = shadingFill });
        var sp = new SpacingBetweenLines();
        if (beforeTw.HasValue) sp.Before = beforeTw.Value.ToString();
        if (afterTw.HasValue) sp.After = afterTw.Value.ToString();
        pPr.Append(sp);
        if (contextual) pPr.Append(new ContextualSpacing());
        return new Paragraph(pPr, new Run(new Text(text)));
    }

    [Test]
    public void AfterAndBefore_EmitPaddingBottomPlusMarginTop_SoTheySumLikeWord()
    {
        // Word: gap = after(P1) + before(P2) = 12pt + 12pt = 24pt. Marginesy CSS skolapsowałyby
        // do 12pt; padding-bottom + margin-top renderują pełną sumę.
        using var ms = Docx(P("P1", afterTw: 240), P("P2", beforeTw: 240));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("padding-bottom:12pt;");
        html.Should().Contain("margin-top:12pt;");
        html.Should().NotContain("margin-bottom:12pt;");
    }

    [Test]
    public void ShadedParagraph_KeepsAfterAsMarginBottom_SoBackgroundDoesNotPaintTheGap()
    {
        // Cieniowanie w Wordzie NIE pokrywa odstępu „po" — padding malowałby go tłem.
        using var ms = Docx(P("Cieniowany", afterTw: 240, shadingFill: "D9D9D9"), P("Dalej"));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("background-color:#D9D9D9;");
        html.Should().Contain("margin-bottom:12pt;");
        html.Should().NotContain("padding-bottom:12pt;");
        // Jawne zero neutralizuje klasowy domyślny odstęp (--doc-par-margin jest paddingiem)
        // — bez niego akapit z tłem dostawałby odstęp PODWÓJNIE i zamalowany tłem.
        html.Should().Contain("padding-bottom:0;");
    }

    [Test]
    public void ShadedParagraph_RoundTripsAfterValue_DespiteZeroPaddingReset()
    {
        // margin-bottom:12pt + padding-bottom:0 (reset klasowego domyślnego) →
        // writer musi zapisać w:after=240, nie 0.
        using var ms = Docx(P("Cieniowany", afterTw: 240, shadingFill: "D9D9D9"), P("Dalej"));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var saved = new MemoryStream(savedBytes);
        using var reopened = WordprocessingDocument.Open(saved, false);
        var spacing = reopened.MainDocumentPart!.Document.Body!
            .Descendants<SpacingBetweenLines>().First(s => s.After?.Value != null);

        spacing.After!.Value.Should().Be("240");
    }

    [Test]
    public void ContextualSpacing_SameStyleNeighbours_ZeroesTheSpacingBetweenThem()
    {
        // Trzy akapity tego samego (domyślnego) stylu z contextualSpacing: środkowy traci
        // before i after, skrajne tylko stronę sąsiadującą z tym samym stylem.
        using var ms = Docx(
            P("A", beforeTw: 240, afterTw: 240, contextual: true),
            P("B", beforeTw: 240, afterTw: 240, contextual: true),
            P("C", beforeTw: 240, afterTw: 240, contextual: true));

        var html = _reader.Convert(ms).Html;
        var styles = System.Text.RegularExpressions.Regex.Matches(html, "<p style=\"([^\"]*)\"")
            .Select(m => m.Groups[1].Value).ToList();

        styles.Should().HaveCount(3);
        styles[0].Should().Contain("margin-top:12pt;").And.Contain("padding-bottom:0;");
        styles[1].Should().Contain("margin-top:0;").And.Contain("padding-bottom:0;");
        styles[2].Should().Contain("margin-top:0;").And.Contain("padding-bottom:12pt;");
    }

    [Test]
    public void ContextualSpacing_DifferentStyleNeighbour_KeepsSpacing()
    {
        // Sąsiad INNEGO stylu — Word nie znosi odstępów (styl jawny vs domyślny ≠ ten sam).
        using var ms = Docx(
            P("A", afterTw: 240, contextual: true, styleId: "Quote"),
            P("B", beforeTw: 240));

        var html = _reader.Convert(ms).Html;
        var first = System.Text.RegularExpressions.Regex.Match(html, "<p[^>]*style=\"([^\"]*)\"");

        first.Groups[1].Value.Should().Contain("padding-bottom:12pt;");
        first.Groups[1].Value.Should().NotContain("padding-bottom:0;");
    }

    [Test]
    public void Writer_MapsPaddingBottomBackToSpacingAfter_RoundTrip()
    {
        using var ms = Docx(P("Tekst", beforeTw: 240, afterTw: 120));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var saved = new MemoryStream(savedBytes);

        using var reopened = WordprocessingDocument.Open(saved, false);
        var spacing = reopened.MainDocumentPart!.Document.Body!
            .Descendants<SpacingBetweenLines>().First();

        spacing.Before!.Value.Should().Be("240");
        spacing.After!.Value.Should().Be("120");
    }
}

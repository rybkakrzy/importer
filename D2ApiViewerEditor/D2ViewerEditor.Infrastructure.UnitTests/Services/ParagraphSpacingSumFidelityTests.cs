using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// ADR-0107 (koryguje ADR-0053): Word — pomiar COM (Word 16, także tryby zgodności 11/12/14)
/// — liczy odstęp między akapitami jako MAX(after A, before B); SUMUJE wyłącznie z flagą
/// zgodności w:doNotUseHTMLParagraphAutoSpacing. Reader emituje więc w:after jako
/// margin-bottom (kolaps marginesów CSS = max Worda), a padding-bottom (nie kolapsuje = suma)
/// tylko dla dokumentów z flagą. Akapity z tłem/obramowaniem w modelu sumy wracają na margin
/// (padding malowałby odstęp tłem). w:contextualSpacing = markery prezentacyjne, wartości
/// inline zostają (round-trip bezstratny).
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

    private static MemoryStream Docx(bool sumCompat, params OpenXmlElement[] bodyChildren)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var c in bodyChildren) body.Append(c);
            mainPart.Document = new Document(body);
            if (sumCompat)
            {
                var settings = mainPart.AddNewPart<DocumentSettingsPart>();
                settings.Settings = new Settings(new Compatibility(new DoNotUseHTMLParagraphAutoSpacing()));
                settings.Settings.Save();
            }
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static MemoryStream Docx(params OpenXmlElement[] bodyChildren) => Docx(false, bodyChildren);

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

    private static List<string> ParagraphStyles(string html) =>
        System.Text.RegularExpressions.Regex.Matches(html, "<p style=\"([^\"]*)\"")
            .Select(m => m.Groups[1].Value).ToList();

    [Test]
    public void Default_AfterIsMarginBottom_SoAdjacentMarginsCollapseToMaxLikeWord()
    {
        // Word (pomiar): after(A)=12pt + before(B)=4pt → 12pt (przykład ze specyfikacji).
        // Marginesy CSS rodzeństwa kolapsują do max — dokładnie ten model.
        using var ms = Docx(P("P1", afterTw: 240), P("P2", beforeTw: 80));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("margin-bottom:12pt;");
        html.Should().Contain("margin-top:4pt;");
        html.Should().NotContain("padding-bottom:");
        html.Should().NotContain("data-para-spacing-sum");
    }

    [Test]
    public void CompatFlag_AfterIsPaddingBottom_SoSpacingsSum()
    {
        // w:doNotUseHTMLParagraphAutoSpacing: Word SUMUJE (pomiar: 12+4 → 16pt) — padding
        // nie kolapsuje; kontener niesie flagę dla GUI (klasa para-spacing-sum) i writera.
        using var ms = Docx(sumCompat: true, P("P1", afterTw: 240), P("P2", beforeTw: 80));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("padding-bottom:12pt;");
        html.Should().Contain("margin-top:4pt;");
        html.Should().NotContain("margin-bottom:12pt;");
        html.Should().Contain("data-para-spacing-sum=\"1\"");
    }

    [Test]
    public void CompatFlag_RoundTripsToSettingsXml()
    {
        using var ms = Docx(sumCompat: true, P("P1", afterTw: 240), P("P2", beforeTw: 80));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(savedBytes), false);

        reopened.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .GetFirstChild<Compatibility>()!.Elements<DoNotUseHTMLParagraphAutoSpacing>()
            .Should().ContainSingle("bez flagi Word po zapisie przeszedłby na max");
        var spacing = reopened.MainDocumentPart.Document.Body!
            .Descendants<SpacingBetweenLines>().First(s => s.After?.Value != null);
        spacing.After!.Value.Should().Be("240");
    }

    [Test]
    public void DefaultDocument_DoesNotEmitCompatFlag()
    {
        using var ms = Docx(P("P1", afterTw: 240), P("P2", beforeTw: 80));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(savedBytes), false);

        var compat = reopened.MainDocumentPart!.DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>();
        (compat?.Elements<DoNotUseHTMLParagraphAutoSpacing>().Any() ?? false).Should().BeFalse();
    }

    [Test]
    public void SumModel_ShadedParagraph_KeepsAfterAsMarginBottom_SoBackgroundDoesNotPaintTheGap()
    {
        // Cieniowanie w Wordzie NIE pokrywa odstępu „po" — padding malowałby go tłem
        // (wyjątek ADR-0053 potrzebny tylko w modelu sumy).
        using var ms = Docx(sumCompat: true, P("Cieniowany", afterTw: 240, shadingFill: "D9D9D9"), P("Dalej"));

        var html = _reader.Convert(ms).Html;

        html.Should().Contain("background-color:#D9D9D9;");
        html.Should().Contain("margin-bottom:12pt;");
        html.Should().NotContain("padding-bottom:12pt;");
        // Jawne zero neutralizuje klasowy domyślny odstęp (w modelu sumy jest paddingiem).
        html.Should().Contain("padding-bottom:0;");
    }

    [Test]
    public void SumModel_ShadedParagraph_RoundTripsAfterValue_DespiteZeroPaddingReset()
    {
        // margin-bottom:12pt + padding-bottom:0 → writer musi zapisać w:after=240, nie 0.
        using var ms = Docx(sumCompat: true, P("Cieniowany", afterTw: 240, shadingFill: "D9D9D9"), P("Dalej"));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(savedBytes), false);
        var spacing = reopened.MainDocumentPart!.Document.Body!
            .Descendants<SpacingBetweenLines>().First(s => s.After?.Value != null);

        spacing.After!.Value.Should().Be("240");
    }

    [Test]
    public void ContextualSpacing_SameStyleNeighbours_MarksSuppressedSidesAndKeepsValues()
    {
        // Trzy akapity tego samego (domyślnego) stylu z contextualSpacing: środkowy ma obie
        // strony zniesione, skrajne tylko stronę sąsiadującą. Wartości inline ZOSTAJĄ
        // (dialog pokazuje zadeklarowane, writer nie utrwala zer — R-38 c).
        using var ms = Docx(
            P("A", beforeTw: 240, afterTw: 240, contextual: true),
            P("B", beforeTw: 240, afterTw: 240, contextual: true),
            P("C", beforeTw: 240, afterTw: 240, contextual: true));

        var styles = ParagraphStyles(_reader.Convert(ms).Html);

        styles.Should().HaveCount(3);
        styles[0].Should().Contain("margin-top:12pt;").And.Contain("margin-bottom:12pt;")
            .And.Contain("--w-ctx-next:1;").And.NotContain("--w-ctx-prev");
        styles[1].Should().Contain("--w-ctx-prev:1;").And.Contain("--w-ctx-next:1;");
        styles[2].Should().Contain("--w-ctx-prev:1;").And.NotContain("--w-ctx-next");
    }

    [Test]
    public void ContextualSpacing_DifferentStyleNeighbour_KeepsSpacing()
    {
        // Sąsiad INNEGO stylu — Word nie znosi odstępów (styl jawny vs domyślny ≠ ten sam).
        using var ms = Docx(
            P("A", afterTw: 240, contextual: true, styleId: "Quote"),
            P("B", beforeTw: 240));

        var styles = ParagraphStyles(_reader.Convert(ms).Html);

        styles[0].Should().Contain("margin-bottom:12pt;").And.NotContain("--w-ctx-");
        styles[1].Should().NotContain("--w-ctx-");
    }

    [Test]
    public void ContextualSpacing_RoundTripsDeclaredValuesAndFlag()
    {
        using var ms = Docx(
            P("A", beforeTw: 240, afterTw: 200, contextual: true),
            P("B", beforeTw: 120, afterTw: 240, contextual: true));
        var content = _reader.Convert(ms);

        var savedBytes = _writer.Convert(content.Html);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(savedBytes), false);
        var paragraphs = reopened.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();

        var a = paragraphs[0].ParagraphProperties!;
        a.GetFirstChild<SpacingBetweenLines>()!.After!.Value.Should().Be("200", "zniesiony odstęp nie może wracać jako 0");
        a.GetFirstChild<ContextualSpacing>().Should().NotBeNull();
        var b = paragraphs[1].ParagraphProperties!;
        b.GetFirstChild<SpacingBetweenLines>()!.Before!.Value.Should().Be("120");
        b.GetFirstChild<ContextualSpacing>().Should().NotBeNull();
    }

    [Test]
    public void Writer_MapsBothAfterCarriersBackToSpacingAfter()
    {
        using var ms = Docx(P("Tekst", beforeTw: 240, afterTw: 120));
        var content = _reader.Convert(ms);
        var savedBytes = _writer.Convert(content.Html);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(savedBytes), false);
        var spacing = reopened.MainDocumentPart!.Document.Body!.Descendants<SpacingBetweenLines>().First();
        spacing.Before!.Value.Should().Be("240");
        spacing.After!.Value.Should().Be("120");

        // Treść autorstwa GUI w modelu sumy (padding-bottom) nadal jest czytana.
        var legacyBytes = _writer.Convert("<p style=\"margin-top:12pt;padding-bottom:6pt;\">Tekst</p>");
        using var legacy = WordprocessingDocument.Open(new MemoryStream(legacyBytes), false);
        legacy.MainDocumentPart!.Document.Body!.Descendants<SpacingBetweenLines>().First()
            .After!.Value.Should().Be("120");
    }
}

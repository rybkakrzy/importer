using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Wierność list wielopoziomowych (numPr/numId/ilvl/abstractNum/lvl/numFmt/lvlText/start/
/// startOverride) w round-tripie DOCX↔HTML.
///
/// Semantyka Worda, którą te testy utrwalają:
/// - licznik numeracji żyje per ABSTRAKT (różne w:num na ten sam abstrakt kontynuują numerację,
///   chyba że instancja ma startOverride — wtedy restart przy pierwszym użyciu),
/// - lista przerwana zwykłym akapitem kontynuuje numerację (ten sam numId → `start` na drugim
///   fragmencie ol + wspólny data-num-id scalający fragmenty przy zapisie),
/// - niezależne listy o identycznym wyglądzie NIE są sklejane (tożsamość = numId, nie format),
/// - poziomy głębsze restartują po powrocie na poziom płytszy,
/// - format poziomu (numFmt/lvlText/start/font punktatora) wraca do DOCX z data-*.
/// </summary>
[TestFixture]
public class ListNumberingFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    // ── Helpery budowy DOCX in-memory ────────────────────────────────────────────

    private static Level Lvl(int index, NumberFormatValues fmt, string lvlText, int start = 1)
    {
        var level = new Level { LevelIndex = index };
        level.Append(new StartNumberingValue { Val = start });
        level.Append(new NumberingFormat { Val = fmt });
        level.Append(new LevelText { Val = lvlText });
        level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
        return level;
    }

    private static AbstractNum Abstract(int id, params Level[] levels)
    {
        var abs = new AbstractNum { AbstractNumberId = id };
        abs.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
        foreach (var l in levels) abs.Append(l);
        return abs;
    }

    private static NumberingInstance Num(int numId, int absId, params (int lvl, int startOverride)[] overrides)
    {
        var num = new NumberingInstance { NumberID = numId };
        num.Append(new AbstractNumId { Val = absId });
        foreach (var (lvl, so) in overrides)
        {
            num.Append(new LevelOverride(
                new StartOverrideNumberingValue { Val = so })
            { LevelIndex = lvl });
        }
        return num;
    }

    private static Paragraph ListItem(string text, int numId, int ilvl) =>
        new(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = ilvl },
                    new NumberingId { Val = numId })),
            new Run(new Text(text)));

    private static Paragraph Plain(string text) => new(new Run(new Text(text)));

    private static MemoryStream BuildDocx(
        IEnumerable<AbstractNum> abstracts,
        IEnumerable<NumberingInstance> nums,
        IEnumerable<OpenXmlElement> bodyElements)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
            foreach (var a in abstracts) numberingPart.Numbering.Append(a);
            foreach (var n in nums) numberingPart.Numbering.Append(n);
            numberingPart.Numbering.Save();

            var body = mainPart.Document.Body!;
            foreach (var el in bodyElements) body.Append(el);
            body.Append(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    /// <summary>Wszystkie otwierające tagi ol/ul (z atrybutami) w kolejności wystąpienia.</summary>
    private static List<string> ListOpenTags(string html) =>
        Regex.Matches(html, "<(ol|ul)[^>]*>").Select(m => m.Value).ToList();

    // ── Reader: DOCX → HTML ─────────────────────────────────────────────────────

    [Test]
    public void NestedBullets_MainNumberingContinues_AfterNesting()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."), Lvl(1, NumberFormatValues.Bullet, "•"))],
            [Num(1, 1)],
            [
                ListItem("Pierwszy", 1, 0),
                ListItem("Podpunkt", 1, 1),
                ListItem("Kolejny podpunkt", 1, 1),
                ListItem("Drugi", 1, 0),
                ListItem("Trzeci", 1, 0),
            ]);

        var html = _reader.Convert(docx).Html;

        var tags = ListOpenTags(html);
        // Jedna lista główna (ol) z zagnieżdżoną ul — bez rozbicia głównego poziomu.
        tags.Count(t => t.StartsWith("<ol")).Should().Be(1);
        tags.Count(t => t.StartsWith("<ul")).Should().Be(1);
        // Główny <ol> bez `start` — numeracja 1..3 ciągła mimo zagnieżdżenia.
        tags.First(t => t.StartsWith("<ol")).Should().NotContain("start=");
        tags.First(t => t.StartsWith("<ol")).Should().Contain("data-num-id=\"1\"").And.Contain("data-ilvl=\"0\"");
        tags.First(t => t.StartsWith("<ul")).Should().Contain("data-ilvl=\"1\"");
        Regex.Matches(html, "<li").Count.Should().Be(5);
    }

    [Test]
    public void ListInterruptedByParagraph_ContinuesNumbering()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."))],
            [Num(1, 1)],
            [
                ListItem("Jeden", 1, 0),
                ListItem("Dwa", 1, 0),
                Plain("Zwykły akapit w środku"),
                ListItem("Trzy", 1, 0),
                ListItem("Cztery", 1, 0),
            ]);

        var html = _reader.Convert(docx).Html;

        var ols = ListOpenTags(html).Where(t => t.StartsWith("<ol")).ToList();
        ols.Should().HaveCount(2);
        ols[0].Should().NotContain("start=");
        // Kontynuacja: drugi fragment zaczyna od 3 i niesie TĘ SAMĄ tożsamość listy.
        ols[1].Should().Contain("start=\"3\"");
        ols[0].Should().Contain("data-num-id=\"1\"");
        ols[1].Should().Contain("data-num-id=\"1\"");
    }

    [Test]
    public void IndependentLists_SameFormat_AreNotMerged()
    {
        using var docx = BuildDocx(
            [
                Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1.")),
                Abstract(2, Lvl(0, NumberFormatValues.Decimal, "%1.")),
            ],
            [Num(1, 1), Num(2, 2)],
            [
                ListItem("A1", 1, 0),
                ListItem("A2", 1, 0),
                ListItem("B1", 2, 0),
                ListItem("B2", 2, 0),
            ]);

        var html = _reader.Convert(docx).Html;

        var ols = ListOpenTags(html).Where(t => t.StartsWith("<ol")).ToList();
        // Identyczny wygląd ≠ ta sama lista: dwa numId/abstrakty → dwa osobne <ol>.
        ols.Should().HaveCount(2);
        ols[0].Should().Contain("data-num-id=\"1\"");
        ols[1].Should().Contain("data-num-id=\"2\"");
        // Druga lista zaczyna od 1 (niezależny licznik) — brak `start`.
        ols[1].Should().NotContain("start=");
    }

    [Test]
    public void StartValueOtherThanOne_IsHonored()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1.", start: 5))],
            [Num(1, 1)],
            [ListItem("Piąty", 1, 0), ListItem("Szósty", 1, 0)]);

        var html = _reader.Convert(docx).Html;

        var ol = ListOpenTags(html).First(t => t.StartsWith("<ol"));
        ol.Should().Contain("start=\"5\"");
        ol.Should().Contain("data-start=\"5\""); // definicja w:start round-tripuje osobno od prezentacji
    }

    [Test]
    public void SharedAbstract_DifferentNums_ContinueNumbering()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."))],
            [Num(1, 1), Num(2, 1)],
            [
                ListItem("Jeden", 1, 0),
                ListItem("Dwa", 1, 0),
                Plain("Przerwa"),
                ListItem("Trzy", 2, 0),
            ]);

        var html = _reader.Convert(docx).Html;

        var ols = ListOpenTags(html).Where(t => t.StartsWith("<ol")).ToList();
        ols.Should().HaveCount(2);
        // Wspólny abstrakt bez startOverride = „Kontynuuj numerację" w Wordzie.
        ols[1].Should().Contain("start=\"3\"");
    }

    [Test]
    public void StartOverride_RestartsNumbering_DespiteSharedAbstract()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."))],
            [Num(1, 1), Num(2, 1, (0, 1))],
            [
                ListItem("Jeden", 1, 0),
                ListItem("Dwa", 1, 0),
                Plain("Przerwa"),
                ListItem("Restart-jeden", 2, 0),
            ]);

        var html = _reader.Convert(docx).Html;

        var ols = ListOpenTags(html).Where(t => t.StartsWith("<ol")).ToList();
        ols.Should().HaveCount(2);
        // startOverride=1 = „Rozpocznij od nowa" — mimo wspólnego abstraktu.
        ols[1].Should().NotContain("start=");
    }

    [Test]
    public void DeeperLevel_RestartsAfterReturnToMainLevel()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."), Lvl(1, NumberFormatValues.LowerLetter, "%2."))],
            [Num(1, 1)],
            [
                ListItem("Jeden", 1, 0),
                ListItem("a", 1, 1),
                ListItem("b", 1, 1),
                ListItem("Dwa", 1, 0),
                ListItem("a-znowu", 1, 1),
            ]);

        var html = _reader.Convert(docx).Html;

        // Druga zagnieżdżona lista (po powrocie na poziom główny) restartuje: bez `start`.
        var nestedOls = ListOpenTags(html)
            .Where(t => t.StartsWith("<ol") && t.Contains("data-ilvl=\"1\""))
            .ToList();
        nestedOls.Should().HaveCount(2);
        nestedOls[1].Should().NotContain("start=");
    }

    [Test]
    public void MixedFormatsPerLevel_EmitCorrectTagsAndDataAttributes()
    {
        using var docx = BuildDocx(
            [Abstract(1,
                Lvl(0, NumberFormatValues.UpperRoman, "%1)"),
                Lvl(1, NumberFormatValues.Bullet, "•"),
                Lvl(2, NumberFormatValues.LowerLetter, "%3."))],
            [Num(1, 1)],
            [
                ListItem("Rzymski", 1, 0),
                ListItem("Punktor", 1, 1),
                ListItem("Litera", 1, 2),
            ]);

        var html = _reader.Convert(docx).Html;

        var tags = ListOpenTags(html);
        var outer = tags.First(t => t.StartsWith("<ol") && t.Contains("data-ilvl=\"0\""));
        outer.Should().Contain("data-num-fmt=\"upperRoman\"");
        outer.Should().Contain("data-lvl-text=\"%1)\"");
        outer.Should().Contain("upper-roman");
        tags.Should().Contain(t => t.StartsWith("<ul") && t.Contains("data-num-fmt=\"bullet\""));
        var innerOl = tags.First(t => t.StartsWith("<ol") && t.Contains("data-ilvl=\"2\""));
        innerOl.Should().Contain("data-num-fmt=\"lowerLetter\"");
        innerOl.Should().Contain("lower-alpha");
    }

    // ── Writer: HTML → DOCX ─────────────────────────────────────────────────────

    private static (List<(int numId, int ilvl)> items, Numbering numbering) ReadListParagraphs(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var items = doc.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Select(p => p.ParagraphProperties?.NumberingProperties)
            .Where(n => n?.NumberingId?.Val?.Value is > 0)
            .Select(n => (n!.NumberingId!.Val!.Value, n.NumberingLevelReference?.Val?.Value ?? 0))
            .ToList();
        var numbering = (Numbering)doc.MainDocumentPart.NumberingDefinitionsPart!.Numbering.CloneNode(true);
        return (items, numbering);
    }

    [Test]
    public void Writer_SameDataNumId_SharesOneNumberingInstance()
    {
        var html =
            "<ol data-num-id=\"7\" data-ilvl=\"0\" data-num-fmt=\"decimal\"><li>a</li><li>b</li></ol>" +
            "<p>przerwa</p>" +
            "<ol start=\"3\" data-num-id=\"7\" data-ilvl=\"0\" data-num-fmt=\"decimal\"><li>c</li></ol>";

        var (items, numbering) = ReadListParagraphs(_writer.Convert(html));

        items.Should().HaveCount(3);
        items.Select(i => i.numId).Distinct().Should().HaveCount(1,
            "fragmenty tej samej listy logicznej muszą współdzielić numId — Word kontynuuje numerację");
        numbering.Elements<NumberingInstance>().Should().HaveCount(1);
    }

    [Test]
    public void Writer_DistinctDataNumIds_CreateSeparateInstances()
    {
        var html =
            "<ol data-num-id=\"7\" data-num-fmt=\"decimal\"><li>a</li></ol>" +
            "<ol data-num-id=\"8\" data-num-fmt=\"decimal\"><li>b</li></ol>";

        var (items, numbering) = ReadListParagraphs(_writer.Convert(html));

        items.Should().HaveCount(2);
        items.Select(i => i.numId).Distinct().Should().HaveCount(2,
            "niezależne listy nie mogą zostać sklejone w jedną numerację");
        numbering.Elements<NumberingInstance>().Should().HaveCount(2);
    }

    [Test]
    public void Writer_CustomFormat_RoundTripsToAbstractNum()
    {
        var html =
            "<ol data-num-id=\"7\" data-ilvl=\"0\" data-num-fmt=\"upperRoman\" data-lvl-text=\"%1)\" data-start=\"3\">" +
            "<li>rzymski</li></ol>";

        var (_, numbering) = ReadListParagraphs(_writer.Convert(html));

        var lvl0 = numbering.Elements<AbstractNum>().Single().Elements<Level>()
            .First(l => l.LevelIndex?.Value == 0);
        lvl0.NumberingFormat!.Val!.Value.Should().Be(NumberFormatValues.UpperRoman);
        lvl0.LevelText!.Val!.Value.Should().Be("%1)");
        lvl0.StartNumberingValue!.Val!.Value.Should().Be(3);
    }

    // ── Pełny round-trip: DOCX → HTML → DOCX → HTML ─────────────────────────────

    [Test]
    public void FullRoundTrip_PreservesListStructureAndContinuation()
    {
        using var docx = BuildDocx(
            [Abstract(1, Lvl(0, NumberFormatValues.Decimal, "%1."), Lvl(1, NumberFormatValues.Bullet, "•"))],
            [Num(1, 1)],
            [
                ListItem("Jeden", 1, 0),
                ListItem("Podpunkt", 1, 1),
                ListItem("Dwa", 1, 0),
                Plain("Przerwa"),
                ListItem("Trzy", 1, 0),
            ]);

        var html1 = _reader.Convert(docx).Html;
        var regenerated = _writer.Convert(html1);
        using var second = new MemoryStream(regenerated);
        var html2 = _reader.Convert(second).Html;

        // Struktura po pełnym round-tripie: główna numeracja ciągła (drugi fragment start=3),
        // zagnieżdżony punktor zachowany, fragmenty wciąż jedną listą logiczną (ten sam data-num-id).
        var ols = ListOpenTags(html2).Where(t => t.StartsWith("<ol")).ToList();
        ols.Should().HaveCount(2);
        ols[1].Should().Contain("start=\"3\"");
        ListOpenTags(html2).Should().Contain(t => t.StartsWith("<ul"));

        string NumIdOf(string tag) => Regex.Match(tag, "data-num-id=\"(\\d+)\"").Groups[1].Value;
        NumIdOf(ols[0]).Should().NotBeEmpty();
        NumIdOf(ols[0]).Should().Be(NumIdOf(ols[1]),
            "po zapisie i ponownym otwarciu fragmenty muszą nadal należeć do jednej listy");

        // Format głównego poziomu przeżył round-trip (nie zdegradował się do drabinki domyślnej).
        ols[0].Should().Contain("data-num-fmt=\"decimal\"");
    }
}

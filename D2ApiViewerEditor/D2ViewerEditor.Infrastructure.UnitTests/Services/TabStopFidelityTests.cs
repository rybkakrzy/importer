using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Tab-stopy per akapit (układ „lewa⇥środek⇥prawa" w nagłówku/stopce).
///
/// Poprzednio: pozycje w:tabs nie były czytane (flex przybliżał center/right jako 50%/100%),
/// zapis emitował tab-stopy tylko w STYLACH Header/Footer (sztywne 4536/9072), a znak taba
/// w tekście trafiał do w:t (Word go nie renderuje). Teraz: reader emituje
/// data-tab-stops="pos:align[:leader]" (round-trip per akapit) i w nagłówku/stopce renderuje
/// segmenty POZYCYJNIE (span.docx-tab-seg na pozycji stopu); writer odtwarza w:tabs z
/// atrybutu i w:tab z segmentów/znaków \t.
/// </summary>
[TestFixture]
public class TabStopFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    /// <summary>Dokument z nagłówkiem „Lewa ⇥ Środek ⇥ Prawa" (center 4536, right 9072 tw).</summary>
    private static MemoryStream BuildDocxWithTabbedHeader(int centerPos = 4536, int rightPos = 9072)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(
                new ParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Center, Position = centerPos },
                    new TabStop { Val = TabStopValues.Right, Position = rightPos })),
                new Run(new Text("Lewa")),
                new Run(new TabChar()),
                new Run(new Text("Środek")),
                new Run(new TabChar()),
                new Run(new Text("Prawa"))));
            headerPart.Header.Save();

            body.Append(new Paragraph(new Run(new Text("Treść"))));
            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_HeaderWithTabStops_EmitsDataTabStopsAndPositionedSegments()
    {
        using var stream = BuildDocxWithTabbedHeader();

        var content = _reader.Convert(stream);

        content.Header.Should().NotBeNull();
        var html = content.Header!.Html;
        html.Should().Contain("data-tab-stops=\"4536:center;9072:right\"");
        html.Should().Contain("docx-tab-seg");
        // center = wyśrodkowany NA pozycji, right = kończy się NA pozycji.
        html.Should().Contain("data-tab-align=\"center\"");
        html.Should().Contain("translateX(-50%)");
        html.Should().Contain("data-tab-align=\"right\"");
        html.Should().Contain("translateX(-100%)");
        // Pozycja center: 4536 tw = 302 px (96 DPI).
        html.Should().Contain("left:302px");
        html.Should().Contain("Środek");
    }

    [Test]
    public void Read_BodyParagraphWithTabStops_RendersPositionedSegmentsAtRealStops()
    {
        // Body paragraphs must honour the real tab-stop positions (right stop = following segment's
        // END aligned to the stop), not a flex row that spreads segments evenly and ignores where
        // the stops sit. 9000 tw = 600 px @96 DPI.
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Right, Position = 9000 })),
                new Run(new Text("Podpis")),
                new Run(new TabChar()),
                new Run(new Text("Data")))));
            mainPart.Document.Save();
        }
        ms.Position = 0;

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("data-tab-stops=\"9000:right\"");
        content.Html.Should().NotContain("display:flex", "body używa teraz pozycjonowania na realnych stopach");
        content.Html.Should().Contain("docx-tab-seg");
        content.Html.Should().Contain("data-tab-align=\"right\"");
        content.Html.Should().Contain("left:600px");
        content.Html.Should().Contain("translateX(-100%)");
    }

    [Test]
    public void Write_DataTabStops_RecreatesTabsInParagraphProperties()
    {
        var html = "<p data-tab-stops=\"4536:center;9072:right:dot\">Lewa\tŚrodek\tPrawa</p>";

        var bytes = _writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var para = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var tabs = para.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().ToList();

        tabs.Should().HaveCount(2);
        ((int)tabs[0].Position!.Value).Should().Be(4536);
        tabs[0].Val!.Value.Should().Be(TabStopValues.Center);
        ((int)tabs[1].Position!.Value).Should().Be(9072);
        tabs[1].Val!.Value.Should().Be(TabStopValues.Right);
        tabs[1].Leader!.Value.Should().Be(TabStopLeaderCharValues.Dot);
    }

    [Test]
    public void Write_LiteralTabInText_BecomesTabCharElement_NotText()
    {
        var bytes = _writer.Convert("<p>Lewa\tPrawa</p>");

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var para = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        para.Descendants<TabChar>().Should().HaveCount(1, "Word nie renderuje \\t w w:t — tab musi być elementem w:tab");
        para.InnerText.Should().NotContain("\t");
    }

    [Test]
    public void RoundTrip_TabbedHeader_PreservesStopsAndTabs()
    {
        using var stream = BuildDocxWithTabbedHeader();
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert("<p>Treść</p>",
            header: content.Header, footer: content.Footer,
            margins: content.Margins, pageSize: content.PageSize);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.First();
        var para = headerPart.Header!.Elements<Paragraph>().First();

        // Pozycje/wyrównania wracają per akapit (nie tylko sztywne style Header/Footer).
        var tabs = para.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().ToList();
        tabs.Should().HaveCount(2);
        ((int)tabs[0].Position!.Value).Should().Be(4536);
        ((int)tabs[1].Position!.Value).Should().Be(9072);

        // Dwa tabulatory jako w:tab; kolejność treści zachowana.
        para.Descendants<TabChar>().Should().HaveCount(2);
        para.InnerText.Should().Contain("Lewa").And.Contain("Środek").And.Contain("Prawa");
    }

    [Test]
    public void RoundTrip_TabbedHeader_SecondPass_IsStable()
    {
        // DOCX→HTML→DOCX→HTML: pozycyjne segmenty nie mogą się mnożyć ani gubić.
        using var stream = BuildDocxWithTabbedHeader();
        var first = _reader.Convert(stream);
        var bytes = _writer.Convert("<p>Treść</p>", header: first.Header);
        var second = _reader.Convert(new MemoryStream(bytes));

        second.Header!.Html.Should().Contain("data-tab-stops=\"4536:center;9072:right\"");
        var segCount = System.Text.RegularExpressions.Regex.Matches(second.Header.Html, "docx-tab-seg").Count;
        segCount.Should().Be(2);
    }

    /// <summary>Dokument z jednym akapitem body o zadanych pPr (opcjonalnie ze stylem akapitowym).</summary>
    private static MemoryStream BuildDocxWithBodyParagraph(ParagraphProperties paraProps, Style? paragraphStyle = null)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            if (paragraphStyle != null)
            {
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(paragraphStyle);
                stylesPart.Styles.Save();
            }
            mainPart.Document = new Document(new Body(new Paragraph(
                paraProps,
                new Run(new Text("Lewa")),
                new Run(new TabChar()),
                new Run(new Text("Prawa")))));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Style ParagraphStyleWithCenterTab(int position = 4536) => new(
        new StyleParagraphProperties(new Tabs(
            new TabStop { Val = TabStopValues.Center, Position = position })))
    {
        Type = StyleValues.Paragraph,
        StyleId = "TabbedStyle",
        StyleName = new StyleName { Val = "TabbedStyle" }
    };

    [Test]
    public void Read_TabLeader_IsSerializedInDataTabStops()
    {
        using var ms = BuildDocxWithBodyParagraph(new ParagraphProperties(new Tabs(
            new TabStop { Val = TabStopValues.Right, Position = 9072, Leader = TabStopLeaderCharValues.Dot })));

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("data-tab-stops=\"9072:right:dot\"",
            "leader musi przetrwać round-trip w trzecim segmencie atrybutu");
    }

    [Test]
    public void Read_ClearTab_RemovesStopInheritedFromParagraphStyle()
    {
        // Styl definiuje center 4536; akapit czyści go (w:val=clear) i dodaje right 9000 — jak Word.
        using var ms = BuildDocxWithBodyParagraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "TabbedStyle" },
                new Tabs(
                    new TabStop { Val = TabStopValues.Clear, Position = 4536 },
                    new TabStop { Val = TabStopValues.Right, Position = 9000 })),
            ParagraphStyleWithCenterTab());

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("data-tab-stops=\"9000:right\"");
        content.Html.Should().NotContain("4536:center");
    }

    [Test]
    public void Read_DirectTab_OverridesStyleStopAtSamePosition()
    {
        // Ta sama pozycja co w stylu, inne wyrównanie → wygrywa definicja z akapitu.
        using var ms = BuildDocxWithBodyParagraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "TabbedStyle" },
                new Tabs(new TabStop { Val = TabStopValues.Right, Position = 4536 })),
            ParagraphStyleWithCenterTab());

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("data-tab-stops=\"4536:right\"");
        content.Html.Should().NotContain("4536:center");
    }

    [Test]
    public void Read_BarTab_IsSkipped_RealStopsRemain()
    {
        // Bar-tab rysuje pionową linię, nie pozycjonuje tekstu — nie wchodzi do data-tab-stops.
        using var ms = BuildDocxWithBodyParagraph(new ParagraphProperties(new Tabs(
            new TabStop { Val = TabStopValues.Bar, Position = 3000 },
            new TabStop { Val = TabStopValues.Right, Position = 9000 })));

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("data-tab-stops=\"9000:right\"");
        content.Html.Should().NotContain("3000:");
    }

    /// <summary>Dokument z nagłówkiem/stopką o JEDNYM własnym stopie right 9360 (≠ 4536/9072).</summary>
    private static MemoryStream BuildDocxWithSingleRightStopBands()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(
                new ParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Right, Position = 9360 })),
                new Run(new Text("Logo")),
                new Run(new TabChar()),
                new Run(new Text("Katowice, data"))));
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                new Paragraph(new Run(new Text("Qutas  — linia bez tabulatora"))),
                new Paragraph(
                    new ParagraphProperties(new Tabs(
                        new TabStop { Val = TabStopValues.Right, Position = 9360 })),
                    new Run(new Text("Miejscowość: Katowice")),
                    new Run(new TabChar()),
                    new Run(new Text("Data"))));
            footerPart.Footer.Save();

            body.Append(new Paragraph(new Run(new Text("Treść"))));
            body.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void RoundTrip_BandParagraphWithOwnStop_StyleMustNotInjectWordDefaultStops()
    {
        // Regresja bug 13261178 („po Zapisz rozwala się formatowanie stopki"): writer stemplował
        // pasma stylem Header/Footer, którego definicja niosła klasyczne stopy Worda
        // (4536:center/9072:right). Stopy stylu SUMUJĄ się z bezpośrednimi w:tabs (ECMA-376),
        // więc po zapisie tabulator „Miejscowość ⇥ Data" łapał 4536:center zamiast własnego
        // 9360:right — w podglądzie i w realnym Wordzie.
        using var stream = BuildDocxWithSingleRightStopBands();
        var first = _reader.Convert(stream);

        var bytes = _writer.Convert("<p>Treść</p>", header: first.Header, footer: first.Footer);
        var second = new DocxToHtmlConverter().Convert(new MemoryStream(bytes));

        second.Footer!.Html.Should().Contain("data-tab-stops=\"9360:right\"",
            "efektywne stopy akapitu stopki to wyłącznie jego własny right 9360");
        second.Footer.Html.Should().NotContain("4536").And.NotContain("9072");
        second.Footer.Html.Should().Contain("data-tab-align=\"right\"",
            "segment po tabulatorze musi zostać na prawym stopie, nie przeskoczyć na center");
        second.Footer.Html.Should().NotContain("data-tab-align=\"center\"");
        second.Header!.Html.Should().Contain("data-tab-stops=\"9360:right\"");
        second.Header.Html.Should().NotContain("4536").And.NotContain("9072");
    }

    [Test]
    public void RoundTrip_BandParagraphs_ThirdPass_NoFlexAndStable()
    {
        // Progresja tej samej regresji: w 3. przebiegu akapity BEZ znaku tabulatora, którym
        // poprzedni zapis zmaterializował stopy center/right w pPr, dostawały display:flex
        // (fallback aligned-tab) i układ pasm rozjeżdżał się z każdym kolejnym zapisem.
        using var stream = BuildDocxWithSingleRightStopBands();
        var first = _reader.Convert(stream);
        var bytes2 = _writer.Convert("<p>Treść</p>", header: first.Header, footer: first.Footer);
        var second = new DocxToHtmlConverter().Convert(new MemoryStream(bytes2));
        var bytes3 = new HtmlToDocxConverter().Convert("<p>Treść</p>", header: second.Header, footer: second.Footer);
        var third = new DocxToHtmlConverter().Convert(new MemoryStream(bytes3));

        third.Footer!.Html.Should().NotContain("display:flex");
        third.Header!.Html.Should().NotContain("display:flex");
        third.Footer.Html.Should().Contain("data-tab-stops=\"9360:right\"");
        third.Footer.Html.Should().Contain("data-tab-align=\"right\"");
        // Geometria tabów ustabilizowana: kolejne zapisy nie zmieniają zestawu stopów.
        third.Footer.Html.Should().NotContain("4536").And.NotContain("9072");
    }

    [Test]
    public void PreservedPackage_OriginalFooterStyleWithWordStops_DoesNotHijackBandTab()
    {
        // Realna ścieżka /save: ConvertPreservingPackage podmienia styles.xml na ORYGINALNY,
        // w którym wordowy styl Footer MA stopy 4536:center/9072:right. Akapit stopki bez
        // własnego pStyle był stemplowany stylem Footer → stopy oryginalnego stylu sumowały
        // się z jego bezpośrednimi w:tabs i tabulator przeskakiwał na 4536:center.
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new Style(
                new StyleName { Val = "footer" },
                new StyleParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Center, Position = 4536 },
                    new TabStop { Val = TabStopValues.Right, Position = 9072 })))
            { Type = StyleValues.Paragraph, StyleId = "Footer" });
            stylesPart.Styles.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(
                new ParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Right, Position = 9360 })),
                new Run(new Text("Miejscowość: Katowice")),
                new Run(new TabChar()),
                new Run(new Text("Data"))));
            footerPart.Footer.Save();

            mainPart.Document.Body!.Append(new Paragraph(new Run(new Text("Treść"))));
            mainPart.Document.Body!.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        var original = ms.ToArray();

        var first = _reader.Convert(new MemoryStream(original));
        var saved = _writer.ConvertPreservingPackage(
            "<p>Treść</p>", new MemoryStream(original), footer: first.Footer);
        var second = new DocxToHtmlConverter().Convert(new MemoryStream(saved));

        second.Footer!.Html.Should().Contain("data-tab-stops=\"9360:right\"");
        second.Footer.Html.Should().NotContain("4536:center",
            "stopy wordowego stylu Footer z oryginalnego styles.xml nie mogą dokleić się do akapitu");
        second.Footer.Html.Should().Contain("data-tab-align=\"right\"");
        second.Footer.Html.Should().NotContain("data-tab-align=\"center\"");
    }

    [Test]
    public void Read_ParagraphDeclaringAlignedStops_WithoutTabChar_DoesNotUseFlexLayout()
    {
        // Akapit, który tylko DEKLARUJE stopy center/right (w:tabs w pPr), ale nie zawiera
        // znaku tabulatora, musi renderować się normalnie — flex zmieniał układ linii.
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Tabs(
                    new TabStop { Val = TabStopValues.Center, Position = 4536 },
                    new TabStop { Val = TabStopValues.Right, Position = 9072 })),
                new Run(new Text("Linia bez tabulatora")))));
            mainPart.Document.Save();
        }
        ms.Position = 0;

        var content = _reader.Convert(ms);

        content.Html.Should().NotContain("display:flex");
        content.Html.Should().Contain("data-tab-stops=\"4536:center;9072:right\"",
            "stopy nadal round-tripują — zmienia się tylko renderowanie");
    }
}

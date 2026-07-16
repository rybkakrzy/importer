using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// R-10: dokumenty wielosekcyjne (różne orientacje/rozmiary/marginesy per sekcja).
///
/// W OOXML sectPr będące dzieckiem w:body opisuje OSTATNIĄ sekcję; wcześniejsze sekcje
/// kończą się paragrafem z pPr/sectPr. Poprzednio reader brał body-level sectPr
/// ("FirstOrDefault"), więc dokument pionowy z poziomym aneksem otwierał się cały jako
/// poziomy, z nagłówkami ostatniej sekcji, a writer spłaszczał wszystko do jednej sekcji
/// (utrata danych przy autosave). Teraz: geometria/nagłówki z PIERWSZEJ sekcji + niewidoczne
/// markery div.docx-section-break niosące geometrię kolejnych sekcji w data-*, odtwarzane
/// przez writer jako paragraph-level sectPr.
/// </summary>
[TestFixture]
public class MultiSectionFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    // A4: portrait 11906x16838 twips, landscape odwrotnie.
    private const int A4W = 11906;
    private const int A4H = 16838;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    /// <summary>
    /// Sekcja 1: A4 pionowo, marginesy 1417 tw (2.5 cm), nagłówek "SEC1-HEADER".
    /// Sekcja 2: A4 poziomo, marginesy 720 tw (1.27 cm).
    /// </summary>
    private static MemoryStream BuildTwoSectionDocx(
        bool headerOnFirstSection = true,
        SectionMarkValues? secondSectionType = null)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text("SEC1-HEADER"))));
            headerPart.Header.Save();
            var headerId = mainPart.GetIdOfPart(headerPart);

            body.Append(new Paragraph(new Run(new Text("Sekcja pionowa"))));

            // Koniec sekcji 1 — paragraph-level sectPr.
            var sect1 = new SectionProperties();
            if (headerOnFirstSection)
                sect1.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId });
            sect1.Append(new PageSize { Width = A4W, Height = A4H });
            sect1.Append(new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(new Paragraph(new ParagraphProperties(sect1)));

            body.Append(new Paragraph(new Run(new Text("Sekcja pozioma"))));

            // Sekcja 2 (ostatnia) — body-level sectPr, landscape.
            var sect2 = new SectionProperties();
            if (!headerOnFirstSection)
                sect2.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId });
            if (secondSectionType is { } t)
                sect2.Append(new SectionType { Val = t });
            sect2.Append(new PageSize { Width = A4H, Height = A4W, Orient = PageOrientationValues.Landscape });
            sect2.Append(new PageMargin { Top = 720, Bottom = 720, Left = 720, Right = 720, Header = 360, Footer = 360 });
            body.Append(sect2);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    // ---------- READ: geometria i nagłówki z PIERWSZEJ sekcji ----------

    [Test]
    public void Read_TwoSections_PageSizeComesFromFirstSection_NotLast()
    {
        using var stream = BuildTwoSectionDocx();

        var content = _reader.Convert(stream);

        content.PageSize.Should().NotBeNull();
        content.PageSize!.Orientation.Should().Be("portrait");
        content.PageSize.WidthCm.Should().BeApproximately(21.0, 0.05);
        content.PageSize.HeightCm.Should().BeApproximately(29.7, 0.05);
    }

    [Test]
    public void Read_TwoSections_MarginsComeFromFirstSection()
    {
        using var stream = BuildTwoSectionDocx();

        var content = _reader.Convert(stream);

        content.Margins.Should().NotBeNull();
        content.Margins!.Top.Should().BeApproximately(2.5, 0.05);
        content.Margins.Left.Should().BeApproximately(2.5, 0.05);
    }

    [Test]
    public void Read_TwoSections_HeaderResolvedFromFirstSection()
    {
        using var stream = BuildTwoSectionDocx(headerOnFirstSection: true);

        var content = _reader.Convert(stream);

        content.Header.Should().NotBeNull();
        content.Header!.Html.Should().Contain("SEC1-HEADER");
    }

    [Test]
    public void Read_HeaderOnlyOnLastSection_StillResolved()
    {
        // Sekcja 1 bez referencji — model ma jeden nagłówek, więc bierzemy pierwszą
        // sekcję Z referencją (lepsze niż brak nagłówka).
        using var stream = BuildTwoSectionDocx(headerOnFirstSection: false);

        var content = _reader.Convert(stream);

        content.Header.Should().NotBeNull();
        content.Header!.Html.Should().Contain("SEC1-HEADER");
    }

    // ---------- READ: marker przerwy sekcji ----------

    [Test]
    public void Read_TwoSections_EmitsSectionBreakMarkerWithNextSectionGeometry()
    {
        using var stream = BuildTwoSectionDocx();

        var content = _reader.Convert(stream);

        content.Html.Should().Contain("docx-section-break");
        content.Html.Should().Contain("data-break-type=\"nextPage\"");
        // Geometria sekcji NASTĘPNEJ (poziomej).
        content.Html.Should().Contain("data-orientation=\"landscape\"");
        content.Html.Should().Contain("data-page-width-cm=\"29.7");
        content.Html.Should().Contain("data-margin-top-cm=\"1.27\"");
        // Przerwa nextPage łamie stronę w edytorze.
        content.Html.Should().Contain("class=\"page-break\"");
    }

    [Test]
    public void Read_ContinuousSectionBreak_EmitsMarkerWithoutPageBreak()
    {
        using var stream = BuildTwoSectionDocx(secondSectionType: SectionMarkValues.Continuous);

        var content = _reader.Convert(stream);

        content.Html.Should().Contain("data-break-type=\"continuous\"");
        content.Html.Should().NotContain("class=\"page-break\"");
    }

    [Test]
    public void Read_BareSectionMarkParagraph_DoesNotEmitEmptyContentLine()
    {
        // Pusty akapit niosący wyłącznie pPr/sectPr to w Wordzie ZNAK przerwy sekcji —
        // nie renderuje osobnej pustej linii. Emisja <p>&nbsp;</p> dawała widoczny „enter"
        // przed sekcją (np. kolumnową), a writer duplikował akapit przy każdym zapisie.
        using var stream = BuildTwoSectionDocx(secondSectionType: SectionMarkValues.Continuous);

        var content = _reader.Convert(stream);

        content.Html.Should().Contain("docx-section-break");
        content.Html.Should().NotContain("&nbsp;</p><div class=\"docx-section-break\"",
            "goły akapit sectPr nie może emitować pustej linii treści przed markerem");
        content.Html.Should().NotContain(">&nbsp;</p>", "fixture nie ma innych pustych akapitów");
    }

    [Test]
    public void Read_SectionEndParagraphWithText_KeepsContentBeforeMarker()
    {
        // Akapit kończący sekcję MOŻE nieść treść — wtedy treść zostaje, marker za nią.
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            var sect1 = new SectionProperties(
                new PageSize { Width = A4W, Height = A4H },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(new Paragraph(
                new ParagraphProperties(sect1),
                new Run(new Text("Ostatni akapit sekcji 1"))));
            body.Append(new Paragraph(new Run(new Text("Sekcja 2"))));
            body.Append(new SectionProperties(
                new SectionType { Val = SectionMarkValues.Continuous },
                new PageSize { Width = A4W, Height = A4H },
                new PageMargin { Top = 720, Bottom = 720, Left = 720, Right = 720, Header = 360, Footer = 360 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;

        var content = _reader.Convert(ms);

        content.Html.Should().Contain("Ostatni akapit sekcji 1");
        content.Html.Should().Contain("docx-section-break");
        content.Html.IndexOf("docx-section-break", StringComparison.Ordinal)
            .Should().BeGreaterThan(content.Html.IndexOf("Ostatni akapit sekcji 1", StringComparison.Ordinal));
    }

    [Test]
    public void RoundTrip_BareSectionMarkParagraph_DoesNotDuplicateParagraphs()
    {
        // Reader nie emituje pustej linii, writer odtwarza w:p/pPr/sectPr z markera —
        // liczba akapitów body musi się zgadzać z oryginałem (3: treść, znak sekcji, treść).
        using var stream = BuildTwoSectionDocx(secondSectionType: SectionMarkValues.Continuous);
        var content = _reader.Convert(stream);

        var docxBytes = _writer.Convert(content.Html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docxBytes), false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();
        paragraphs.Should().HaveCount(3);
        paragraphs.Count(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
            .Should().Be(1, "dokładnie jeden akapit-znak przerwy sekcji");
    }

    [Test]
    public void Read_SingleSection_EmitsNoSectionBreakMarker()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var sectPr = new SectionProperties(new PageSize { Width = A4W, Height = A4H });
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Jedna sekcja"))), sectPr));
            mainPart.Document.Save();
        }
        ms.Position = 0;

        var content = _reader.Convert(ms);

        content.Html.Should().NotContain("docx-section-break");
    }

    // ---------- WRITE: marker → paragraph-level sectPr ----------

    private const string TwoSectionHtml =
        "<p>Sekcja pionowa</p>" +
        "<div class=\"page-break\"></div>" +
        "<div class=\"docx-section-break\" data-break-type=\"nextPage\"" +
        " data-page-width-cm=\"29.7\" data-page-height-cm=\"21\" data-orientation=\"landscape\"" +
        " data-margin-top-cm=\"1.27\" data-margin-bottom-cm=\"1.27\"" +
        " data-margin-left-cm=\"1.27\" data-margin-right-cm=\"1.27\"" +
        " data-header-distance-cm=\"0.64\" data-footer-distance-cm=\"0.64\"></div>" +
        "<p>Sekcja pozioma</p>";

    private static (Body body, List<SectionProperties> orderedSections) OpenBody(byte[] docx)
    {
        var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var ordered = body.Descendants<SectionProperties>()
            .Where(sp => sp.Parent is ParagraphProperties)
            .ToList();
        var bodyLevel = body.Elements<SectionProperties>().FirstOrDefault();
        if (bodyLevel != null) ordered.Add(bodyLevel);
        return (body, ordered);
    }

    [Test]
    public void Write_SectionMarker_RecreatesParagraphLevelSectPr()
    {
        var bytes = _writer.Convert(TwoSectionHtml,
            margins: new Domain.Models.PageMargins { Top = 2.5, Bottom = 2.5, Left = 2.5, Right = 2.5 },
            pageSize: new Domain.Models.PageSize { WidthCm = 21, HeightCm = 29.7, Orientation = "portrait" });

        var (_, sections) = OpenBody(bytes);

        sections.Should().HaveCount(2);

        // Sekcja 1 (paragraph-level): geometria z argumentów Convert (pionowa A4).
        var s1Size = sections[0].GetFirstChild<PageSize>()!;
        ((int)s1Size.Width!.Value).Should().BeCloseTo(A4W, 10);
        ((int)s1Size.Height!.Value).Should().BeCloseTo(A4H, 10);
        (s1Size.Orient == null || s1Size.Orient.Value == PageOrientationValues.Portrait).Should().BeTrue();

        // Sekcja 2 (body-level): geometria z markera (pozioma, marginesy 1.27 cm).
        var s2Size = sections[1].GetFirstChild<PageSize>()!;
        s2Size.Orient!.Value.Should().Be(PageOrientationValues.Landscape);
        ((int)s2Size.Width!.Value).Should().BeCloseTo(A4H, 10);
        var s2Margin = sections[1].GetFirstChild<PageMargin>()!;
        ((int)s2Margin.Top!.Value).Should().BeCloseTo(720, 10);
        ((int)s2Margin.Header!.Value).Should().BeCloseTo(363, 10); // 0.64 cm
    }

    [Test]
    public void Write_PageBreakFollowedBySectionMarker_DoesNotEmitExtraPageBreakRun()
    {
        var bytes = _writer.Convert(TwoSectionHtml);

        var (body, _) = OpenBody(bytes);

        body.Descendants<Break>()
            .Where(b => b.Type != null && b.Type.Value == BreakValues.Page)
            .Should().BeEmpty("sectPr typu nextPage sam łamie stronę — dodatkowy w:br dawałby pustą stronę");
    }

    [Test]
    public void Write_HeaderReference_LandsOnFirstSection_SoAllSectionsInheritIt()
    {
        var header = new Domain.Models.HeaderFooterContent { Html = "<p>NAGŁÓWEK</p>", Height = 1.25 };

        var bytes = _writer.Convert(TwoSectionHtml, header: header);

        var (_, sections) = OpenBody(bytes);
        sections.Should().HaveCount(2);
        sections[0].Elements<HeaderReference>().Should().NotBeEmpty(
            "referencja na pierwszej sekcji jest dziedziczona przez kolejne; na ostatniej — wcześniejsze strony nie miałyby nagłówka");
    }

    [Test]
    public void Write_ContinuousMarker_EmitsSectionTypeOnBodySectPr()
    {
        var html =
            "<p>A</p>" +
            "<div class=\"docx-section-break\" data-break-type=\"continuous\"" +
            " data-page-width-cm=\"21\" data-page-height-cm=\"29.7\" data-orientation=\"portrait\"></div>" +
            "<p>B</p>";

        var bytes = _writer.Convert(html);

        var (_, sections) = OpenBody(bytes);
        sections.Should().HaveCount(2);
        // Typ przerwy opisuje, jak zaczyna się sekcja OSTATNIA (body-level).
        sections[1].GetFirstChild<SectionType>()!.Val!.Value.Should().Be(SectionMarkValues.Continuous);
        sections[0].GetFirstChild<SectionType>().Should().BeNull("pierwsza sekcja nie ma przerwy przed sobą");
    }

    // ---------- FULL ROUND-TRIP: DOCX → HTML → DOCX ----------

    [Test]
    public void RoundTrip_TwoSections_PreservesBothOrientations()
    {
        using var stream = BuildTwoSectionDocx();
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html,
            header: content.Header, footer: content.Footer,
            margins: content.Margins, pageSize: content.PageSize);

        var (_, sections) = OpenBody(bytes);
        sections.Should().HaveCount(2, "dokument wielosekcyjny nie może być spłaszczany do jednej sekcji przy zapisie");

        var s1 = sections[0].GetFirstChild<PageSize>()!;
        (s1.Orient == null || s1.Orient.Value == PageOrientationValues.Portrait).Should().BeTrue();

        var s2 = sections[1].GetFirstChild<PageSize>()!;
        s2.Orient!.Value.Should().Be(PageOrientationValues.Landscape);
        ((int)s2Width(s2)).Should().BeGreaterThan((int)s2Height(s2));

        static uint s2Width(PageSize p) => p.Width!.Value;
        static uint s2Height(PageSize p) => p.Height!.Value;
    }

    // ---------- Nagłówki/stopki per sekcja (SectionHeadersFooters) ----------

    /// <summary>Dwie sekcje, każda z WŁASNYM nagłówkiem default.</summary>
    private static MemoryStream BuildTwoSectionDocxWithTwoHeaders()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            HeaderPart AddHeader(string text)
            {
                var part = mainPart.AddNewPart<HeaderPart>();
                part.Header = new Header(new Paragraph(new Run(new Text(text))));
                part.Header.Save();
                return part;
            }

            var h1 = AddHeader("SEC1-HEADER");
            var h2 = AddHeader("SEC2-HEADER");

            body.Append(new Paragraph(new Run(new Text("Sekcja 1"))));
            var sect1 = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(h1) },
                new PageSize { Width = A4W, Height = A4H },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(new Paragraph(new ParagraphProperties(sect1)));

            body.Append(new Paragraph(new Run(new Text("Sekcja 2"))));
            var sect2 = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(h2) },
                new PageSize { Width = A4H, Height = A4W, Orient = PageOrientationValues.Landscape },
                new PageMargin { Top = 720, Bottom = 720, Left = 720, Right = 720, Header = 360, Footer = 360 });
            body.Append(sect2);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_TwoSectionsWithOwnHeaders_ExposesSectionEntry()
    {
        using var stream = BuildTwoSectionDocxWithTwoHeaders();

        var content = _reader.Convert(stream);

        content.Header!.Html.Should().Contain("SEC1-HEADER", "baza = pierwsza sekcja");
        content.SectionHeadersFooters.Should().NotBeNull();
        var entry = content.SectionHeadersFooters!.Single();
        entry.SectionIndex.Should().Be(1);
        entry.Header!.Html.Should().Contain("SEC2-HEADER");
    }

    [Test]
    public void Read_SecondSectionInheritsHeader_NoSectionEntry()
    {
        // Sekcja 2 bez własnych referencji → dziedziczy z sekcji 1 (brak wpisu).
        using var stream = BuildTwoSectionDocx(headerOnFirstSection: true);

        var content = _reader.Convert(stream);

        content.SectionHeadersFooters.Should().BeNull();
    }

    [Test]
    public void Write_SectionHeaderEntry_LandsOnThatSectionsSectPr()
    {
        var header = new Domain.Models.HeaderFooterContent { Html = "<p>BAZOWY</p>", Height = 1.25 };
        var sectionHf = new List<Domain.Models.SectionHeaderFooter>
        {
            new()
            {
                SectionIndex = 1,
                Header = new Domain.Models.HeaderFooterContent { Html = "<p>SEKCYJNY</p>", Height = 1.25 }
            }
        };

        var bytes = _writer.Convert(TwoSectionHtml, header: header, sectionHeadersFooters: sectionHf);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;
        var paragraphSect = body.Descendants<SectionProperties>().First(sp => sp.Parent is ParagraphProperties);
        var bodySect = body.Elements<SectionProperties>().First();

        // Baza (sekcja 0) na pierwszym sectPr; własny nagłówek sekcji 1 na body-level sectPr.
        var baseRef = paragraphSect.Elements<HeaderReference>().Single();
        var sectionRef = bodySect.Elements<HeaderReference>().Single();
        (mainPart.GetPartById(baseRef.Id!.Value!) as HeaderPart)!.Header!.InnerText.Should().Contain("BAZOWY");
        (mainPart.GetPartById(sectionRef.Id!.Value!) as HeaderPart)!.Header!.InnerText.Should().Contain("SEKCYJNY");
    }

    [Test]
    public void RoundTrip_TwoSectionHeaders_SurviveDocxHtmlDocx()
    {
        using var stream = BuildTwoSectionDocxWithTwoHeaders();
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html,
            header: content.Header, footer: content.Footer,
            margins: content.Margins, pageSize: content.PageSize,
            sectionHeadersFooters: content.SectionHeadersFooters);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.Header!.Html.Should().Contain("SEC1-HEADER");
        reimported.SectionHeadersFooters.Should().NotBeNull();
        reimported.SectionHeadersFooters!.Single().Header!.Html.Should().Contain("SEC2-HEADER");
    }

    [Test]
    public void Write_SectionFooterOnlyEntry_LandsOnThatSectionsSectPr()
    {
        // Wpis wyłącznie ze stopką (bez nagłówka) — referencja tylko na sectPr sekcji 1.
        var sectionHf = new List<Domain.Models.SectionHeaderFooter>
        {
            new()
            {
                SectionIndex = 1,
                Footer = new Domain.Models.HeaderFooterContent { Html = "<p>STOPKA-SEKCJI-2</p>", Height = 1.25 }
            }
        };

        var bytes = _writer.Convert(TwoSectionHtml, sectionHeadersFooters: sectionHf);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;
        var paragraphSect = body.Descendants<SectionProperties>().First(sp => sp.Parent is ParagraphProperties);
        var bodySect = body.Elements<SectionProperties>().First();

        paragraphSect.Elements<FooterReference>().Should().BeEmpty("wpis dotyczy sekcji 1, nie sekcji 0");
        paragraphSect.Elements<HeaderReference>().Should().BeEmpty();
        bodySect.Elements<HeaderReference>().Should().BeEmpty("wpis nie definiuje nagłówka");
        var footerRef = bodySect.Elements<FooterReference>().Single();
        (mainPart.GetPartById(footerRef.Id!.Value!) as FooterPart)!.Footer!.InnerText
            .Should().Contain("STOPKA-SEKCJI-2");
    }

    [Test]
    public void Write_SectionEntryBeyondExistingSections_IsIgnoredWithoutError()
    {
        // Użytkownik mógł skasować marker sekcji w edytorze — wpis wskazuje wtedy sekcję,
        // której już nie ma. Zapis nie może się wywrócić ani zostawić osieroconych partów.
        var sectionHf = new List<Domain.Models.SectionHeaderFooter>
        {
            new()
            {
                SectionIndex = 5,
                Header = new Domain.Models.HeaderFooterContent { Html = "<p>OSIEROCONY</p>", Height = 1.25 }
            }
        };

        var bytes = _writer.Convert(TwoSectionHtml, sectionHeadersFooters: sectionHf);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<HeaderReference>().Should().BeEmpty();
        doc.MainDocumentPart.HeaderParts.Should().BeEmpty();
    }

    [Test]
    public void Write_SectionEntryForSectionZero_IsIgnored_BaseFieldsOwnIt()
    {
        // Kontrakt: sekcja 0 żyje w polach bazowych Header/Footer — wpis z indeksem 0
        // nie może dublować referencji na pierwszym sectPr.
        var sectionHf = new List<Domain.Models.SectionHeaderFooter>
        {
            new()
            {
                SectionIndex = 0,
                Header = new Domain.Models.HeaderFooterContent { Html = "<p>DUBEL</p>", Height = 1.25 }
            }
        };

        var bytes = _writer.Convert(TwoSectionHtml, sectionHeadersFooters: sectionHf);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        doc.MainDocumentPart!.Document!.Body!.Descendants<HeaderReference>().Should().BeEmpty();
        doc.MainDocumentPart.HeaderParts.Should().BeEmpty();
    }

    [Test]
    public void RoundTrip_TwoSections_ReimportShowsFirstSectionGeometry()
    {
        using var stream = BuildTwoSectionDocx();
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html,
            header: content.Header, footer: content.Footer,
            margins: content.Margins, pageSize: content.PageSize);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.PageSize!.Orientation.Should().Be("portrait");
        reimported.Margins!.Top.Should().BeApproximately(2.5, 0.06);
        reimported.Html.Should().Contain("docx-section-break", "marker musi przetrwać kolejne round-tripy");
        reimported.Html.Should().Contain("data-orientation=\"landscape\"");
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Word's "Different first page" (w:titlePg) semantics. Opting in means the first page of
/// the section NEVER shows the default header/footer: a missing or empty first part is a
/// BLANK band, not a fallback to the default. The setting and the first-page part must also
/// survive export (a first-only header without a default part used to disappear on save).
/// </summary>
[TestFixture]
public class TitlePageHeaderFidelityTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static HeaderPart AddHeader(MainDocumentPart mainPart, string? text)
    {
        var part = mainPart.AddNewPart<HeaderPart>();
        part.Header = text == null
            ? new Header(new Paragraph())
            : new Header(new Paragraph(new Run(new Text(text))));
        part.Header.Save();
        return part;
    }

    private static FooterPart AddFooter(MainDocumentPart mainPart, string? text)
    {
        var part = mainPart.AddNewPart<FooterPart>();
        part.Footer = text == null
            ? new Footer(new Paragraph())
            : new Footer(new Paragraph(new Run(new Text(text))));
        part.Footer.Save();
        return part;
    }

    /// <param name="titlePgVal">null = element without w:val; otherwise the raw attribute value.</param>
    private static MemoryStream BuildDocx(
        bool withTitlePg,
        string? titlePgVal = null,
        bool withFirstHeaderRef = true,
        bool emptyFirstHeader = false,
        bool withEvenAndOddSetting = false,
        bool withEvenHeaderRef = false)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            if (withEvenAndOddSetting)
            {
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(new EvenAndOddHeaders());
                settingsPart.Settings.Save();
            }

            var defaultHeader = AddHeader(mainPart, "DEFAULT-HEADER");
            var defaultFooter = AddFooter(mainPart, "DEFAULT-FOOTER");

            var sectPr = new SectionProperties();
            sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultHeader) });
            sectPr.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultFooter) });

            if (withFirstHeaderRef)
            {
                var firstHeader = AddHeader(mainPart, emptyFirstHeader ? null : "FIRST-HEADER");
                sectPr.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstHeader) });
            }
            if (withEvenHeaderRef)
            {
                var evenHeader = AddHeader(mainPart, "EVEN-HEADER");
                sectPr.Append(new HeaderReference { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(evenHeader) });
            }

            if (withTitlePg)
            {
                var titlePg = new TitlePage();
                if (titlePgVal != null)
                    titlePg.SetAttribute(new OpenXmlAttribute("w", "val", W, titlePgVal));
                sectPr.Append(titlePg);
            }

            sectPr.Append(new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(new Paragraph(new Run(new Text("body"))));
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    // ---------- READ: titlePg + missing/empty first part ----------

    [Test]
    public void Read_TitlePg_EmptyFirstPart_FirstPageIsBlank_NotDefault()
    {
        using var stream = BuildDocx(withTitlePg: true, emptyFirstHeader: true);

        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeTrue("titlePg is on");
        content.Header.FirstPageHtml.Should().NotBeNull();
        content.Header.FirstPageHtml.Should().NotContain("DEFAULT-HEADER",
            "an empty first part means a BLANK first page, not the default header");
        content.Header.Html.Should().Contain("DEFAULT-HEADER");
    }

    [Test]
    public void Read_TitlePg_NoFirstReference_FirstPageIsBlank_NotDefault()
    {
        using var stream = BuildDocx(withTitlePg: true, withFirstHeaderRef: false);

        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeTrue();
        content.Header.FirstPageHtml.Should().Be(string.Empty,
            "section 0 has no previous section to inherit from — the band is blank");
        content.Header.Html.Should().Contain("DEFAULT-HEADER");
    }

    [Test]
    public void Read_TitlePg_AppliesToFooterIndependently()
    {
        // No first FOOTER reference at all — the first-page footer must still be blank.
        using var stream = BuildDocx(withTitlePg: true);

        var content = _reader.Convert(stream);

        content.Footer!.DifferentFirstPage.Should().BeTrue();
        content.Footer.FirstPageHtml.Should().Be(string.Empty);
        content.Footer.Html.Should().Contain("DEFAULT-FOOTER");
    }

    // ---------- READ: ST_OnOff parsing of w:titlePg ----------

    [TestCase(null, true, TestName = "Read_TitlePg_NoValAttribute_MeansOn")]
    [TestCase("true", true)]
    [TestCase("1", true)]
    [TestCase("on", true)]
    [TestCase("false", false)]
    [TestCase("0", false)]
    [TestCase("off", false)]
    public void Read_TitlePg_OnOffValues(string? val, bool expected)
    {
        using var stream = BuildDocx(withTitlePg: true, titlePgVal: val);

        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().Be(expected);
        if (!expected)
            content.Header.FirstPageHtml.Should().BeNull("titlePg is explicitly disabled");
    }

    [Test]
    public void Read_NoTitlePg_NoFirstPageVariant_EvenWhenFirstPartExists()
    {
        using var stream = BuildDocx(withTitlePg: false);

        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeFalse();
        content.Header.FirstPageHtml.Should().BeNull();
    }

    // ---------- READ: evenAndOddHeaders without an even reference ----------

    [Test]
    public void Read_EvenAndOdd_NoEvenReference_EvenPagesBlank_NotDefault()
    {
        using var stream = BuildDocx(withTitlePg: false, withEvenAndOddSetting: true, withEvenHeaderRef: false);

        var content = _reader.Convert(stream);

        content.Header!.DifferentOddEven.Should().BeTrue("evenAndOddHeaders is on");
        content.Header.EvenHtml.Should().Be(string.Empty,
            "no even part referenced — even pages are blank in Word, not the default header");
    }

    // ---------- WRITE: titlePg and the first part must survive export ----------

    [Test]
    public void Write_TitlePgWithBlankFirstPage_KeepsTitlePgAndValidFirstPart()
    {
        var header = new Domain.Models.HeaderFooterContent
        {
            Html = "<p>DEF</p>",
            Height = 1.25,
            DifferentFirstPage = true,
            FirstPageHtml = string.Empty
        };

        var bytes = _writer.Convert("<p>body</p>", header: header);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().First();
        sectPr.Elements<TitlePage>().Should().NotBeEmpty("titlePg must survive a blank first-page band");
        var firstRef = sectPr.Elements<HeaderReference>().Single(r => r.Type!.Value == HeaderFooterValues.First);
        var firstPart = (HeaderPart)doc.MainDocumentPart.GetPartById(firstRef.Id!.Value!);
        firstPart.Header!.Elements<Paragraph>().Should().NotBeEmpty("CT_HdrFtr requires a block-level child");

        var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc).ToList();
        errors.Should().BeEmpty();
    }

    [Test]
    public void Write_TitlePgWithoutFirstPart_KeepsTitlePg_WithoutFirstReference()
    {
        var header = new Domain.Models.HeaderFooterContent
        {
            Html = "<p>DEF</p>",
            Height = 1.25,
            DifferentFirstPage = true,
            FirstPageHtml = null
        };

        var bytes = _writer.Convert("<p>body</p>", header: header);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().First();
        sectPr.Elements<TitlePage>().Should().NotBeEmpty();
        sectPr.Elements<HeaderReference>().Should().NotContain(r => r.Type!.Value == HeaderFooterValues.First);
    }

    [Test]
    public void Write_FirstOnlyHeader_WithoutDefault_SurvivesSave()
    {
        // The Qutalo-class document: titlePg + first-page header, intentionally NO default.
        // The old writer skipped ALL variants when Html was empty — header lost on 1st save.
        var header = new Domain.Models.HeaderFooterContent
        {
            Html = string.Empty,
            Height = 1.25,
            DifferentFirstPage = true,
            FirstPageHtml = "<p>FIRST-ONLY</p>"
        };

        var bytes = _writer.Convert("<p>body</p>", header: header);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var mainPart = doc.MainDocumentPart!;
        var sectPr = mainPart.Document!.Body!.Elements<SectionProperties>().First();
        sectPr.Elements<TitlePage>().Should().NotBeEmpty();
        sectPr.Elements<HeaderReference>().Should().NotContain(r => r.Type!.Value == HeaderFooterValues.Default,
            "ordinary pages are intentionally blank");
        var firstRef = sectPr.Elements<HeaderReference>().Single(r => r.Type!.Value == HeaderFooterValues.First);
        ((HeaderPart)mainPart.GetPartById(firstRef.Id!.Value!)).Header!.InnerText.Should().Contain("FIRST-ONLY");
    }

    // ---------- ROUND-TRIP: save and reopen must not lose the setting ----------

    [Test]
    public void RoundTrip_TitlePgEmptyFirst_PreservedAcrossSaveAndReopen()
    {
        using var stream = BuildDocx(withTitlePg: true, emptyFirstHeader: true);
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html, header: content.Header, footer: content.Footer);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.Header!.DifferentFirstPage.Should().BeTrue();
        reimported.Header.FirstPageHtml.Should().NotContain("DEFAULT-HEADER");
        reimported.Header.Html.Should().Contain("DEFAULT-HEADER");
    }

    [Test]
    public void RoundTrip_FirstOnlyHeader_PreservedAcrossSaveAndReopen()
    {
        var header = new Domain.Models.HeaderFooterContent
        {
            Html = string.Empty,
            Height = 1.25,
            DifferentFirstPage = true,
            FirstPageHtml = "<p>FIRST-ONLY</p>"
        };

        var bytes = _writer.Convert("<p>body</p>", header: header);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.Header!.DifferentFirstPage.Should().BeTrue();
        reimported.Header.FirstPageHtml.Should().Contain("FIRST-ONLY");
        reimported.Header.Html.Should().NotContain("FIRST-ONLY");
    }

    // ---------- MULTI-SECTION: per-section titlePg in section entries ----------

    private static MemoryStream BuildTwoSectionDocxWithTitlePgOnSecond(bool secondHasFirstRef)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var h1 = AddHeader(mainPart, "SEC1-HEADER");
            var h2 = AddHeader(mainPart, "SEC2-DEFAULT");

            body.Append(new Paragraph(new Run(new Text("Sekcja 1"))));
            var sect1 = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(h1) },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(new Paragraph(new ParagraphProperties(sect1)));

            body.Append(new Paragraph(new Run(new Text("Sekcja 2"))));
            var sect2 = new SectionProperties();
            sect2.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(h2) });
            if (secondHasFirstRef)
            {
                var h2First = AddHeader(mainPart, "SEC2-FIRST");
                sect2.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(h2First) });
            }
            sect2.Append(new TitlePage());
            sect2.Append(new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708 });
            body.Append(sect2);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void Read_SecondSectionWithTitlePg_ExposesFirstVariantInSectionEntry()
    {
        using var stream = BuildTwoSectionDocxWithTitlePgOnSecond(secondHasFirstRef: true);

        var content = _reader.Convert(stream);

        content.Header!.DifferentFirstPage.Should().BeFalse("section 1 has no titlePg");
        var entry = content.SectionHeadersFooters!.Single(e => e.SectionIndex == 1);
        entry.Header!.DifferentFirstPage.Should().BeTrue();
        entry.Header.FirstPageHtml.Should().Contain("SEC2-FIRST");
        entry.Header.Html.Should().Contain("SEC2-DEFAULT");
    }

    [Test]
    public void Read_SecondSectionWithTitlePg_NoFirstRef_InheritsFirstVariant()
    {
        using var stream = BuildTwoSectionDocxWithTitlePgOnSecond(secondHasFirstRef: false);

        var content = _reader.Convert(stream);

        var entry = content.SectionHeadersFooters!.Single(e => e.SectionIndex == 1);
        entry.Header!.DifferentFirstPage.Should().BeTrue();
        entry.Header.FirstPageHtml.Should().BeNull("no own first part — inherited from the previous section");
    }

    [Test]
    public void Write_SectionEntryWithTitlePg_LandsOnThatSectionsSectPr()
    {
        const string twoSectionHtml =
            "<p>Sekcja 1</p>" +
            "<div class=\"page-break\"></div>" +
            "<div class=\"docx-section-break\" data-break-type=\"nextPage\"" +
            " data-page-width-cm=\"21\" data-page-height-cm=\"29.7\" data-orientation=\"portrait\"" +
            " data-margin-top-cm=\"2.5\" data-margin-bottom-cm=\"2.5\"" +
            " data-margin-left-cm=\"2.5\" data-margin-right-cm=\"2.5\"></div>" +
            "<p>Sekcja 2</p>";
        var sectionHf = new List<Domain.Models.SectionHeaderFooter>
        {
            new()
            {
                SectionIndex = 1,
                Header = new Domain.Models.HeaderFooterContent
                {
                    Html = "<p>S2-DEF</p>",
                    Height = 1.25,
                    DifferentFirstPage = true,
                    FirstPageHtml = "<p>S2-FIRST</p>"
                }
            }
        };

        var bytes = _writer.Convert(twoSectionHtml, sectionHeadersFooters: sectionHf);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document!.Body!;
        var paragraphSect = body.Descendants<SectionProperties>().First(sp => sp.Parent is ParagraphProperties);
        var bodySect = body.Elements<SectionProperties>().First();

        paragraphSect.Elements<TitlePage>().Should().BeEmpty("section 0 has no titlePg entry");
        bodySect.Elements<TitlePage>().Should().NotBeEmpty();
        var firstRef = bodySect.Elements<HeaderReference>().Single(r => r.Type!.Value == HeaderFooterValues.First);
        ((HeaderPart)mainPart.GetPartById(firstRef.Id!.Value!)).Header!.InnerText.Should().Contain("S2-FIRST");
    }
}

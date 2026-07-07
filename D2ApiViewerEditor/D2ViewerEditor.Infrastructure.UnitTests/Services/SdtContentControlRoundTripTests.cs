using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Formanty (SDT / Content Controls / „obszary specjalne" Worda) muszą zachować TYP i
/// właściwości przez round-trip DOCX→HTML→DOCX. Wcześniej writer odtwarzał tylko tag/alias,
/// więc dropdown/date/checkbox degradowały do generycznego formantu przy 1. eksporcie/autosave.
/// Reader niesie pełne `w:sdtPr` w `data-sdt-props` (base64), writer odtwarza je 1:1.
/// </summary>
[TestFixture]
public class SdtContentControlRoundTripTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream DocxWith(OpenXmlElement bodyChild)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(bodyChild, new Paragraph()));
            main.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void DropDownList_TypeTagAndOptions_SurviveRoundTrip()
    {
        var sdt = new SdtBlock(
            new SdtProperties(
                new Tag { Val = "Wybor" },
                new SdtAlias { Val = "Lista" },
                new SdtId { Val = 111 },
                new SdtContentDropDownList(
                    new ListItem { DisplayText = "Opcja A", Value = "A" },
                    new ListItem { DisplayText = "Opcja B", Value = "B" })),
            new SdtContentBlock(new Paragraph(new Run(new Text("Opcja A")))));

        using var ms = DocxWith(sdt);
        var html = _reader.Convert(ms).Html;
        html.Should().Contain("data-sdt-props=");
        html.Should().Contain("Opcja A");

        var outSdt = RoundTripFirstSdt(html);
        var props = outSdt.SdtProperties!;
        props.Elements<Tag>().First().Val!.Value.Should().Be("Wybor");
        var ddl = props.Elements<SdtContentDropDownList>().Should().ContainSingle().Subject;
        var items = ddl.Elements<ListItem>().ToList();
        items.Should().HaveCount(2);
        items[0].DisplayText!.Value.Should().Be("Opcja A");
        items[1].Value!.Value.Should().Be("B");
    }

    [Test]
    public void DatePicker_TypeAndFormat_SurviveRoundTrip()
    {
        var sdt = new SdtBlock(
            new SdtProperties(
                new Tag { Val = "Data" },
                new SdtContentDate(new DateFormat { Val = "dd.MM.yyyy" })
                { FullDate = System.DateTime.Parse("2026-07-06") }),
            new SdtContentBlock(new Paragraph(new Run(new Text("06.07.2026")))));

        using var ms = DocxWith(sdt);
        var html = _reader.Convert(ms).Html;

        var props = RoundTripFirstSdt(html).SdtProperties!;
        var date = props.Elements<SdtContentDate>().Should().ContainSingle().Subject;
        date.GetFirstChild<DateFormat>()!.Val!.Value.Should().Be("dd.MM.yyyy");
    }

    [Test]
    public void InlineTextControl_TypeAndTag_SurviveRoundTrip()
    {
        var sdt = new SdtRun(
            new SdtProperties(
                new Tag { Val = "Imie" },
                new SdtContentText()),
            new SdtContentRun(new Run(new Text("Jan"))));
        var para = new Paragraph(new Run(new Text("Klient: ")), sdt);

        using var ms = DocxWith(para);
        var html = _reader.Convert(ms).Html;
        html.Should().Contain("sdt-inline");

        var bytes = _writer.Convert(html);
        using var outDoc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var outSdt = outDoc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().First();
        outSdt.SdtProperties!.Elements<Tag>().First().Val!.Value.Should().Be("Imie");
        outSdt.SdtProperties.Elements<SdtContentText>().Should().ContainSingle();
        outSdt.Descendants<Text>().Any(t => t.Text == "Jan").Should().BeTrue();
    }

    [Test]
    public void DuplicateId_IsDropped_SoWordDoesNotSeeCorruption()
    {
        var sdt = new SdtBlock(
            new SdtProperties(new Tag { Val = "X" }, new SdtId { Val = 999 }),
            new SdtContentBlock(new Paragraph(new Run(new Text("v")))));

        using var ms = DocxWith(sdt);
        var html = _reader.Convert(ms).Html;

        var props = RoundTripFirstSdt(html).SdtProperties!;
        props.Elements<SdtId>().Should().BeEmpty("id jest usuwane — Word nadaje własne, unikając kolizji");
        props.Elements<Tag>().First().Val!.Value.Should().Be("X");
    }

    [Test]
    public void BlockContentControl_InFooter_ContentIsNotDropped()
    {
        using var ms = DocxWithFooterSdt(
            new SdtBlock(
                new SdtProperties(new Tag { Val = "removeif_nondigitalversion" }),
                new SdtContentBlock(new Paragraph(new Run(new Text(
                    "Dokument wygenerowany elektronicznie, nie wymaga pieczeci ani podpisu."))))));

        var footer = _reader.Convert(ms).Footer;
        footer.Should().NotBeNull();
        footer!.Html.Should().Contain("Dokument wygenerowany elektronicznie");
        footer.Html.Should().Contain("sdt-block");
        footer.Html.Should().Contain("data-sdt-tag=\"removeif_nondigitalversion\"");
    }

    private static MemoryStream DocxWithFooterSdt(OpenXmlElement footerChild)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var footerPart = main.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(footerChild);
            footerPart.Footer.Save();
            var rid = main.GetIdOfPart(footerPart);

            var sectPr = new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = rid });
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("body"))), sectPr));
            main.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private SdtBlock RoundTripFirstSdt(string html)
    {
        var bytes = _writer.Convert(html);
        var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().First();
    }
}

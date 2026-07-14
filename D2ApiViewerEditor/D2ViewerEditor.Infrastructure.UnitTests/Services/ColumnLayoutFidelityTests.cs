using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// ADR-0039: układ wielokolumnowy (w:cols). Import (reader) → model/HTML data-col-*, eksport
/// (writer) → w:cols, round-trip DOCX→HTML→DOCX oraz podział kolumny (w:br type=column).
/// Testy budują syntetyczne DOCX 1:1 wg semantyki OOXML (jak pozostałe *FidelityTests) — brak
/// dokumentu regulaminu w repo.
/// </summary>
[TestFixture]
public class ColumnLayoutFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    private const int A4W = 11906;
    private const int A4H = 16838;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream BuildDocx(Columns? cols, params OpenXmlElement[] bodyChildren)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            foreach (var el in bodyChildren) body.Append(el);

            var sectPr = new SectionProperties(new PageSize { Width = A4W, Height = A4H });
            if (cols != null) sectPr.Append(cols);
            sectPr.Append(new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720 });
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Body OpenBody(byte[] docx)
    {
        var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return doc.MainDocumentPart!.Document!.Body!;
    }

    // ---------- IMPORT ----------

    [Test]
    public void Read_TwoEqualColumns_EmitsBaseColumnsAndDataAttrs()
    {
        using var stream = BuildDocx(
            new Columns { ColumnCount = 2, Space = "720", Separator = true },
            new Paragraph(new Run(new Text("Regulamin dwukolumnowy"))));

        var content = _reader.Convert(stream);

        content.Columns.Should().NotBeNull();
        content.Columns!.Count.Should().Be(2);
        content.Columns.SpaceTwips.Should().Be(720);
        content.Columns.Separator.Should().BeTrue();
        // Atrybuty na kontenerze .document-content (nośnik kolumn sekcji bazowej).
        content.Html.Should().Contain("data-col-count=\"2\"");
        content.Html.Should().Contain("data-col-space-tw=\"720\"");
        content.Html.Should().Contain("data-col-sep=\"1\"");
    }

    [Test]
    public void Read_SingleColumn_DoesNotEmitColumnDataAttrs()
    {
        // qutable.docx: <w:cols w:space="708"/> — jedna kolumna. NIE może udawać wielokolumnowej.
        using var stream = BuildDocx(
            new Columns { Space = "708" },
            new Paragraph(new Run(new Text("Jednokolumnowy"))));

        var content = _reader.Convert(stream);

        content.Html.Should().NotContain("data-col-count");
        (content.Columns == null || content.Columns.Count <= 1).Should().BeTrue();
    }

    [Test]
    public void Read_ColumnBreak_EmitsMarker()
    {
        using var stream = BuildDocx(
            new Columns { ColumnCount = 2, Space = "720" },
            new Paragraph(new Run(new Text("Przed"), new Break { Type = BreakValues.Column }, new Text("Po"))));

        var content = _reader.Convert(stream);

        content.Html.Should().Contain("docx-column-break");
    }

    // ---------- EXPORT ----------

    [Test]
    public void Write_TwoEqualColumns_EmitsColsAfterPageMargin_SchemaValid()
    {
        var html = "<div class=\"document-content\" data-col-count=\"2\" data-col-space-tw=\"567\" data-col-equal=\"1\" data-col-sep=\"1\">"
                 + "<p>A</p></div>";

        var bytes = _writer.Convert(html);

        var body = OpenBody(bytes);
        var sectPr = body.Elements<SectionProperties>().Single();
        var cols = sectPr.GetFirstChild<Columns>();
        cols.Should().NotBeNull();
        cols!.ColumnCount!.Value.Should().Be(2);
        cols.Space!.Value.Should().Be("567");
        (cols.Separator?.Value ?? false).Should().BeTrue();

        // Kolejność CT_SectPr: pgSz → pgMar → cols.
        var children = sectPr.ChildElements.ToList();
        children.IndexOf(cols).Should().BeGreaterThan(children.IndexOf(sectPr.GetFirstChild<PageMargin>()!));

        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(WordprocessingDocument.Open(new MemoryStream(bytes), false)).ToList();
        errors.Should().BeEmpty("dokument z w:cols musi być schema-valid: {0}",
            string.Join(" | ", errors.Select(e => e.Description)));
    }

    [Test]
    public void Write_SingleColumnHtml_DoesNotEmitCols()
    {
        var bytes = _writer.Convert("<div class=\"document-content\"><p>A</p></div>");

        var body = OpenBody(bytes);
        body.Elements<SectionProperties>().Single().GetFirstChild<Columns>()
            .Should().BeNull("brak data-col-count > 1 = brak w:cols");
    }

    [Test]
    public void Write_UnequalColumns_EmitsIndividualColElements()
    {
        var html = "<div class=\"document-content\" data-col-count=\"2\" data-col-equal=\"0\""
                 + " data-col-widths-tw=\"3000,6000\" data-col-spaces-tw=\"500,0\"><p>A</p></div>";

        var bytes = _writer.Convert(html);

        var cols = OpenBody(bytes).Elements<SectionProperties>().Single().GetFirstChild<Columns>()!;
        (cols.EqualWidth?.Value ?? true).Should().BeFalse();
        var colList = cols.Elements<Column>().ToList();
        colList.Should().HaveCount(2);
        colList[0].Width!.Value.Should().Be("3000");
        colList[0].Space!.Value.Should().Be("500");
        colList[1].Width!.Value.Should().Be("6000");
    }

    [Test]
    public void Write_ColumnBreakDiv_EmitsBreakTypeColumn()
    {
        var html = "<div class=\"document-content\" data-col-count=\"2\" data-col-space-tw=\"720\">"
                 + "<p>Przed<div class=\"docx-column-break\"></div>Po</p></div>";

        var bytes = _writer.Convert(html);

        OpenBody(bytes).Descendants<Break>()
            .Any(b => b.Type != null && b.Type.Value == BreakValues.Column)
            .Should().BeTrue("podział kolumny musi wrócić jako w:br type=column");
    }

    // ---------- ROUND-TRIP ----------

    [Test]
    public void RoundTrip_TwoEqualColumns_PreservesLayout()
    {
        using var stream = BuildDocx(
            new Columns { ColumnCount = 2, Space = "708", Separator = true },
            new Paragraph(new Run(new Text("Treść"))));
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html,
            margins: content.Margins, pageSize: content.PageSize);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.Columns.Should().NotBeNull();
        reimported.Columns!.Count.Should().Be(2);
        reimported.Columns.SpaceTwips.Should().Be(708);
        reimported.Columns.Separator.Should().BeTrue();
    }

    [Test]
    public void RoundTrip_UnequalColumns_PreservesWidths()
    {
        var cols = new Columns { ColumnCount = 2, EqualWidth = false };
        cols.Append(new Column { Width = "3402", Space = "425" });
        cols.Append(new Column { Width = "5000" });
        using var stream = BuildDocx(cols, new Paragraph(new Run(new Text("Treść"))));
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html, margins: content.Margins, pageSize: content.PageSize);
        var reimported = _reader.Convert(new MemoryStream(bytes));

        reimported.Columns!.EqualWidth.Should().BeFalse();
        reimported.Columns.Columns.Should().HaveCount(2);
        reimported.Columns.Columns![0].WidthTwips.Should().Be(3402);
        reimported.Columns.Columns[1].WidthTwips.Should().Be(5000);
    }

    [Test]
    public void RoundTrip_SingleColumn_StaysClean()
    {
        using var stream = BuildDocx(
            new Columns { Space = "708" },
            new Paragraph(new Run(new Text("Jednokolumnowy"))));
        var content = _reader.Convert(stream);

        var bytes = _writer.Convert(content.Html, margins: content.Margins, pageSize: content.PageSize);

        OpenBody(bytes).Elements<SectionProperties>().Single().GetFirstChild<Columns>()
            .Should().BeNull("dokument jednokolumnowy nie może zyskać w:cols z >1 kolumną");
    }
}

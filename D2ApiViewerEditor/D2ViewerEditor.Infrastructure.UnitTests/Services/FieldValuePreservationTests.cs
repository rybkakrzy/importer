using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Complex Word fields must never lose their cached value on import (KR-05), and
/// DATE must keep the value stored in the document rather than being overwritten
/// with the server's current date (KR-08). Page-number fields (PAGE/NUMPAGES)
/// stay dynamic placeholders while the surrounding literal text is preserved
/// (the "Strona: 1 z 1." footer scenario, item 3).
/// </summary>
[TestFixture]
public class FieldValuePreservationTests
{
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup() => _reader = new DocxToHtmlConverter();

    [Test]
    public void ComplexRefField_KeepsCachedValueText()
    {
        // { REF _Ref1 } with cached result "Rozdział 3.2" — the value must survive.
        var docx = BuildBody(
            Field(" REF _Ref1 \\h ", "Rozdział 3.2"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("Rozdział 3.2");
    }

    [Test]
    public void ComplexDateField_KeepsStoredDate_NotServerToday()
    {
        var storedDate = "5 stycznia 2024";
        var docx = BuildBody(
            Field(" DATE \\@ \"d MMMM yyyy\" ", storedDate));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain(storedDate);
        html.Should().NotContain(System.DateTime.Now.ToString("dd.MM.yyyy"));
    }

    [Test]
    public void PageAndNumPagesFooter_KeepsLiteralText_AndDynamicPlaceholders()
    {
        // "Strona: {PAGE} z {NUMPAGES}. Przedłużenie tekstu"
        var runs = new List<OpenXmlElement> { new Run(new Text("Strona: ") { Space = SpaceProcessingModeValues.Preserve }) };
        runs.AddRange(FieldRuns(" PAGE ", "1"));
        runs.Add(new Run(new Text(" z ") { Space = SpaceProcessingModeValues.Preserve }));
        runs.AddRange(FieldRuns(" NUMPAGES ", "1"));
        runs.Add(new Run(new Text(". Przedłużenie tekstu") { Space = SpaceProcessingModeValues.Preserve }));

        var docx = BuildBody(new Paragraph(runs));
        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain("Strona:");
        html.Should().Contain(" z ");
        html.Should().Contain("Przedłużenie tekstu");
        html.Should().Contain("class=\"field-page\"");
        html.Should().Contain("class=\"field-numpages\"");
    }

    // --- builders -------------------------------------------------------------

    private static Paragraph Field(string instruction, string cachedValue) =>
        new Paragraph(FieldRuns(instruction, cachedValue).Cast<OpenXmlElement>());

    private static IEnumerable<Run> FieldRuns(string instruction, string cachedValue) => new[]
    {
        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new Run(new Text(cachedValue) { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
    };

    private static byte[] BuildBody(params Paragraph[] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var p in paragraphs) body.Append(p);
            body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }));
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }
}

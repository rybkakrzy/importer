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
    public void ComplexDateField_RendersCurrentDatePerPicture_AndCarriesInstruction()
    {
        // ADR-0084 (koryguje KR-08 dla pól AUTO-daty): DATE/TIME to w Wordzie pola
        // aktualizowane — edytor pokazuje BIEŻĄCĄ datę wg obrazu \@, a instrukcja jedzie
        // w data-fld-instr, żeby zapis odtworzył żywe pole zamiast martwego tekstu.
        var storedDate = "5 stycznia 2024";
        var docx = BuildBody(
            Field(" TIME \\@ \"dd-MM-yyyy\" ", storedDate));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain(System.DateTime.Now.ToString("dd-MM-yyyy"));
        html.Should().NotContain(storedDate);
        html.Should().Contain("class=\"field-date\"");
        html.Should().Contain("data-fld-instr=\"TIME \\@ &quot;dd-MM-yyyy&quot;\"");
    }

    [Test]
    public void ComplexDateField_PolishMonthPicture_UsesPolishCulture()
    {
        var docx = BuildBody(Field(" DATE \\@ \"d MMMM yyyy\" ", "stara wartość"));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        var expected = System.DateTime.Now.ToString(
            "d MMMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("pl-PL"));
        html.Should().Contain(expected);
    }

    [Test]
    public void ComplexCreateDateField_KeepsStoredDate_HistoricalDatesAreNotRefreshed()
    {
        // CREATEDATE/SAVEDATE/PRINTDATE niosą daty HISTORYCZNE — KR-08 pozostaje w mocy.
        var storedDate = "5 stycznia 2024";
        var docx = BuildBody(Field(" CREATEDATE \\@ \"d MMMM yyyy\" ", storedDate));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain(storedDate);
        html.Should().NotContain("data-fld-instr=\"CREATEDATE");
    }

    [Test]
    public void ComplexDateField_WithLockSwitch_KeepsStoredDate()
    {
        // \! = zablokowana aktualizacja pola — wartość z pliku zostaje.
        var storedDate = "11-05-2022";
        var docx = BuildBody(Field(" TIME \\@ \"dd-MM-yyyy\" \\! ", storedDate));

        var html = _reader.Convert(new MemoryStream(docx)).Html;

        html.Should().Contain(storedDate);
    }

    [Test]
    public void DateFieldSpan_RoundTripsToLiveSimpleField()
    {
        // Pełny round-trip: DOCX z polem TIME → HTML (span.field-date) → DOCX.
        // Wynik musi mieć ŻYWE pole (fldSimple z instrukcją), nie martwy tekst.
        var docx = BuildBody(Field(" TIME \\@ \"dd-MM-yyyy\" ", "11-05-2022"));
        var html = _reader.Convert(new MemoryStream(docx)).Html;

        var writer = new HtmlToDocxConverter();
        var bytes = writer.Convert(html);

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var field = doc.MainDocumentPart!.Document.Body!.Descendants<SimpleField>().FirstOrDefault();
        field.Should().NotBeNull("pole daty musi wrócić jako żywe pole, nie tekst");
        field!.Instruction!.Value.Should().Contain("TIME");
        field.Instruction.Value.Should().Contain("dd-MM-yyyy");
        field.InnerText.Trim().Should().Be(System.DateTime.Now.ToString("dd-MM-yyyy"),
            "wartość zbuforowana = data pokazana w edytorze");
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

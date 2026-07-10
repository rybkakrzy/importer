using System.IO.Compression;
using System.Xml.Linq;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using D2ViewerEditor.Infrastructure.UnitTests.Fixtures;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;
using WpFootnote = DocumentFormat.OpenXml.Wordprocessing.Footnote;
using DomainFootnote = D2ViewerEditor.Domain.Models.Footnote;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// End-to-end footnote coverage: import (9.1), export (9.2), round-trip (9.3), editing (9.4)
/// and the no-footnote regression (9.5), plus OOXML validation (section 12). Fixtures come from
/// <see cref="FootnoteTestDocuments"/> (built at the OOXML level, not by the production exporter).
/// </summary>
[TestFixture]
public class FootnoteFidelityTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private DocumentContent Import(byte[] docx) => _reader.Convert(new MemoryStream(docx));

    // The reader HTML-escapes non-ASCII (e.g. ó → &#243;); assert on the decoded text so we check
    // semantic content, not the entity form.
    private static string Decoded(string html) => System.Net.WebUtility.HtmlDecode(html);

    private static bool IsUserFootnote(WpFootnote footnote) =>
        footnote.Type == null ||
        (footnote.Type.Value != FootnoteEndnoteValues.Separator &&
         footnote.Type.Value != FootnoteEndnoteValues.ContinuationSeparator);

    private byte[] Export(DocumentContent content) =>
        _writer.Convert(content.Html, content.Metadata, content.Header, content.Footer,
            content.Margins, content.PageSize, content.SectionHeadersFooters, content.Footnotes);

    // ---------- 9.1 Import ----------

    [Test]
    public void Import_TwoFootnotes_ProducesModelWithReferencesAndContent()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        content.Footnotes.Should().NotBeNull();
        content.Footnotes!.Should().HaveCount(2, "dokument ma dwa odwołania do dwóch przypisów");

        content.Footnotes[0].Id.Should().Be("fn-1");
        content.Footnotes[1].Id.Should().Be("fn-2");

        Decoded(content.Footnotes[0].Html).Should().Contain("zażółć gęślą jaźń")
            .And.Contain("€");
        content.Footnotes[1].Html.Should().Contain("Drugi akapit tego samego przypisu.",
            "przypis wieloakapitowy nie może tracić treści");
    }

    [Test]
    public void Import_References_RenderAsSupWithStableIdAndDisplayNumber()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        content.Html.Should().Contain(
            "<sup class=\"footnote-ref\" data-footnote-id=\"fn-1\" aria-label=\"Przypis 1\">1</sup>");
        content.Html.Should().Contain(
            "<sup class=\"footnote-ref\" data-footnote-id=\"fn-2\" aria-label=\"Przypis 2\">2</sup>");

        // Reference order == display order.
        content.Html.IndexOf("fn-1", StringComparison.Ordinal)
            .Should().BeLessThan(content.Html.IndexOf("fn-2", StringComparison.Ordinal));
    }

    [Test]
    public void Import_PreservesFootnoteFormatting()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        // Bold + italic are rendered by the SAME converters the body uses (semantic tags).
        content.Footnotes![1].Html.Should().Contain("<strong>");
        content.Footnotes[1].Html.Should().Contain("<em>");
    }

    [Test]
    public void Import_TechnicalSeparators_AreNotUserFootnotes()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        content.Footnotes!.Select(f => f.Id).Should().OnlyContain(id => id == "fn-1" || id == "fn-2");
        content.Footnotes.Should().HaveCount(2, "separator i continuationSeparator nie są przypisami użytkownika");
    }

    [Test]
    public void Import_SharedFootnote_TwoReferencesShareOneEntryAndNumber()
    {
        var content = Import(FootnoteTestDocuments.SharedFootnoteReferencedTwice());

        content.Footnotes!.Should().HaveCount(1, "jeden przypis wskazywany dwukrotnie = jedna treść");
        System.Text.RegularExpressions.Regex.Matches(content.Html, "data-footnote-id=\"fn-1\"")
            .Count.Should().Be(2, "oba odwołania wskazują ten sam przypis");
        content.Html.Should().NotContain("fn-1\" aria-label=\"Przypis 2",
            "drugie odwołanie do tego samego przypisu współdzieli numer 1");
    }

    [Test]
    public void Import_OrphanReference_IsHandledWithoutCrash()
    {
        var content = Import(FootnoteTestDocuments.OrphanReference());

        // Import completes; the reference is preserved with empty content rather than crashing.
        content.Footnotes.Should().NotBeNull();
        content.Footnotes!.Should().ContainSingle();
        content.Footnotes[0].Html.Should().BeEmpty();
        content.Html.Should().Contain("data-footnote-id=\"fn-5\"");
    }

    // ---------- 9.2 Export ----------

    [Test]
    public void Export_ProducesValidZipWithFootnotesPartRelationshipAndContentType()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());
        var docx = Export(content);

        using var archive = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);

        archive.GetEntry("word/footnotes.xml").Should().NotBeNull("część przypisów musi istnieć");

        var contentTypes = XDocument.Load(archive.GetEntry("[Content_Types].xml")!.Open());
        contentTypes.Descendants().Any(e =>
                (string?)e.Attribute("PartName") == "/word/footnotes.xml" ||
                (string?)e.Attribute("ContentType") ==
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")
            .Should().BeTrue("wpis content-type dla footnotes.xml");

        var rels = XDocument.Load(archive.GetEntry("word/_rels/document.xml.rels")!.Open());
        rels.Descendants().Any(e => ((string?)e.Attribute("Type"))?.EndsWith("/footnotes") == true)
            .Should().BeTrue("relacja z głównej części do części przypisów");
    }

    [Test]
    public void Export_FootnotesXml_HasSeparatorsAndUserFootnotesWithUniqueIds()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());
        var docx = Export(content);

        using var archive = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        var footnotesXml = XDocument.Load(archive.GetEntry("word/footnotes.xml")!.Open());
        XName Fn(string n) => XName.Get(n, WNs);

        var footnotes = footnotesXml.Descendants(Fn("footnote")).ToList();

        // Technical separators present.
        footnotes.Any(f => (string?)f.Attribute(Fn("type")) == "separator").Should().BeTrue();
        footnotes.Any(f => (string?)f.Attribute(Fn("type")) == "continuationSeparator").Should().BeTrue();

        // Ids unique across the whole part.
        var ids = footnotes.Select(f => (int)f.Attribute(Fn("id"))!).ToList();
        ids.Should().OnlyHaveUniqueItems();

        // User footnote ids are positive and distinct from reserved technical ids (-1, 0).
        var userIds = footnotes
            .Where(f => f.Attribute(Fn("type")) == null)
            .Select(f => (int)f.Attribute(Fn("id"))!)
            .ToList();
        userIds.Should().OnlyContain(id => id >= 1);
        userIds.Should().HaveCount(2);
    }

    [Test]
    public void Export_EveryBodyReferenceHasMatchingFootnoteContent()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());
        var docx = Export(content);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var refIds = doc.MainDocumentPart!.Document.Body!
            .Descendants<FootnoteReference>().Select(r => r.Id!.Value).ToList();
        var contentIds = doc.MainDocumentPart.FootnotesPart!.Footnotes!
            .Elements<WpFootnote>()
            .Where(IsUserFootnote)
            .Select(f => f.Id!.Value)
            .ToList();

        refIds.Should().HaveCount(2);
        refIds.Should().OnlyHaveUniqueItems();
        refIds.Should().BeSubsetOf(contentIds, "każde odwołanie musi mieć odpowiadającą treść (brak osieroconych)");
        contentIds.Should().BeEquivalentTo(refIds, "brak nieużywanych przypisów użytkownika");
    }

    [Test]
    public void Export_GeneratesSchemaValidDocx()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());
        var docx = Export(content);

        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(doc).ToList();

        errors.Should().BeEmpty(
            "wygenerowany dokument z przypisami musi przejść walidację OOXML: {0}",
            string.Join(" | ", errors.Select(e => e.Description)));
    }

    // ---------- 9.3 Round-trip ----------

    [Test]
    public void RoundTrip_PreservesFootnoteCountContentFormattingAndLinks()
    {
        var first = Import(FootnoteTestDocuments.TwoFootnotes());
        var exported = Export(first);
        var second = Import(exported);

        second.Footnotes.Should().HaveCount(first.Footnotes!.Count);

        Decoded(second.Footnotes![0].Html).Should().Contain("zażółć gęślą jaźń").And.Contain("€");
        second.Footnotes[1].Html.Should().Contain("Drugi akapit tego samego przypisu.");
        second.Footnotes[1].Html.Should().Contain("<strong>").And.Contain("<em>");

        // Reference order and links survive (display numbers recomputed identically).
        second.Html.Should().Contain("data-footnote-id=\"fn-1\"");
        second.Html.Should().Contain("data-footnote-id=\"fn-2\"");
        second.Html.IndexOf("fn-1", StringComparison.Ordinal)
            .Should().BeLessThan(second.Html.IndexOf("fn-2", StringComparison.Ordinal));
    }

    // ---------- 9.4 Editing operations ----------

    [Test]
    public void Edit_AddFootnote_AppearsInExportAndReimport()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        // Simulate the editor inserting a new footnote reference + content at the cursor.
        content.Html += "<p>Nowy akapit z odwołaniem" +
            "<sup class=\"footnote-ref\" data-footnote-id=\"fn-new\" aria-label=\"Przypis 3\">3</sup>.</p>";
        content.Footnotes!.Add(new DomainFootnote { Id = "fn-new", Html = "<p>Świeżo dodany przypis.</p>" });

        var reimported = Import(Export(content));

        reimported.Footnotes!.Should().HaveCount(3);
        reimported.Footnotes.Last().Html.Should().Contain("Świeżo dodany przypis.");
    }

    [Test]
    public void Edit_ChangeFootnoteContent_IsPersisted()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());
        content.Footnotes![0].Html = "<p>Zmieniona treść pierwszego przypisu.</p>";

        var reimported = Import(Export(content));

        reimported.Footnotes![0].Html.Should().Contain("Zmieniona treść pierwszego przypisu.");
        reimported.Footnotes[0].Html.Should().NotContain("zażółć gęślą jaźń");
    }

    [Test]
    public void Edit_DeleteFootnote_RemovesReferenceAndContent_NoOrphans()
    {
        var content = Import(FootnoteTestDocuments.TwoFootnotes());

        // Remove the second reference from the body and its content from the model.
        content.Html = System.Text.RegularExpressions.Regex.Replace(
            content.Html, "<sup class=\"footnote-ref\" data-footnote-id=\"fn-2\"[^>]*>.*?</sup>", "");
        content.Footnotes = content.Footnotes!.Where(f => f.Id != "fn-2").ToList();

        var docx = Export(content);
        var reimported = Import(docx);

        reimported.Footnotes.Should().ContainSingle();
        reimported.Html.Should().NotContain("fn-2");

        // No orphan references or unused footnotes in the exported package.
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        var refIds = doc.MainDocumentPart!.Document.Body!
            .Descendants<FootnoteReference>().Select(r => r.Id!.Value).ToList();
        var contentIds = doc.MainDocumentPart.FootnotesPart!.Footnotes!
            .Elements<WpFootnote>()
            .Where(IsUserFootnote)
            .Select(f => f.Id!.Value).ToList();
        contentIds.Should().BeEquivalentTo(refIds);
    }

    [Test]
    public void Edit_ReorderReferences_RecomputesDisplayNumbers()
    {
        // Two footnotes, but the body references fn-2 BEFORE fn-1.
        var footnotes = new List<DomainFootnote>
        {
            new() { Id = "fn-a", Html = "<p>Przypis A.</p>" },
            new() { Id = "fn-b", Html = "<p>Przypis B.</p>" }
        };
        var html =
            "<p>Najpierw B<sup class=\"footnote-ref\" data-footnote-id=\"fn-b\">1</sup>, " +
            "potem A<sup class=\"footnote-ref\" data-footnote-id=\"fn-a\">2</sup>.</p>";

        var docx = _writer.Convert(html, footnotes: footnotes);
        var reimported = Import(docx);

        // Internal ids are re-derived from OOXML ids on import (they need not survive the round-trip);
        // what MUST hold is that the reference appearing FIRST is numbered 1 and its content is B.
        reimported.Footnotes![0].Html.Should().Contain("Przypis B.", "pierwsze odwołanie w treści dotyczy B");
        reimported.Footnotes[1].Html.Should().Contain("Przypis A.");

        var firstId = reimported.Footnotes[0].Id;
        var secondId = reimported.Footnotes[1].Id;
        reimported.Html.Should().Contain($"data-footnote-id=\"{firstId}\" aria-label=\"Przypis 1\">1</sup>");
        reimported.Html.Should().Contain($"data-footnote-id=\"{secondId}\" aria-label=\"Przypis 2\">2</sup>");
    }

    // ---------- 9.5 Regression: no footnotes ----------

    [Test]
    public void NoFootnotes_ImportProducesNoFootnoteModel()
    {
        var content = Import(FootnoteTestDocuments.NoFootnotes());

        content.Footnotes.Should().BeNull();
        content.Html.Should().NotContain("footnote-ref");
    }

    [Test]
    public void NoFootnotes_ExportDoesNotCreateFootnotesPart()
    {
        var content = Import(FootnoteTestDocuments.NoFootnotes());
        var docx = Export(content);

        using var archive = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        archive.GetEntry("word/footnotes.xml").Should().BeNull(
            "dokument bez przypisów nie dostaje nadmiarowej części footnotes.xml");
    }

    [Test]
    public void NoFootnotes_RoundTripStaysClean()
    {
        var content = Import(FootnoteTestDocuments.NoFootnotes());
        var second = Import(Export(content));

        second.Footnotes.Should().BeNull();
        Decoded(second.Html).Should().Contain("Dokument bez przypisów.");
    }
}

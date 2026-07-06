using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Golden;

/// <summary>
/// Etap 0 regression harness. Each golden document is converted and its HTML is
/// compared to an approved snapshot, plus a few human-readable point assertions on
/// the unit-conversion outputs (font-size pt, indent px, table widths px, image px).
/// These lock current behaviour ahead of the Etap 2+ pipeline refactor.
/// </summary>
[TestFixture]
public class GoldenSnapshotTests
{
    private DocxToHtmlConverter _converter = null!;

    [SetUp]
    public void Setup() => _converter = new DocxToHtmlConverter();

    [Test]
    public void SimpleParagraphs_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.SimpleParagraphs());
        content.Html.Should().Contain("text-align:center");
        HtmlSnapshot.Verify(content.Html, "simple-paragraphs");
    }

    [Test]
    public void StyledRuns_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.StyledRuns());

        content.Html.Should().Contain("<strong>pogrubiony");
        content.Html.Should().Contain("<em>kursywa");
        content.Html.Should().Contain("<u>podkreślony");
        content.Html.Should().Contain("color:#FF0000");
        content.Html.Should().Contain("font-size:16pt"); // sz=32 half-points → 16 pt
        HtmlSnapshot.Verify(content.Html, "styled-runs");
    }

    [Test]
    public void CharacterStyleRun_AppliesNamedRunStyle()
    {
        var content = _converter.Convert(GoldenDocuments.CharacterStyleRun());

        // All three properties come from the "Akcent" character style via w:rStyle.
        content.Html.Should().Contain("font-weight:bold");
        content.Html.Should().Contain("color:#C00000");
        content.Html.Should().Contain("font-size:14pt"); // sz=28 half-points → 14 pt
        HtmlSnapshot.Verify(content.Html, "character-style-run");
    }

    [Test]
    public void ParagraphSpacingAndIndent_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.ParagraphSpacingAndIndent());

        content.Html.Should().Contain("margin-top:12pt");     // 240 twips → 12 pt
        content.Html.Should().Contain("margin-bottom:6pt");   // 120 twips → 6 pt
        content.Html.Should().Contain("margin-left:48px");    // 720 twips → 48 px
        content.Html.Should().Contain("text-indent:32px");    // 480 twips → 32 px
        HtmlSnapshot.Verify(content.Html, "paragraph-spacing-indent");
    }

    [Test]
    public void TabStopLeftCenterRight_RendersPositionedSegments()
    {
        var content = _converter.Convert(GoldenDocuments.TabStopLeftCenterRight());

        // Segments are positioned at the real stops (center 4536 tw = 302 px, right 9072 tw = 604 px)
        // rather than spread evenly by a flex row that ignores the declared geometry.
        content.Html.Should().Contain("docx-tab-seg");
        content.Html.Should().Contain("data-tab-align=\"center\"").And.Contain("translateX(-50%)");
        content.Html.Should().Contain("data-tab-align=\"right\"").And.Contain("translateX(-100%)");
        content.Html.Should().NotContain("display:flex");
        content.Html.Should().Contain("Lewy");
        content.Html.Should().Contain("Środek");
        content.Html.Should().Contain("Prawy");
        HtmlSnapshot.Verify(content.Html, "tab-stops-lcr");
    }

    [Test]
    public void SimpleTable_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.SimpleTableWithBordersAndWidths());

        content.Html.Should().Contain("<table");
        content.Html.Should().Contain("table-layout:fixed");          // explicit width → fixed layout
        content.Html.Should().Contain("<colgroup>");                  // tblGrid → authoritative columns
        content.Html.Should().Contain("<col style=\"width:200px;\""); // 3000 twips → 200 px
        content.Html.Should().Contain("<col style=\"width:133px;\""); // 2000 twips → 133 px (truncated)
        HtmlSnapshot.Verify(content.Html, "simple-table");
    }

    [Test]
    public void MergedCellsTable_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.MergedCellsTable());

        content.Html.Should().Contain("table-layout:fixed");      // tblLayout=fixed
        content.Html.Should().Contain("width:399px");             // grid sum 3×133 px (truncated)
        content.Html.Should().Contain("colspan=\"2\"");           // gridSpan → colspan
        content.Html.Should().Contain("rowspan=\"2\"");           // vMerge restart+continue → rowspan
        HtmlSnapshot.Verify(content.Html, "merged-cells-table");
    }

    [Test]
    public void HeaderFooterWithImage_MatchesSnapshot()
    {
        var content = _converter.Convert(GoldenDocuments.HeaderFooterWithImage());

        content.Header.Should().NotBeNull();
        content.Footer.Should().NotBeNull();
        // 1270000 EMU → ~133 px, 317500 EMU → ~33 px (aspect preserved).
        content.Header!.Html.Should().Contain("data-width-emu=\"1270000\"");
        content.Header.Html.Should().Contain("data-height-emu=\"317500\"");

        HtmlSnapshot.Verify(content.Header.Html, "headerfooter-header");
        HtmlSnapshot.Verify(content.Footer!.Html, "headerfooter-footer");
    }
}

using System.IO.Compression;
using System.Text.RegularExpressions;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Fast package-level fidelity gate (CI). Complements the per-feature round-trip tests by checking
/// the whole-package invariants after a round-trip: styles/theme preserved (pass-through, R-16),
/// manual page breaks kept (R-15), a logical table not multiplied (R-17), no fabricated row
/// heights (R-18). Uses a parser-independent ZIP/regex reader so it judges the writer objectively.
/// </summary>
[TestFixture]
public class DocxFidelityRegressionTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    [Test]
    public void PassThrough_PreservesStylesAndTheme_R16()
    {
        var original = BuildRich(paragraphStyles: 30, tableStyles: 50, minorFont: "Cambria");
        var before = Analyze(original);
        before.Styles.Should().BeGreaterThan(60);

        var content = _reader.Convert(new MemoryStream(original));
        var result = _writer.ConvertPreservingPackage(content.Html, new MemoryStream(original));
        var after = Analyze(result);

        after.Styles.Should().BeGreaterThanOrEqualTo(before.Styles);       // not collapsed
        after.TableStyles.Should().Be(before.TableStyles);
        after.ThemeMinorFont.Should().Be("Cambria");                       // font not swapped
    }

    [Test]
    public void RoundTrip_KeepsManualPageBreaks_R15()
    {
        var original = BuildWithPageBreaks(2);
        Analyze(original).PageBreaks.Should().Be(2);

        var content = _reader.Convert(new MemoryStream(original));
        var result = _writer.Convert(content.Html);

        Analyze(result).PageBreaks.Should().Be(2); // not lost, not duplicated
    }

    [Test]
    public void RoundTrip_DoesNotMultiplyTables_R17()
    {
        var original = BuildWithTables(3, rows: 4);
        var before = Analyze(original);

        var content = _reader.Convert(new MemoryStream(original));
        var result = _writer.Convert(content.Html);
        var after = Analyze(result);

        after.Tables.Should().BeLessThanOrEqualTo(before.Tables); // one logical table stays one
        after.Tables.Should().Be(3);
    }

    [Test]
    public void RoundTrip_DoesNotFabricateRowHeights_R18()
    {
        var original = BuildWithTables(2, rows: 5); // no explicit trHeight
        Analyze(original).RowHeights.Should().Be(0);

        var content = _reader.Convert(new MemoryStream(original));
        var result = _writer.Convert(content.Html);

        Analyze(result).RowHeights.Should().Be(0); // writer must not invent row heights
    }

    // --- tiny parser-independent package reader -------------------------------

    private sealed record Report(int Styles, int TableStyles, string? ThemeMinorFont, int Tables, int PageBreaks, int RowHeights);

    private static Report Analyze(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        string Read(Predicate<string> m)
        {
            var e = zip.Entries.FirstOrDefault(x => m(x.FullName));
            if (e == null) return string.Empty;
            using var r = new StreamReader(e.Open());
            return r.ReadToEnd();
        }
        var styles = Read(n => Regex.IsMatch(n, @"word/styles\d*\.xml$"));
        var doc = Read(n => n.EndsWith("word/document.xml"));
        var theme = Read(n => Regex.IsMatch(n, @"word/theme/theme\d*\.xml$"));
        var minor = Regex.Match(theme, "<a:minorFont>.*?<a:latin typeface=\"([^\"]*)\"", RegexOptions.Singleline);
        return new Report(
            Regex.Matches(styles, "<w:style ").Count,
            Regex.Matches(styles, "<w:style [^>]*w:type=\"table\"").Count,
            minor.Success ? minor.Groups[1].Value : null,
            Regex.Matches(doc, "<w:tbl>").Count,
            Regex.Matches(doc, "<w:br w:type=\"page\"").Count,
            Regex.Matches(doc, "<w:trHeight ").Count);
    }

    // --- builders -------------------------------------------------------------

    private static byte[] BuildRich(int paragraphStyles, int tableStyles, string minorFont)
    {
        return Build(
            body => body.Append(new Paragraph(new Run(new Text("Treść")))),
            styles =>
            {
                for (var i = 0; i < paragraphStyles; i++)
                    styles.Append(new Style(new StyleName { Val = $"P{i}" }) { Type = StyleValues.Paragraph, StyleId = $"P{i}" });
                for (var i = 0; i < tableStyles; i++)
                    styles.Append(new Style(new StyleName { Val = $"T{i}" }) { Type = StyleValues.Table, StyleId = $"T{i}" });
            },
            minorFont);
    }

    private static byte[] BuildWithPageBreaks(int count) => Build(body =>
    {
        for (var i = 0; i <= count; i++)
        {
            body.Append(new Paragraph(new Run(new Text($"Sekcja {i}"))));
            if (i < count) body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }
    });

    private static byte[] BuildWithTables(int tables, int rows) => Build(body =>
    {
        for (var t = 0; t < tables; t++)
        {
            var table = new Table(new TableProperties(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }),
                new TableGrid(new GridColumn { Width = "5000" }, new GridColumn { Width = "5000" }));
            for (var r = 0; r < rows; r++)
                table.Append(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text($"R{r}A")))),
                    new TableCell(new Paragraph(new Run(new Text($"R{r}B"))))));
            body.Append(table);
            body.Append(new Paragraph());
        }
    });

    private static byte[] Build(Action<Body> fill, Action<Styles>? configureStyles = null, string minorFont = "Cambria")
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi },
                new FontSize { Val = "22" }))));
            styles.Append(new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            configureStyles?.Invoke(styles);
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();

            mainPart.AddNewPart<ThemePart>().FeedXml(ThemeXml(minorFont));

            fill(mainPart.Document.Body!);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    private static string ThemeXml(string minorLatin) =>
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\">" +
        "<a:themeElements><a:clrScheme name=\"C\">" +
        string.Concat(new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" }
            .Select(c => $"<a:{c}><a:srgbClr val=\"000000\"/></a:{c}>")) +
        "</a:clrScheme><a:fontScheme name=\"F\">" +
        "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        $"<a:minorFont><a:latin typeface=\"{minorLatin}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
        "<a:fmtScheme name=\"S\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>" +
        "<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>" +
        "</a:themeElements></a:theme>";
}

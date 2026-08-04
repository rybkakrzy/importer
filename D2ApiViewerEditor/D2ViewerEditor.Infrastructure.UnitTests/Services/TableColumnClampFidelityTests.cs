using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Word DOSKALOWUJE tabelę szerszą niż szpalta (sekcja wielokolumnowa) do jej szerokości —
/// siatka tblGrid w pliku zostaje oryginalna, kompresja jest czysto renderowa. Reader musi
/// robić to samo: skalować px podglądu (colgroup style + width tabeli), zachowując
/// data-w-tw z ORYGINALNYMI twipami (zapis nie utrwala kompresji). Repro: tabela wzoru
/// Rachunku Wirtualnego w dwukolumnowej umowie deweloperskiej wyjeżdżała poza szpaltę.
/// </summary>
[TestFixture]
public class TableColumnClampFidelityTests
{
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup() => _reader = new DocxToHtmlConverter();

    // Tabela 2×4000tw (533px) — szersza niż szpalta 2-kolumnowej A4 (~4182tw ≈ 278px).
    private const string WideTable = @"<w:tbl>
  <w:tblPr><w:tblW w:w=""8000"" w:type=""dxa""/><w:tblLayout w:type=""fixed""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""4000""/><w:gridCol w:w=""4000""/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcW w:w=""4000"" w:type=""dxa""/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w=""4000"" w:type=""dxa""/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>";

    private static MemoryStream DocxWithBody(string bodyInnerXml, string sectPrXml)
    {
        var doc = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>{bodyInnerXml}{sectPrXml}</w:body>
</w:document>";
        var ms = new MemoryStream();
        using (var docx = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = docx.AddMainDocumentPart();
            using var w = new StreamWriter(mainPart.GetStream(FileMode.Create), Encoding.UTF8);
            w.Write(doc);
        }
        ms.Position = 0;
        return ms;
    }

    private const string TwoColumnSectPr = @"<w:sectPr>
  <w:pgSz w:w=""11906"" w:h=""16838""/>
  <w:pgMar w:top=""1417"" w:right=""1417"" w:bottom=""1417"" w:left=""1417"" w:header=""708"" w:footer=""708""/>
  <w:cols w:num=""2"" w:space=""708""/>
</w:sectPr>";

    private const string SingleColumnSectPr = @"<w:sectPr>
  <w:pgSz w:w=""11906"" w:h=""16838""/>
  <w:pgMar w:top=""1417"" w:right=""1417"" w:bottom=""1417"" w:left=""1417"" w:header=""708"" w:footer=""708""/>
</w:sectPr>";

    [Test]
    public void Read_WideTableInTwoColumnSection_ScalesPreviewToColumnWidth()
    {
        var html = _reader.Convert(DocxWithBody(WideTable, TwoColumnSectPr)).Html;

        // Szpalta: (11906−2·1417−708)/2 = 4182tw ≈ 278px. Suma px colgroup ≤ szpalta.
        var colPx = System.Text.RegularExpressions.Regex.Matches(html, "<col style=\"width:(\\d+)px;\"")
            .Select(m => int.Parse(m.Groups[1].Value)).ToList();
        colPx.Should().HaveCount(2);
        colPx.Sum().Should().BeLessThanOrEqualTo(280, "tabela szersza niż szpalta musi być doskalowana");
        colPx.Sum().Should().BeGreaterThan(240, "skala proporcjonalna, nie zerowanie");
        // Kompresja jest TYLKO renderowa — oryginalne twipy w data-w-tw (round-trip zapisu).
        html.Should().Contain("data-w-tw=\"4000\"");
        // Szerokość tabeli też doskalowana (nie 533px z tblW).
        html.Should().NotContain("width:533px");
    }

    [Test]
    public void Read_WideTableInSingleColumnSection_KeepsOriginalWidths()
    {
        var html = _reader.Convert(DocxWithBody(WideTable, SingleColumnSectPr)).Html;

        // 8000tw = 533px < obszar treści 9072tw — bez clampu, px 1:1 z twipów.
        var colPx = System.Text.RegularExpressions.Regex.Matches(html, "<col style=\"width:(\\d+)px;\"")
            .Select(m => int.Parse(m.Groups[1].Value)).ToList();
        colPx.Should().HaveCount(2);
        colPx.Sum().Should().BeInRange(530, 536);
    }

    [Test]
    public void Read_TableWiderThanPage_ScalesToContentWidth()
    {
        var hugeTable = WideTable.Replace("4000", "6000").Replace("8000", "12000");
        var html = _reader.Convert(DocxWithBody(hugeTable, SingleColumnSectPr)).Html;

        // 12000tw = 800px > obszar treści 9072tw ≈ 604px → doskalowanie.
        var colPx = System.Text.RegularExpressions.Regex.Matches(html, "<col style=\"width:(\\d+)px;\"")
            .Select(m => int.Parse(m.Groups[1].Value)).ToList();
        colPx.Sum().Should().BeLessThanOrEqualTo(606);
        html.Should().Contain("data-w-tw=\"6000\"");
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Direct w:ind akapitu listowego (Word: bezpośrednie wcięcie akapitu NADPISUJE wcięcie
/// z definicji poziomu numeracji).
///
/// Poprzednio StripIndentationCss kasował wcięcia z &lt;li&gt; bez śladu — punktory list
/// (zwłaszcza w komórkach tabel, gdzie dokumenty klienta nadpisują wcięcia na małe wartości)
/// stały za daleko od tekstu, a po zapisie direct w:ind ginął całkowicie.
/// Teraz: delta wizualna (margin-left / --ind-hanging na li) + kontrakt data-ind-*-tw,
/// z którego writer odtwarza w:ind.
/// </summary>
[TestFixture]
public class ListDirectIndentFidelityTests
{
    private DocxToHtmlConverter _reader = null!;
    private HtmlToDocxConverter _writer = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
        _writer = new HtmlToDocxConverter();
    }

    private static MemoryStream BuildDocx(Indentation? directInd, bool inTableCell)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var level = new Level { LevelIndex = 0 };
            level.Append(new StartNumberingValue { Val = 1 });
            level.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
            level.Append(new LevelText { Val = "•" });
            level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            level.Append(new PreviousParagraphProperties(
                new Indentation { Left = "720", Hanging = "360" }));

            var abstractNum = new AbstractNum { AbstractNumberId = 1 };
            abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
            abstractNum.Append(level);

            var num = new NumberingInstance { NumberID = 1 };
            num.Append(new AbstractNumId { Val = 1 });

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(abstractNum, num);
            numberingPart.Numbering.Save();

            var paraProps = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            if (directInd != null)
                paraProps.Append(directInd);
            var listPara = new Paragraph(paraProps, new Run(new Text("Punkt")));

            var body = mainPart.Document.Body!;
            if (inTableCell)
            {
                var table = new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn { Width = "5000" }),
                    new TableRow(new TableCell(
                        new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                        listPara)));
                body.Append(table);
            }
            else
            {
                body.Append(listPara);
            }

            body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static Paragraph FirstListParagraph(byte[] docx)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return (Paragraph)doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .First(p => p.ParagraphProperties?.GetFirstChild<NumberingProperties>() != null)
            .CloneNode(true);
    }

    [Test]
    public void Read_DirectIndInTableCell_OverridesLevelIndentation()
    {
        // Definicja poziomu: left=720 (48 px), hanging=360 (24 px).
        // Akapit w komórce nadpisuje na left=284 (18 px), hanging=142 (9 px).
        var html = _reader.Convert(BuildDocx(new Indentation { Left = "284", Hanging = "142" }, inTableCell: true)).Html;

        var li = System.Text.RegularExpressions.Regex.Match(html, "<li[^>]*>").Value;
        li.Should().Contain("data-ind-left-tw=\"284\"");
        li.Should().Contain("data-ind-hanging-tw=\"142\"");
        // Delta wizualna względem paddingu kontenera (48 px z poziomu): 18 − 48 = −30 px.
        li.Should().Contain("margin-left:-30px");
        li.Should().Contain("--ind-hanging:9px");
    }

    [Test]
    public void Read_NoDirectInd_KeepsLevelIndentationWithoutItemOverrides()
    {
        var html = _reader.Convert(BuildDocx(null, inTableCell: true)).Html;

        var li = System.Text.RegularExpressions.Regex.Match(html, "<li[^>]*>").Value;
        li.Should().NotContain("data-ind-left-tw");
        li.Should().NotContain("margin-left:");
        // Kontener niesie wcięcie poziomu.
        System.Text.RegularExpressions.Regex.Match(html, "<ul[^>]*>").Value
            .Should().Contain("padding-left:48px").And.Contain("--ind-hanging:24px");
    }

    [Test]
    public void RoundTrip_DirectInd_IsRestoredOnListParagraph()
    {
        var html = _reader.Convert(BuildDocx(new Indentation { Left = "284", Hanging = "142" }, inTableCell: true)).Html;

        var para = FirstListParagraph(_writer.Convert(html));
        var ind = para.ParagraphProperties!.GetFirstChild<Indentation>();
        ind.Should().NotBeNull("direct w:ind nie może ginąć przy zapisie");
        ind!.Left!.Value.Should().Be("284");
        ind.Hanging!.Value.Should().Be("142");
    }

    [Test]
    public void RoundTrip_BodyList_DirectIndSurvives()
    {
        var html = _reader.Convert(BuildDocx(new Indentation { Left = "1418", Hanging = "360" }, inTableCell: false)).Html;

        var para = FirstListParagraph(_writer.Convert(html));
        var ind = para.ParagraphProperties!.GetFirstChild<Indentation>();
        ind.Should().NotBeNull();
        ind!.Left!.Value.Should().Be("1418");
        ind.Hanging!.Value.Should().Be("360");
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Infrastructure.UnitTests.Golden;

/// <summary>
/// Deterministic in-memory DOCX builders used as golden samples for the regression
/// harness. Each document targets a specific conversion area (geometry, runs, spacing,
/// tables, header/footer images) so a snapshot mismatch points at the regressed feature.
/// No randomness, no date fields, no metafile images — output is byte-stable across runs.
/// </summary>
public static class GoldenDocuments
{
    // 1x1 red PNG — cheapest legal image that round-trips deterministically.
    private static readonly byte[] OnePixelPng = System.Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    /// <summary>Plain paragraphs — baseline text + alignment, no direct formatting.</summary>
    public static MemoryStream SimpleParagraphs()
    {
        return Build(body =>
        {
            body.Append(Paragraph(null, Run(null, "Pierwszy akapit.")));
            body.Append(Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                Run(null, "Wyśrodkowany akapit.")));
        });
    }

    /// <summary>Direct character formatting — bold, italic, underline, colour, explicit size.</summary>
    public static MemoryStream StyledRuns()
    {
        return Build(body =>
        {
            body.Append(Paragraph(null,
                Run(new RunProperties(new Bold()), "pogrubiony "),
                Run(new RunProperties(new Italic()), "kursywa "),
                Run(new RunProperties(new Underline { Val = UnderlineValues.Single }), "podkreślony "),
                Run(new RunProperties(new Color { Val = "FF0000" }), "czerwony "),
                Run(new RunProperties(new FontSize { Val = "32" }), "16pt")));
        });
    }

    /// <summary>Paragraph spacing + indentation — exercises twips→pt and twips→px paths.</summary>
    public static MemoryStream ParagraphSpacingAndIndent()
    {
        return Build(body =>
        {
            var pPr = new ParagraphProperties(
                new SpacingBetweenLines { Before = "240", After = "120", Line = "360", LineRule = LineSpacingRuleValues.Auto },
                new Indentation { Left = "720", Right = "360", FirstLine = "480" });
            body.Append(Paragraph(pPr, Run(null, "Akapit z odstępami i wcięciami.")));
        });
    }

    /// <summary>
    /// The classic header/footer "left ⇥ center ⇥ right" single line built with a center
    /// and a right tab stop — exercises Etap 7 flex tab-stop rendering.
    /// </summary>
    public static MemoryStream TabStopLeftCenterRight()
    {
        return Build(body =>
        {
            var pPr = new ParagraphProperties(
                new Tabs(
                    new TabStop { Val = TabStopValues.Center, Position = 4536 },
                    new TabStop { Val = TabStopValues.Right, Position = 9072 }));
            body.Append(new Paragraph(pPr,
                new Run(new Text("Lewy")),
                new Run(new TabChar()),
                new Run(new Text("Środek")),
                new Run(new TabChar()),
                new Run(new Text("Prawy"))));
        });
    }

    /// <summary>2x2 table with an explicit grid, preferred cell widths and borders.</summary>
    public static MemoryStream SimpleTableWithBordersAndWidths()
    {
        return Build(body =>
        {
            var borders = new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" });

            var table = new Table(
                new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa },
                    borders),
                new TableGrid(
                    new GridColumn { Width = "3000" },
                    new GridColumn { Width = "2000" }),
                TableRow(("A1", 3000), ("B1", 2000)),
                TableRow(("A2", 3000), ("B2", 2000)));
            body.Append(table);
        });
    }

    /// <summary>
    /// A run whose only formatting comes from a named character style (w:rStyle) plus a
    /// second plain run. Exercises Etap 3 character-style resolution with inheritance.
    /// </summary>
    public static MemoryStream CharacterStyleRun()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                    new FontSize { Val = "22" }))),
                new Style(
                    new StyleName { Val = "Akcent" },
                    new StyleRunProperties(
                        new Bold(),
                        new Color { Val = "C00000" },
                        new FontSize { Val = "28" }))
                {
                    Type = StyleValues.Character,
                    StyleId = "Akcent"
                });
            stylesPart.Styles.Save();

            var body = mainPart.Document.Body!;
            body.Append(new Paragraph(
                new Run(new RunProperties(new RunStyle { Val = "Akcent" }), new Text("Akcentowany")),
                new Run(new Text(" zwykły") { Space = SpaceProcessingModeValues.Preserve })));
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Fixed-layout table driven only by tblGrid (no explicit table width) with a
    /// horizontally merged cell (gridSpan) and a vertically merged cell (vMerge).
    /// Exercises colgroup emission, the grid-sum width fallback and colspan/rowspan.
    /// </summary>
    public static MemoryStream MergedCellsTable()
    {
        return Build(body =>
        {
            var table = new Table(
                new TableProperties(new TableLayout { Type = TableLayoutValues.Fixed }),
                new TableGrid(
                    new GridColumn { Width = "2000" },
                    new GridColumn { Width = "2000" },
                    new GridColumn { Width = "2000" }),
                // Row 1: a cell spanning the first two columns + a cell starting a vertical merge.
                new TableRow(
                    new TableCell(
                        new TableCellProperties(new GridSpan { Val = 2 }),
                        new Paragraph(new Run(new Text("Scalone poziomo")))),
                    new TableCell(
                        new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                        new Paragraph(new Run(new Text("Scalone pionowo"))))),
                // Row 2: two normal cells + the continuation of the vertical merge.
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("A2")))),
                    new TableCell(new Paragraph(new Run(new Text("B2")))),
                    new TableCell(
                        new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Continue }),
                        new Paragraph(new Run(new Text(string.Empty))))));
            body.Append(table);
        });
    }

    /// <summary>Header with an inline logo (EMU sized) + footer with L/C/R aligned paragraphs.</summary>
    public static MemoryStream HeaderFooterWithImage()
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            AppendDefaultStyles(mainPart);

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var imagePart = headerPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedBytes(OnePixelPng);
            var relId = headerPart.GetIdOfPart(imagePart);
            headerPart.Header = new Header(new Paragraph(InlineImageRun(relId, 1270000, 317500)));
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }), Run(null, "L")),
                Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }), Run(null, "C")),
                Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Right }), Run(null, "R")));
            footerPart.Footer.Save();

            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Bottom = 1417, Left = 1417, Right = 1417, Header = 708, Footer = 708, Gutter = 0 });
            body.Append(sectPr);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    // --- builder helpers -----------------------------------------------------

    private static MemoryStream Build(Action<Body> fill)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            AppendDefaultStyles(mainPart);
            fill(mainPart.Document.Body!);
            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static void AppendDefaultStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
            new RunPropertiesBaseStyle(
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                new FontSize { Val = "22" }))));
        stylesPart.Styles.Save();
    }

    private static Paragraph Paragraph(ParagraphProperties? pPr, params Run[] runs)
    {
        var p = new Paragraph();
        if (pPr != null) p.Append(pPr);
        foreach (var r in runs) p.Append(r);
        return p;
    }

    private static Run Run(RunProperties? rPr, string text)
    {
        var r = new Run();
        if (rPr != null) r.Append(rPr);
        r.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return r;
    }

    private static TableRow TableRow(params (string text, int widthTwips)[] cells)
    {
        var row = new TableRow();
        foreach (var (text, widthTwips) in cells)
        {
            row.Append(new TableCell(
                new TableCellProperties(new TableCellWidth { Width = widthTwips.ToString(), Type = TableWidthUnitValues.Dxa }),
                new Paragraph(new Run(new Text(text)))));
        }
        return row;
    }

    private static Run InlineImageRun(string relId, long cx, long cy)
    {
        var xml = $@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:drawing>
    <wp:inline xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" distT=""0"" distB=""0"" distL=""0"" distR=""0"">
      <wp:extent cx=""{cx}"" cy=""{cy}""/>
      <wp:docPr id=""1"" name=""Picture 1""/>
      <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
        <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
          <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
            <pic:nvPicPr><pic:cNvPr id=""1"" name=""Picture 1""/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill>
              <a:blip xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:embed=""{relId}""/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{cx}"" cy=""{cy}""/></a:xfrm>
              <a:prstGeom prst=""rect""><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>";
        return new Run(xml);
    }
}

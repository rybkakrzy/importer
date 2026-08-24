using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Benchmarks;

/// <summary>
/// Deterministic, self-contained DOCX samples for the benchmarks. Synthetic documents are
/// generated in memory so the suite runs on any environment with no committed (and possibly
/// sensitive) files. A real regression document can additionally be supplied via the
/// D2_BENCH_ASSETS directory (env var) or the repo default — used only when present.
/// </summary>
public static class BenchmarkAssets
{
    // 1x1 red PNG — cheapest legal image that round-trips deterministically.
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    public const string RegressionOriginalFileName = "orginał_GOOD.docx";

    private const string DefaultAssetsDirectory = @"sciezka_do_projektu";

    /// <summary>Directory holding optional real regression files. Override with D2_BENCH_ASSETS.</summary>
    public static string AssetsDirectory
    {
        get
        {
            var configured = Environment.GetEnvironmentVariable("D2_BENCH_ASSETS");
            var directory = string.IsNullOrWhiteSpace(configured) ? DefaultAssetsDirectory : configured;
            return Path.TrimEndingDirectorySeparator(Path.GetFullPath(directory));
        }
    }

    public static byte[]? TryLoadRegressionOriginal()
    {
        // The asset name is a fixed constant, but the base directory comes from the environment.
        // Resolve both to canonical paths and read only when the result stays inside that directory.
        var directory = AssetsDirectory;
        var path = Path.GetFullPath(Path.Combine(directory, Path.GetFileName(RegressionOriginalFileName)));

        if (!IsInside(directory, path) || !File.Exists(path))
            return null;

        return File.ReadAllBytes(path);
    }

    private static bool IsInside(string directory, string candidate) =>
        candidate.StartsWith(directory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    public static byte[] Simple() => Build(body =>
    {
        for (var i = 0; i < 20; i++)
            body.Append(Paragraph($"Akapit numer {i} — przykładowa treść dokumentu."));
    });

    public static byte[] MultiStyle(int paragraphStyles = 60, int tableStyles = 100) => Build(
        body =>
        {
            for (var i = 0; i < 40; i++)
                body.Append(Paragraph($"Treść akapitu {i} ze stylem P{i % paragraphStyles}.", styleId: $"P{i % paragraphStyles}"));
        },
        configureStyles: styles =>
        {
            for (var i = 0; i < paragraphStyles; i++)
                styles.Append(new Style(new StyleName { Val = $"P{i}" }) { Type = StyleValues.Paragraph, StyleId = $"P{i}" });
            for (var i = 0; i < tableStyles; i++)
                styles.Append(new Style(new StyleName { Val = $"T{i}" }) { Type = StyleValues.Table, StyleId = $"T{i}" });
        });

    public static byte[] Tables(int tables = 6, int rowsEach = 8) => Build(body =>
    {
        for (var t = 0; t < tables; t++)
        {
            body.Append(Paragraph($"Tabela {t}:"));
            body.Append(BuildTable(rowsEach, columns: 3));
            body.Append(new Paragraph());
        }
    });

    public static byte[] LargeTable(int rows = 400) => Build(body =>
    {
        body.Append(Paragraph("Duża tabela:"));
        body.Append(BuildTable(rows, columns: 4));
        body.Append(new Paragraph());
    });

    public static byte[] Images(int count = 25) => Build((doc, mainPart, body) =>
    {
        for (var i = 0; i < count; i++)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            imagePart.FeedBytes(OnePixelPng);
            var relId = mainPart.GetIdOfPart(imagePart);
            body.Append(new Paragraph(InlineImageRun(relId, 1270000, 317500)));
        }
    });

    public static byte[] HeaderFooter() => Build((doc, mainPart, body) =>
    {
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var imagePart = headerPart.AddImagePart(ImagePartType.Png);
        imagePart.FeedBytes(OnePixelPng);
        headerPart.Header = new Header(new Paragraph(InlineImageRun(headerPart.GetIdOfPart(imagePart), 1270000, 317500)));
        headerPart.Header.Save();

        var footerPart = mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }), new Run(new Text("L"))),
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }), new Run(new Text("C"))),
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Right }), new Run(new Text("R"))));
        footerPart.Footer.Save();

        for (var i = 0; i < 15; i++) body.Append(Paragraph($"Treść {i}."));

        body.Append(new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
            new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin { Top = 567, Bottom = 851, Left = 851, Right = 851, Header = 6, Footer = 340, Gutter = 0 }));
    });

    public static byte[] PageBreaks(int count = 3) => Build(body =>
    {
        for (var i = 0; i < count + 1; i++)
        {
            body.Append(Paragraph($"Sekcja {i} przed podziałem strony."));
            if (i < count)
                body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }
    });

    // --- builders -----------------------------------------------------------

    private static byte[] Build(Action<Body> fill, Action<Styles>? configureStyles = null) =>
        Build((doc, mainPart, body) => fill(body), configureStyles);

    private static byte[] Build(Action<WordprocessingDocument, MainDocumentPart, Body> fill, Action<Styles>? configureStyles = null)
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

            mainPart.AddNewPart<ThemePart>().FeedXml(MinimalThemeXml());

            fill(doc, mainPart, mainPart.Document.Body!);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    private static Paragraph Paragraph(string text, string? styleId = null)
    {
        var p = new Paragraph();
        if (styleId != null) p.Append(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
        p.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return p;
    }

    private static Table BuildTable(int rows, int columns)
    {
        var grid = new TableGrid();
        for (var c = 0; c < columns; c++) grid.Append(new GridColumn { Width = (10000 / columns).ToString() });

        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "808080" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "808080" })),
            grid);

        for (var r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (var c = 0; c < columns; c++)
                row.Append(new TableCell(
                    new TableCellProperties(new TableCellWidth { Width = (10000 / columns).ToString(), Type = TableWidthUnitValues.Dxa }),
                    new Paragraph(new Run(new Text($"R{r}C{c}")))));
            table.Append(row);
        }
        return table;
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

    private static string MinimalThemeXml() =>
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
        "<a:themeElements><a:clrScheme name=\"Office\">" +
        string.Concat(new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" }
            .Select(c => $"<a:{c}><a:srgbClr val=\"000000\"/></a:{c}>")) +
        "</a:clrScheme>" +
        "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        "<a:minorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
        "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>" +
        "<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>" +
        "</a:themeElements></a:theme>";
}

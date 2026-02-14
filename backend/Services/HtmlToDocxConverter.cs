using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImporterApi.Models;
using HtmlAgilityPack;

namespace ImporterApi.Services;

/// <summary>
/// Serwis do konwersji HTML na dokument DOCX
/// Własna implementacja generatora OpenXML
/// </summary>
public class HtmlToDocxConverter
{
    private MainDocumentPart? _mainPart;
    private readonly Dictionary<string, string> _imageRelationships = new();
    private int _imageCounter = 0;

    /// <summary>
    /// Konwertuje HTML na plik DOCX
    /// </summary>
    public byte[] Convert(string html, DocumentMetadata? metadata = null)
    {
        using var memoryStream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
        {
            _mainPart = document.AddMainDocumentPart();
            _mainPart.Document = new Document();
            
            var body = new Body();
            _mainPart.Document.Body = body;

            // Dodaj style dokumentu
            AddDocumentStyles(document);

            // Parsuj HTML i konwertuj na elementy Word
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            
            ConvertHtmlToBody(htmlDoc.DocumentNode, body);

            // Ustaw metadane
            if (metadata != null)
            {
                SetDocumentMetadata(document, metadata);
            }

            // Dodaj ustawienia strony
            AddPageSettings(body);

            document.Save();
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Dodaje domyślne style do dokumentu
    /// </summary>
    private void AddDocumentStyles(WordprocessingDocument document)
    {
        var stylesPart = _mainPart!.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        // Styl normalny
        var normalStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        normalStyle.Append(new StyleName { Val = "Normal" });
        normalStyle.Append(new StyleParagraphProperties(
            new SpacingBetweenLines { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        normalStyle.Append(new StyleRunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new FontSize { Val = "22" }
        ));
        styles.Append(normalStyle);

        // Style nagłówków
        for (int i = 1; i <= 6; i++)
        {
            var headingStyle = CreateHeadingStyle(i);
            styles.Append(headingStyle);
        }

        stylesPart.Styles = styles;
    }

    /// <summary>
    /// Tworzy styl nagłówka
    /// </summary>
    private Style CreateHeadingStyle(int level)
    {
        var fontSize = level switch
        {
            1 => "32",
            2 => "28",
            3 => "26",
            4 => "24",
            5 => "22",
            6 => "20",
            _ => "22"
        };

        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Heading{level}"
        };
        style.Append(new StyleName { Val = $"Heading {level}" });
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new StyleParagraphProperties(
            new SpacingBetweenLines { Before = "240", After = "120" },
            new KeepNext(),
            new KeepLines()
        ));
        style.Append(new StyleRunProperties(
            new Bold(),
            new FontSize { Val = fontSize }
        ));

        return style;
    }

    /// <summary>
    /// Konwertuje węzły HTML na elementy Body
    /// </summary>
    private void ConvertHtmlToBody(HtmlNode node, Body body)
    {
        foreach (var child in node.ChildNodes)
        {
            var elements = ConvertHtmlNode(child);
            foreach (var element in elements)
            {
                body.Append(element);
            }
        }

        // Jeśli body jest puste, dodaj pusty paragraf
        if (!body.Elements<Paragraph>().Any())
        {
            body.Append(new Paragraph());
        }
    }

    /// <summary>
    /// Konwertuje węzeł HTML na elementy OpenXML
    /// </summary>
    private List<OpenXmlElement> ConvertHtmlNode(HtmlNode node)
    {
        var elements = new List<OpenXmlElement>();

        switch (node.NodeType)
        {
            case HtmlNodeType.Text:
                // Tekst bez tagów - jeśli nie jest pusty, dodaj jako paragraf
                var text = System.Net.WebUtility.HtmlDecode(node.InnerText);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    elements.Add(CreateParagraph(text));
                }
                break;

            case HtmlNodeType.Element:
                elements.AddRange(ConvertHtmlElement(node));
                break;
        }

        return elements;
    }

    /// <summary>
    /// Konwertuje element HTML na elementy OpenXML
    /// </summary>
    private List<OpenXmlElement> ConvertHtmlElement(HtmlNode node)
    {
        var elements = new List<OpenXmlElement>();
        var tagName = node.Name.ToLower();

        switch (tagName)
        {
            case "p":
                elements.Add(ConvertParagraphElement(node));
                break;

            case "h1":
            case "h2":
            case "h3":
            case "h4":
            case "h5":
            case "h6":
                var level = int.Parse(tagName[1].ToString());
                elements.Add(ConvertHeadingElement(node, level));
                break;

            case "div":
                // Div może zawierać wiele elementów
                if (node.HasClass("page-break"))
                {
                    elements.Add(CreatePageBreak());
                }
                else
                {
                    foreach (var child in node.ChildNodes)
                    {
                        elements.AddRange(ConvertHtmlNode(child));
                    }
                }
                break;

            case "br":
                // Nowa linia - dodaj pusty paragraf
                elements.Add(new Paragraph());
                break;

            case "ul":
            case "ol":
                elements.AddRange(ConvertListElement(node, tagName == "ol"));
                break;

            case "table":
                elements.Add(ConvertTableElement(node));
                break;

            case "img":
                var imgPara = ConvertImageElement(node);
                if (imgPara != null)
                    elements.Add(imgPara);
                break;

            case "a":
                elements.Add(ConvertAnchorElement(node));
                break;

            case "span":
            case "strong":
            case "b":
            case "em":
            case "i":
            case "u":
            case "s":
            case "strike":
            case "sub":
            case "sup":
                // Inline elementy - utwórz paragraf
                elements.Add(ConvertInlineElement(node));
                break;

            case "blockquote":
                elements.Add(ConvertBlockquoteElement(node));
                break;

            default:
                // Nieznany tag - przetwórz dzieci
                foreach (var child in node.ChildNodes)
                {
                    elements.AddRange(ConvertHtmlNode(child));
                }
                break;
        }

        return elements;
    }

    /// <summary>
    /// Konwertuje element P na Paragraph
    /// </summary>
    private Paragraph ConvertParagraphElement(HtmlNode node)
    {
        var paragraph = new Paragraph();
        var props = new ParagraphProperties();

        // Parsuj style inline
        var style = node.GetAttributeValue("style", "");
        ApplyParagraphStyle(props, style);

        if (props.HasChildren)
            paragraph.Append(props);

        // Dodaj zawartość
        AppendInlineContent(paragraph, node);

        return paragraph;
    }

    /// <summary>
    /// Konwertuje nagłówek na Paragraph
    /// </summary>
    private Paragraph ConvertHeadingElement(HtmlNode node, int level)
    {
        var paragraph = new Paragraph();
        var props = new ParagraphProperties();
        props.Append(new ParagraphStyleId { Val = $"Heading{level}" });

        paragraph.Append(props);
        AppendInlineContent(paragraph, node);

        return paragraph;
    }

    /// <summary>
    /// Konwertuje listę na paragrafy z numeracją
    /// </summary>
    private List<OpenXmlElement> ConvertListElement(HtmlNode node, bool ordered)
    {
        var elements = new List<OpenXmlElement>();
        var listItems = node.SelectNodes(".//li");

        if (listItems != null)
        {
            foreach (var li in listItems)
            {
                var para = new Paragraph();
                var props = new ParagraphProperties();
                
                // Dodaj wcięcie dla listy
                props.Append(new Indentation { Left = "720", Hanging = "360" });
                para.Append(props);

                // Dodaj marker listy
                var bullet = ordered ? "• " : "• ";
                var bulletRun = new Run(new Text(bullet) { Space = SpaceProcessingModeValues.Preserve });
                para.Append(bulletRun);

                AppendInlineContent(para, li);
                elements.Add(para);
            }
        }

        return elements;
    }

    /// <summary>
    /// Konwertuje tabelę HTML na Table
    /// </summary>
    private Table ConvertTableElement(HtmlNode node)
    {
        var table = new Table();

        // Właściwości tabeli
        var tableProps = new TableProperties();
        tableProps.Append(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
        ));
        
        // Sprawdź szerokość ze stylu HTML
        var tableStyle = node.GetAttributeValue("style", "");
        var tableWidthMatch = Regex.Match(tableStyle, @"width:\s*(\d+)(%|px)?");
        if (tableWidthMatch.Success)
        {
            var widthValue = tableWidthMatch.Groups[1].Value;
            var unit = tableWidthMatch.Groups[2].Value;
            
            if (unit == "%" || string.IsNullOrEmpty(unit))
            {
                // Procenty - konwertuj na 50ths (5000 = 100%)
                var pct = int.Parse(widthValue);
                tableProps.Append(new TableWidth { Width = (pct * 50).ToString(), Type = TableWidthUnitValues.Pct });
            }
            else if (unit == "px")
            {
                // Piksele - konwertuj na twips (1px ≈ 15 twips)
                var px = int.Parse(widthValue);
                tableProps.Append(new TableWidth { Width = (px * 15).ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }
        else
        {
            // Domyślnie auto
            tableProps.Append(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        }
        
        tableProps.Append(new TableLayout { Type = TableLayoutValues.Autofit });
        table.Append(tableProps);

        // Zbierz informacje o kolumnach dla GridSpan
        var maxCols = 0;
        var rowNodes = node.SelectNodes(".//tr");
        if (rowNodes != null)
        {
            foreach (var rowNode in rowNodes)
            {
                var cellCount = 0;
                var cells = rowNode.SelectNodes(".//td|.//th");
                if (cells != null)
                {
                    foreach (var cellNode in cells)
                    {
                        var colspan = cellNode.GetAttributeValue("colspan", "1");
                        cellCount += int.TryParse(colspan, out var cs) ? cs : 1;
                    }
                }
                maxCols = Math.Max(maxCols, cellCount);
            }
        }

        // Utwórz siatkę tabeli
        if (maxCols > 0)
        {
            var grid = new TableGrid();
            for (int i = 0; i < maxCols; i++)
            {
                grid.Append(new GridColumn());
            }
            table.Append(grid);
        }

        // Przetwórz wiersze
        if (rowNodes != null)
        {
            foreach (var rowNode in rowNodes)
            {
                var row = new TableRow();
                var cells = rowNode.SelectNodes(".//td|.//th");
                
                if (cells != null)
                {
                    foreach (var cellNode in cells)
                    {
                        var cell = new TableCell();
                        var cellProps = new TableCellProperties();
                        
                        // Parsuj colspan
                        var colspanAttr = cellNode.GetAttributeValue("colspan", "1");
                        if (int.TryParse(colspanAttr, out var colspan) && colspan > 1)
                        {
                            cellProps.Append(new GridSpan { Val = colspan });
                        }
                        
                        // Parsuj rowspan - wymaga VMerge
                        var rowspanAttr = cellNode.GetAttributeValue("rowspan", "1");
                        if (int.TryParse(rowspanAttr, out var rowspan) && rowspan > 1)
                        {
                            cellProps.Append(new VerticalMerge { Val = MergedCellValues.Restart });
                        }
                        
                        // Parsuj szerokość komórki ze stylu
                        var cellStyle = cellNode.GetAttributeValue("style", "");
                        var cellWidthMatch = Regex.Match(cellStyle, @"width:\s*(\d+)(px|%)?");
                        if (cellWidthMatch.Success)
                        {
                            var widthVal = int.Parse(cellWidthMatch.Groups[1].Value);
                            var widthUnit = cellWidthMatch.Groups[2].Value;
                            
                            if (widthUnit == "px" || string.IsNullOrEmpty(widthUnit))
                            {
                                cellProps.Append(new TableCellWidth { Width = (widthVal * 15).ToString(), Type = TableWidthUnitValues.Dxa });
                            }
                            else if (widthUnit == "%")
                            {
                                cellProps.Append(new TableCellWidth { Width = (widthVal * 50).ToString(), Type = TableWidthUnitValues.Pct });
                            }
                        }
                        else
                        {
                            cellProps.Append(new TableCellWidth { Type = TableWidthUnitValues.Auto });
                        }
                        
                        // Parsuj kolor tła
                        if (cellStyle.Contains("background"))
                        {
                            var colorMatch = Regex.Match(cellStyle, @"background(?:-color)?:\s*#?([a-fA-F0-9]{6})");
                            if (colorMatch.Success)
                            {
                                cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = colorMatch.Groups[1].Value });
                            }
                        }
                        
                        // Wyrównanie pionowe
                        if (cellStyle.Contains("vertical-align"))
                        {
                            var vAlignMatch = Regex.Match(cellStyle, @"vertical-align:\s*(top|middle|bottom)");
                            if (vAlignMatch.Success)
                            {
                                var vAlign = vAlignMatch.Groups[1].Value switch
                                {
                                    "top" => TableVerticalAlignmentValues.Top,
                                    "middle" => TableVerticalAlignmentValues.Center,
                                    "bottom" => TableVerticalAlignmentValues.Bottom,
                                    _ => TableVerticalAlignmentValues.Top
                                };
                                cellProps.Append(new TableCellVerticalAlignment { Val = vAlign });
                            }
                        }
                        
                        // Ramki komórki
                        cellProps.Append(new TableCellBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 }
                        ));

                        cell.Append(cellProps);

                        // Dodaj zawartość komórki - może być wiele paragrafów
                        var hasContent = false;
                        foreach (var childNode in cellNode.ChildNodes)
                        {
                            if (childNode.Name.ToLower() == "p" || 
                                childNode.Name.ToLower() == "div" ||
                                childNode.Name.ToLower() == "br")
                            {
                                var elements = ConvertHtmlNode(childNode);
                                foreach (var el in elements)
                                {
                                    if (el is Paragraph para)
                                    {
                                        cell.Append(para);
                                        hasContent = true;
                                    }
                                }
                            }
                            else if (childNode.NodeType == HtmlNodeType.Text || 
                                     childNode.Name.ToLower() == "span" ||
                                     childNode.Name.ToLower() == "strong" ||
                                     childNode.Name.ToLower() == "b" ||
                                     childNode.Name.ToLower() == "em" ||
                                     childNode.Name.ToLower() == "i")
                            {
                                // Inline content - zbierz do jednego paragrafu
                                if (!hasContent)
                                {
                                    var cellPara = new Paragraph();
                                    AppendInlineContent(cellPara, cellNode);
                                    cell.Append(cellPara);
                                    hasContent = true;
                                    break;
                                }
                            }
                        }
                        
                        // Jeśli komórka jest pusta, dodaj pusty paragraf (wymagany w Word)
                        if (!hasContent)
                        {
                            cell.Append(new Paragraph());
                        }

                        row.Append(cell);
                    }
                }

                table.Append(row);
            }
        }

        return table;
    }

    /// <summary>
    /// Konwertuje obraz na Paragraph z obrazem
    /// </summary>
    private Paragraph? ConvertImageElement(HtmlNode node)
    {
        var src = node.GetAttributeValue("src", "");
        if (string.IsNullOrEmpty(src)) return null;

        // Sprawdź czy to base64
        if (src.StartsWith("data:"))
        {
            var match = Regex.Match(src, @"data:([^;]+);base64,(.+)");
            if (match.Success)
            {
                var contentType = match.Groups[1].Value;
                var base64 = match.Groups[2].Value;
                
                try
                {
                    var imageBytes = System.Convert.FromBase64String(base64);
                    return CreateImageParagraph(imageBytes, contentType, node);
                }
                catch
                {
                    return null;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Tworzy paragraf z obrazem
    /// </summary>
    private Paragraph CreateImageParagraph(byte[] imageBytes, string contentType, HtmlNode node)
    {
        var imagePart = contentType switch
        {
            "image/png" => _mainPart!.AddImagePart(ImagePartType.Png),
            "image/gif" => _mainPart!.AddImagePart(ImagePartType.Gif),
            "image/bmp" => _mainPart!.AddImagePart(ImagePartType.Bmp),
            _ => _mainPart!.AddImagePart(ImagePartType.Jpeg)
        };

        using var stream = new MemoryStream(imageBytes);
        imagePart.FeedData(stream);

        var relationshipId = _mainPart!.GetIdOfPart(imagePart);

        // Pobierz wymiary ze stylu
        var style = node.GetAttributeValue("style", "");
        var widthMatch = Regex.Match(style, @"width:\s*(\d+)px");
        var width = widthMatch.Success ? int.Parse(widthMatch.Groups[1].Value) : 200;
        var height = (int)(width * 0.75); // Domyślny aspect ratio

        // Konwertuj px na EMU
        var widthEmu = width * 9525;
        var heightEmu = height * 9525;

        var element = new Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmu, Cy = heightEmu },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = (uint)++_imageCounter, Name = $"Image{_imageCounter}" },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = (uint)_imageCounter, Name = $"Image{_imageCounter}" },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId },
                                new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                    new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
        );

        return new Paragraph(new Run(element));
    }

    /// <summary>
    /// Konwertuje link na Paragraph
    /// </summary>
    private Paragraph ConvertAnchorElement(HtmlNode node)
    {
        var para = new Paragraph();
        var href = node.GetAttributeValue("href", "#");

        // Dodaj hyperlink
        var relationshipId = _mainPart!.AddHyperlinkRelationship(new Uri(href, UriKind.RelativeOrAbsolute), true).Id;
        
        var hyperlink = new Hyperlink { Id = relationshipId };
        
        foreach (var child in node.ChildNodes)
        {
            var runs = CreateRunsFromNode(child);
            foreach (var run in runs)
            {
                // Dodaj styl linku
                run.RunProperties ??= new RunProperties();
                run.RunProperties.Append(new Color { Val = "0000FF" });
                run.RunProperties.Append(new Underline { Val = UnderlineValues.Single });
                hyperlink.Append(run);
            }
        }

        para.Append(hyperlink);
        return para;
    }

    /// <summary>
    /// Konwertuje inline element na Paragraph
    /// </summary>
    private Paragraph ConvertInlineElement(HtmlNode node)
    {
        var para = new Paragraph();
        AppendInlineContent(para, node);
        return para;
    }

    /// <summary>
    /// Konwertuje blockquote na Paragraph
    /// </summary>
    private Paragraph ConvertBlockquoteElement(HtmlNode node)
    {
        var para = new Paragraph();
        var props = new ParagraphProperties();
        props.Append(new Indentation { Left = "720" });
        props.Append(new ParagraphBorders(
            new LeftBorder { Val = BorderValues.Single, Size = 24, Color = "CCCCCC" }
        ));
        para.Append(props);

        AppendInlineContent(para, node);
        return para;
    }

    /// <summary>
    /// Dodaje inline content do paragrafu
    /// </summary>
    private void AppendInlineContent(Paragraph paragraph, HtmlNode node)
    {
        foreach (var child in node.ChildNodes)
        {
            var runs = CreateRunsFromNode(child);
            foreach (var run in runs)
            {
                paragraph.Append(run);
            }
        }

        // Jeśli paragraf jest pusty, dodaj pusty run
        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any())
        {
            paragraph.Append(new Run(new Text("")));
        }
    }

    /// <summary>
    /// Tworzy Run z węzła HTML
    /// </summary>
    private List<Run> CreateRunsFromNode(HtmlNode node, RunProperties? inheritedProps = null)
    {
        var runs = new List<Run>();

        switch (node.NodeType)
        {
            case HtmlNodeType.Text:
                var text = System.Net.WebUtility.HtmlDecode(node.InnerText);
                if (!string.IsNullOrEmpty(text))
                {
                    var run = new Run();
                    if (inheritedProps != null)
                    {
                        run.Append(inheritedProps.CloneNode(true));
                    }
                    run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    runs.Add(run);
                }
                break;

            case HtmlNodeType.Element:
                var tagName = node.Name.ToLower();
                var newProps = (inheritedProps?.CloneNode(true) as RunProperties) ?? new RunProperties();
                
                // Aplikuj formatowanie na podstawie tagu
                switch (tagName)
                {
                    case "strong":
                    case "b":
                        newProps.Append(new Bold());
                        break;
                    case "em":
                    case "i":
                        newProps.Append(new Italic());
                        break;
                    case "u":
                        newProps.Append(new Underline { Val = UnderlineValues.Single });
                        break;
                    case "s":
                    case "strike":
                        newProps.Append(new Strike());
                        break;
                    case "sub":
                        newProps.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
                        break;
                    case "sup":
                        newProps.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
                        break;
                    case "br":
                        runs.Add(new Run(new Break()));
                        return runs;
                }

                // Parsuj style inline
                var style = node.GetAttributeValue("style", "");
                ApplyRunStyle(newProps, style);

                // Rekurencyjnie przetwórz dzieci
                foreach (var child in node.ChildNodes)
                {
                    runs.AddRange(CreateRunsFromNode(child, newProps));
                }
                break;
        }

        return runs;
    }

    /// <summary>
    /// Aplikuje styl CSS do ParagraphProperties
    /// </summary>
    private void ApplyParagraphStyle(ParagraphProperties props, string style)
    {
        if (string.IsNullOrEmpty(style)) return;

        // Text-align
        var alignMatch = Regex.Match(style, @"text-align:\s*(left|center|right|justify)");
        if (alignMatch.Success)
        {
            var align = alignMatch.Groups[1].Value switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                "justify" => JustificationValues.Both,
                _ => JustificationValues.Left
            };
            props.Append(new Justification { Val = align });
        }

        // Margin/padding
        var marginLeftMatch = Regex.Match(style, @"margin-left:\s*(\d+)px");
        var marginRightMatch = Regex.Match(style, @"margin-right:\s*(\d+)px");
        var textIndentMatch = Regex.Match(style, @"text-indent:\s*(\d+)px");

        if (marginLeftMatch.Success || marginRightMatch.Success || textIndentMatch.Success)
        {
            var indentation = new Indentation();
            if (marginLeftMatch.Success)
            {
                indentation.Left = PxToTwips(int.Parse(marginLeftMatch.Groups[1].Value)).ToString();
            }
            if (marginRightMatch.Success)
            {
                indentation.Right = PxToTwips(int.Parse(marginRightMatch.Groups[1].Value)).ToString();
            }
            if (textIndentMatch.Success)
            {
                indentation.FirstLine = PxToTwips(int.Parse(textIndentMatch.Groups[1].Value)).ToString();
            }
            props.Append(indentation);
        }

        // Margin top/bottom
        var marginTopMatch = Regex.Match(style, @"margin-top:\s*(\d+)px");
        var marginBottomMatch = Regex.Match(style, @"margin-bottom:\s*(\d+)px");
        var lineHeightMatch = Regex.Match(style, @"line-height:\s*([\d.]+)");

        if (marginTopMatch.Success || marginBottomMatch.Success || lineHeightMatch.Success)
        {
            var spacing = new SpacingBetweenLines();
            if (marginTopMatch.Success)
            {
                spacing.Before = PxToTwips(int.Parse(marginTopMatch.Groups[1].Value)).ToString();
            }
            if (marginBottomMatch.Success)
            {
                spacing.After = PxToTwips(int.Parse(marginBottomMatch.Groups[1].Value)).ToString();
            }
            if (lineHeightMatch.Success)
            {
                var lineHeight = double.Parse(lineHeightMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                spacing.Line = ((int)(lineHeight * 240)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            props.Append(spacing);
        }
    }

    /// <summary>
    /// Aplikuje styl CSS do RunProperties
    /// </summary>
    private void ApplyRunStyle(RunProperties props, string style)
    {
        if (string.IsNullOrEmpty(style)) return;

        // Font-weight
        if (style.Contains("font-weight:bold") || style.Contains("font-weight: bold"))
        {
            if (!props.Elements<Bold>().Any())
                props.Append(new Bold());
        }

        // Font-style
        if (style.Contains("font-style:italic") || style.Contains("font-style: italic"))
        {
            if (!props.Elements<Italic>().Any())
                props.Append(new Italic());
        }

        // Text-decoration
        if (style.Contains("text-decoration:underline") || style.Contains("text-decoration: underline"))
        {
            if (!props.Elements<Underline>().Any())
                props.Append(new Underline { Val = UnderlineValues.Single });
        }
        if (style.Contains("text-decoration:line-through") || style.Contains("text-decoration: line-through"))
        {
            if (!props.Elements<Strike>().Any())
                props.Append(new Strike());
        }

        // Font-size
        var fontSizeMatch = Regex.Match(style, @"font-size:\s*([\d.]+)(pt|px)");
        if (fontSizeMatch.Success)
        {
            var size = double.Parse(fontSizeMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = fontSizeMatch.Groups[2].Value;
            
            if (unit == "px")
            {
                size = size * 0.75; // px to pt
            }
            
            var halfPoints = ((int)(size * 2)).ToString();
            props.Append(new FontSize { Val = halfPoints });
        }

        // Font-family
        var fontFamilyMatch = Regex.Match(style, @"font-family:\s*'?([^',;]+)'?");
        if (fontFamilyMatch.Success)
        {
            var fontName = fontFamilyMatch.Groups[1].Value.Trim();
            props.Append(new RunFonts { Ascii = fontName, HighAnsi = fontName });
        }

        // Color
        var colorMatch = Regex.Match(style, @"(?<!background-)color:\s*#?([a-fA-F0-9]{6})");
        if (colorMatch.Success)
        {
            props.Append(new Color { Val = colorMatch.Groups[1].Value });
        }

        // Background-color
        var bgColorMatch = Regex.Match(style, @"background-color:\s*#?([a-fA-F0-9]{6})");
        if (bgColorMatch.Success)
        {
            props.Append(new Shading { Fill = bgColorMatch.Groups[1].Value, Val = ShadingPatternValues.Clear });
        }

        // Vertical align
        if (style.Contains("vertical-align:super"))
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }
        if (style.Contains("vertical-align:sub"))
        {
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }
    }

    /// <summary>
    /// Tworzy prosty paragraf z tekstem
    /// </summary>
    private Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    /// <summary>
    /// Tworzy paragraf z page break
    /// </summary>
    private Paragraph CreatePageBreak()
    {
        return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
    }

    /// <summary>
    /// Ustawia metadane dokumentu
    /// </summary>
    private void SetDocumentMetadata(WordprocessingDocument document, DocumentMetadata metadata)
    {
        var props = document.PackageProperties;
        
        if (!string.IsNullOrEmpty(metadata.Title))
            props.Title = metadata.Title;
        
        if (!string.IsNullOrEmpty(metadata.Author))
            props.Creator = metadata.Author;
        
        if (!string.IsNullOrEmpty(metadata.Subject))
            props.Subject = metadata.Subject;
        
        props.Created = DateTime.UtcNow;
        props.Modified = DateTime.UtcNow;
    }

    /// <summary>
    /// Dodaje ustawienia strony (marginesy itp.)
    /// </summary>
    private void AddPageSettings(Body body)
    {
        var sectionProps = new SectionProperties();
        
        // Ustawienia strony A4
        sectionProps.Append(new PageSize 
        { 
            Width = 11906, // A4 width in twips
            Height = 16838 // A4 height in twips
        });
        
        // Marginesy
        sectionProps.Append(new PageMargin
        {
            Top = 1440,    // 1 inch
            Right = 1440,
            Bottom = 1440,
            Left = 1440,
            Header = 720,
            Footer = 720
        });

        body.Append(sectionProps);
    }

    /// <summary>
    /// Konwertuje piksele na twips
    /// </summary>
    private int PxToTwips(int px)
    {
        // 1 inch = 96 px, 1 inch = 1440 twips
        return (int)(px / 96.0 * 1440);
    }
}

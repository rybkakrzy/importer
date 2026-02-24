using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Importer.Domain.Interfaces;
using Importer.Domain.Models;
using HtmlAgilityPack;

namespace Importer.Infrastructure.Services;

/// <summary>
/// Serwis do konwersji HTML na dokument DOCX
/// Własna implementacja generatora OpenXML z wysoką dokładnością odwzorowania stylów
/// </summary>
public class HtmlToDocxConverter : IHtmlToDocxConverter
{
    private MainDocumentPart? _mainPart;
    private readonly Dictionary<string, string> _imageRelationships = new();
    private int _imageCounter = 0;
    private int _numberingId = 1;
    private NumberingDefinitionsPart? _numberingPart;
    private readonly Dictionary<int, int> _abstractNumIds = new(); // track list numbering

    /// <summary>
    /// Konwertuje HTML na plik DOCX
    /// </summary>
    public byte[] Convert(string html, DocumentMetadata? metadata = null, HeaderFooterContent? header = null, HeaderFooterContent? footer = null)
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

            // Dodaj nagłówek i stopkę
            AddHeaderAndFooter(document, header, footer);

            // Dodaj ustawienia strony
            AddPageSettings(body, header, footer);

            document.Save();
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Dodaje nagłówek i stopkę do dokumentu
    /// </summary>
    private void AddHeaderAndFooter(WordprocessingDocument document, HeaderFooterContent? header, HeaderFooterContent? footer)
    {
        if (_mainPart == null) return;

        if (header != null && !string.IsNullOrWhiteSpace(header.Html))
        {
            var headerPart = _mainPart.AddNewPart<HeaderPart>();
            var headerElement = new Header();
            
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(header.Html);
            ConvertHtmlToHeaderFooter(htmlDoc.DocumentNode, headerElement);
            
            headerPart.Header = headerElement;
            headerPart.Header.Save();
            
            var headerPartId = _mainPart.GetIdOfPart(headerPart);
            AddHeaderReference(headerPartId);
        }

        if (footer != null && !string.IsNullOrWhiteSpace(footer.Html))
        {
            var footerPart = _mainPart.AddNewPart<FooterPart>();
            var footerElement = new Footer();
            
            var htmlDoc = new HtmlDocument();
            var footerHtml = footer.Html
                .Replace("{page}", "<span class=\"field-page\"></span>")
                .Replace("{pages}", "<span class=\"field-numpages\"></span>");
            htmlDoc.LoadHtml(footerHtml);
            ConvertHtmlToHeaderFooter(htmlDoc.DocumentNode, footerElement);
            
            footerPart.Footer = footerElement;
            footerPart.Footer.Save();
            
            var footerPartId = _mainPart.GetIdOfPart(footerPart);
            AddFooterReference(footerPartId);
        }
    }

    /// <summary>
    /// Konwertuje HTML na elementy nagłówka/stopki
    /// </summary>
    private void ConvertHtmlToHeaderFooter(HtmlNode node, OpenXmlCompositeElement parent)
    {
        foreach (var child in node.ChildNodes)
        {
            switch (child.Name.ToLower())
            {
                case "p":
                    parent.Append(ConvertParagraphElement(child));
                    break;
                case "span":
                    if (child.HasClass("field-page"))
                    {
                        var pagePara = new Paragraph(new Run(
                            new SimpleField { Instruction = " PAGE " }
                        ));
                        parent.Append(pagePara);
                    }
                    else if (child.HasClass("field-numpages"))
                    {
                        var numPagesPara = new Paragraph(new Run(
                            new SimpleField { Instruction = " NUMPAGES " }
                        ));
                        parent.Append(numPagesPara);
                    }
                    else
                    {
                        ConvertHtmlToHeaderFooter(child, parent);
                    }
                    break;
                case "table":
                    parent.Append(ConvertTableElement(child));
                    break;
                case "#text":
                    var text = child.InnerText.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        parent.Append(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
                    }
                    break;
                default:
                    ConvertHtmlToHeaderFooter(child, parent);
                    break;
            }
        }
    }

    private void AddHeaderReference(string headerPartId)
    {
        var body = _mainPart?.Document?.Body;
        if (body == null) return;

        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }

        sectionProps.InsertAt(new HeaderReference
        {
            Type = HeaderFooterValues.Default,
            Id = headerPartId
        }, 0);
    }

    private void AddFooterReference(string footerPartId)
    {
        var body = _mainPart?.Document?.Body;
        if (body == null) return;

        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }

        sectionProps.InsertAt(new FooterReference
        {
            Type = HeaderFooterValues.Default,
            Id = footerPartId
        }, 0);
    }

    /// <summary>
    /// Dodaje domyślne style do dokumentu z dokładnym odwzorowaniem
    /// </summary>
    private void AddDocumentStyles(WordprocessingDocument document)
    {
        var stylesPart = _mainPart!.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        // Domyślne właściwości dokumentu
        var docDefaults = new DocDefaults(
            new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Calibri" },
                    new FontSize { Val = "22" },
                    new FontSizeComplexScript { Val = "22" },
                    new Languages { Val = "pl-PL", EastAsia = "pl-PL" }
                )
            ),
            new ParagraphPropertiesDefault(
                new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }
                )
            )
        );
        styles.Append(docDefaults);

        // Styl normalny
        var normalStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        normalStyle.Append(new StyleName { Val = "Normal" });
        normalStyle.Append(new PrimaryStyle());
        normalStyle.Append(new StyleParagraphProperties(
            new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }
        ));
        normalStyle.Append(new StyleRunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new FontSize { Val = "22" }
        ));
        styles.Append(normalStyle);

        // Style nagłówków z dokładnymi rozmiarami jak w Word
        string[] headingColors = { "2F5496", "2F5496", "1F3763", "2F5496", "2F5496", "1F3763" };
        string[] headingSizes = { "32", "26", "24", "22", "22", "22" };
        bool[] headingBold = { true, true, true, true, true, false };
        bool[] headingItalic = { false, false, false, true, false, true };
        int[] headingSpaceBefore = { 240, 40, 40, 40, 40, 40 };

        for (int i = 1; i <= 6; i++)
        {
            var headingStyle = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = $"Heading{i}"
            };
            headingStyle.Append(new StyleName { Val = $"Heading {i}" });
            headingStyle.Append(new BasedOn { Val = "Normal" });
            headingStyle.Append(new NextParagraphStyle { Val = "Normal" });
            headingStyle.Append(new PrimaryStyle());
            
            var paraProps = new StyleParagraphProperties(
                new SpacingBetweenLines { Before = headingSpaceBefore[i - 1].ToString(), After = "0" },
                new KeepNext(),
                new KeepLines(),
                new OutlineLevel { Val = i - 1 }
            );
            headingStyle.Append(paraProps);
            
            var runPropsElements = new List<OpenXmlElement>
            {
                new RunFonts { Ascii = "Calibri Light", HighAnsi = "Calibri Light" },
                new FontSize { Val = headingSizes[i - 1] },
                new Color { Val = headingColors[i - 1] }
            };
            
            if (headingBold[i - 1]) runPropsElements.Add(new Bold());
            if (headingItalic[i - 1]) runPropsElements.Add(new Italic());
            
            headingStyle.Append(new StyleRunProperties(runPropsElements.ToArray()));
            styles.Append(headingStyle);
        }

        // Styl hiperłącza
        var hyperlinkStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "Hyperlink"
        };
        hyperlinkStyle.Append(new StyleName { Val = "Hyperlink" });
        hyperlinkStyle.Append(new StyleRunProperties(
            new Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink },
            new Underline { Val = UnderlineValues.Single }
        ));
        styles.Append(hyperlinkStyle);

        // Styl akapitu listy
        var listParagraph = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "ListParagraph"
        };
        listParagraph.Append(new StyleName { Val = "List Paragraph" });
        listParagraph.Append(new BasedOn { Val = "Normal" });
        listParagraph.Append(new StyleParagraphProperties(
            new Indentation { Left = "720" }
        ));
        styles.Append(listParagraph);

        stylesPart.Styles = styles;
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

        if (!body.Elements<Paragraph>().Any() && !body.Elements<Table>().Any())
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
    /// Konwertuje element HTML na elementy OpenXML z pełnym odwzorowaniem
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

            case "h1": case "h2": case "h3": case "h4": case "h5": case "h6":
                var level = int.Parse(tagName[1].ToString());
                elements.Add(ConvertHeadingElement(node, level));
                break;

            case "div":
                if (node.HasClass("page-break"))
                {
                    elements.Add(CreatePageBreak());
                }
                else if (node.HasClass("document-content"))
                {
                    foreach (var child in node.ChildNodes)
                        elements.AddRange(ConvertHtmlNode(child));
                }
                else
                {
                    foreach (var child in node.ChildNodes)
                        elements.AddRange(ConvertHtmlNode(child));
                }
                break;

            case "br":
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

            case "blockquote":
                elements.Add(ConvertBlockquoteElement(node));
                break;

            case "hr":
                elements.Add(CreateHorizontalRule());
                break;

            case "span": case "strong": case "b": case "em": case "i": case "u": case "s": case "strike": case "sub": case "sup":
                elements.Add(ConvertInlineElement(node));
                break;

            default:
                foreach (var child in node.ChildNodes)
                    elements.AddRange(ConvertHtmlNode(child));
                break;
        }

        return elements;
    }

    /// <summary>
    /// Konwertuje element P na Paragraph z pełnym parsowaniem stylów
    /// </summary>
    private Paragraph ConvertParagraphElement(HtmlNode node)
    {
        var paragraph = new Paragraph();
        var props = new ParagraphProperties();

        var style = node.GetAttributeValue("style", "");
        ApplyParagraphStyle(props, style);

        // Sprawdź data-style-id dla zachowania oryginalnego stylu Word
        var styleId = node.GetAttributeValue("data-style-id", "");
        if (!string.IsNullOrEmpty(styleId))
        {
            props.Append(new ParagraphStyleId { Val = styleId });
        }

        if (props.HasChildren)
            paragraph.Append(props);

        AppendInlineContent(paragraph, node);

        return paragraph;
    }

    /// <summary>
    /// Konwertuje nagłówek na Paragraph z odpowiednim stylem
    /// </summary>
    private Paragraph ConvertHeadingElement(HtmlNode node, int level)
    {
        var paragraph = new Paragraph();
        var props = new ParagraphProperties();
        props.Append(new ParagraphStyleId { Val = $"Heading{level}" });

        // Dodaj dodatkowe style inline
        var style = node.GetAttributeValue("style", "");
        if (!string.IsNullOrEmpty(style))
        {
            ApplyParagraphStyleExtras(props, style);
        }

        paragraph.Append(props);
        AppendInlineContent(paragraph, node);

        return paragraph;
    }

    /// <summary>
    /// Konwertuje listę na paragrafy z prawidłową definicją numeracji Word
    /// </summary>
    private List<OpenXmlElement> ConvertListElement(HtmlNode node, bool ordered, int level = 0)
    {
        var elements = new List<OpenXmlElement>();
        
        // Upewnij się że mamy NumberingDefinitionsPart
        EnsureNumberingPart();
        
        // Utwórz definicję numeracji dla tej listy
        var abstractNumId = CreateAbstractNumbering(ordered);
        var numId = CreateNumberingInstance(abstractNumId);

        foreach (var child in node.ChildNodes)
        {
            if (child.Name.ToLower() != "li") continue;
            
            // Sprawdź czy li zawiera zagnieżdżoną listę
            var nestedList = child.SelectSingleNode("./ul|./ol");
            
            // Utwórz paragraf z elementem listy
            var para = new Paragraph();
            var props = new ParagraphProperties();
            props.Append(new ParagraphStyleId { Val = "ListParagraph" });
            props.Append(new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = numId }
            ));
            
            // Parsuj style inline z li
            var liStyle = child.GetAttributeValue("style", "");
            if (!string.IsNullOrEmpty(liStyle))
            {
                ApplyParagraphStyleExtras(props, liStyle);
            }
            
            para.Append(props);

            // Dodaj zawartość (bez zagnieżdżonej listy)
            foreach (var liChild in child.ChildNodes)
            {
                if (liChild.Name.ToLower() == "ul" || liChild.Name.ToLower() == "ol")
                    continue; // Zagnieżdżona lista będzie obsłużona osobno
                    
                var runs = CreateRunsFromNode(liChild);
                foreach (var run in runs)
                    para.Append(run);
            }

            if (!para.Elements<Run>().Any() && !para.Elements<Hyperlink>().Any())
            {
                para.Append(new Run(new Text("") { Space = SpaceProcessingModeValues.Preserve }));
            }

            elements.Add(para);

            // Obsłuż zagnieżdżoną listę
            if (nestedList != null)
            {
                var isOrdered = nestedList.Name.ToLower() == "ol";
                elements.AddRange(ConvertListElement(nestedList, isOrdered, level + 1));
            }
        }

        return elements;
    }

    /// <summary>
    /// Upewnia się że dokument ma NumberingDefinitionsPart
    /// </summary>
    private void EnsureNumberingPart()
    {
        if (_numberingPart != null) return;
        
        _numberingPart = _mainPart!.AddNewPart<NumberingDefinitionsPart>();
        _numberingPart.Numbering = new Numbering();
        _numberingPart.Numbering.Save();
    }

    /// <summary>
    /// Tworzy definicję abstrakcyjnej numeracji
    /// </summary>
    private int CreateAbstractNumbering(bool ordered)
    {
        var abstractNumId = _numberingId++;
        
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
        
        // Zdefiniuj 9 poziomów
        for (int lvl = 0; lvl < 9; lvl++)
        {
            var levelDef = new Level { LevelIndex = lvl };
            levelDef.Append(new StartNumberingValue { Val = 1 });
            
            if (ordered)
            {
                var format = lvl switch
                {
                    0 => NumberFormatValues.Decimal,
                    1 => NumberFormatValues.LowerLetter,
                    2 => NumberFormatValues.LowerRoman,
                    3 => NumberFormatValues.Decimal,
                    4 => NumberFormatValues.LowerLetter,
                    5 => NumberFormatValues.LowerRoman,
                    _ => NumberFormatValues.Decimal
                };
                levelDef.Append(new NumberingFormat { Val = format });
                levelDef.Append(new LevelText { Val = $"%{lvl + 1}." });
                levelDef.Append(new NumberingSymbolRunProperties(
                    new RunFonts { Hint = FontTypeHintValues.Default }
                ));
            }
            else
            {
                levelDef.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
                
                // Użyj standardowych czcionek i znaków Word:
                // Poziom 0,3,6: Symbol font, znak \uF0B7 (bullet)
                // Poziom 1,4,7: Courier New, znak "o" (circle)
                // Poziom 2,5,8: Wingdings, znak \uF0A7 (filled square)
                var bulletType = lvl % 3;
                switch (bulletType)
                {
                    case 0: // Filled circle (Symbol)
                        levelDef.Append(new LevelText { Val = "\uF0B7" });
                        levelDef.Append(new NumberingSymbolRunProperties(
                            new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }
                        ));
                        break;
                    case 1: // Open circle (Courier New)
                        levelDef.Append(new LevelText { Val = "o" });
                        levelDef.Append(new NumberingSymbolRunProperties(
                            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New", Hint = FontTypeHintValues.Default }
                        ));
                        break;
                    case 2: // Filled square (Wingdings)
                        levelDef.Append(new LevelText { Val = "\uF0A7" });
                        levelDef.Append(new NumberingSymbolRunProperties(
                            new RunFonts { Ascii = "Wingdings", HighAnsi = "Wingdings", Hint = FontTypeHintValues.Default }
                        ));
                        break;
                }
            }
            
            levelDef.Append(new LevelJustification { Val = LevelJustificationValues.Left });
            
            var indent = 720 * (lvl + 1);
            levelDef.Append(new PreviousParagraphProperties(
                new Indentation { Left = indent.ToString(), Hanging = "360" }
            ));
            
            abstractNum.Append(levelDef);
        }
        
        // Wstaw na początku (przed instancjami)
        var firstInstance = _numberingPart!.Numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstInstance != null)
            _numberingPart.Numbering.InsertBefore(abstractNum, firstInstance);
        else
            _numberingPart.Numbering.Append(abstractNum);
        
        _numberingPart.Numbering.Save();
        return abstractNumId;
    }

    /// <summary>
    /// Tworzy instancję numeracji
    /// </summary>
    private int CreateNumberingInstance(int abstractNumId)
    {
        var numId = _numberingId++;
        
        var numInstance = new NumberingInstance { NumberID = numId };
        numInstance.Append(new AbstractNumId { Val = abstractNumId });
        
        _numberingPart!.Numbering.Append(numInstance);
        _numberingPart.Numbering.Save();
        
        return numId;
    }

    /// <summary>
    /// Konwertuje tabelę HTML na Table z pełnym odwzorowaniem stylów
    /// </summary>
    private Table ConvertTableElement(HtmlNode node)
    {
        var table = new Table();
        var tableProps = new TableProperties();
        
        // Domyślne obramowania
        var defaultBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
        );
        
        // Parsuj style tabeli
        var tableStyle = node.GetAttributeValue("style", "");
        
        // Szerokość
        var tableWidthMatch = Regex.Match(tableStyle, @"width:\s*(\d+)(px|%)?");
        if (tableWidthMatch.Success)
        {
            var widthValue = tableWidthMatch.Groups[1].Value;
            var unit = tableWidthMatch.Groups[2].Value;
            
            if (unit == "%")
            {
                var pct = int.Parse(widthValue);
                tableProps.Append(new TableWidth { Width = (pct * 50).ToString(), Type = TableWidthUnitValues.Pct });
            }
            else if (unit == "px" || string.IsNullOrEmpty(unit))
            {
                var px = int.Parse(widthValue);
                tableProps.Append(new TableWidth { Width = (px * 15).ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }
        else
        {
            tableProps.Append(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
        }
        
        // Wyrównanie tabeli
        if (tableStyle.Contains("margin-left:auto") && tableStyle.Contains("margin-right:auto"))
        {
            tableProps.Append(new TableJustification { Val = TableRowAlignmentValues.Center });
        }
        else if (tableStyle.Contains("margin-left:auto"))
        {
            tableProps.Append(new TableJustification { Val = TableRowAlignmentValues.Right });
        }

        // Parsuj obramowania tabeli z CSS
        var borderMatch = Regex.Match(tableStyle, @"border:\s*([\d.]+)px\s+(\w+)\s+#?([a-fA-F0-9]{3,6})");
        if (borderMatch.Success)
        {
            var bSize = (uint)(double.Parse(borderMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 8);
            var bStyle = ParseBorderStyle(borderMatch.Groups[2].Value);
            var bColor = NormalizeColor(borderMatch.Groups[3].Value);
            
            defaultBorders = new TableBorders(
                new TopBorder { Val = bStyle, Size = bSize, Color = bColor },
                new BottomBorder { Val = bStyle, Size = bSize, Color = bColor },
                new LeftBorder { Val = bStyle, Size = bSize, Color = bColor },
                new RightBorder { Val = bStyle, Size = bSize, Color = bColor },
                new InsideHorizontalBorder { Val = bStyle, Size = bSize, Color = bColor },
                new InsideVerticalBorder { Val = bStyle, Size = bSize, Color = bColor }
            );
        }
        
        tableProps.Append(defaultBorders);
        tableProps.Append(new TableLayout { Type = TableLayoutValues.Autofit });
        
        // Domyślny padding komórek
        tableProps.Append(new TableCellMarginDefault(
            new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
            new TableCellLeftMargin { Width = 80, Type = TableWidthValues.Dxa },
            new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
            new TableCellRightMargin { Width = 80, Type = TableWidthValues.Dxa }
        ));
        
        table.Append(tableProps);

        // Oblicz liczbę kolumn
        var maxCols = 0;
        var rowNodes = node.SelectNodes(".//tr");
        if (rowNodes != null)
        {
            foreach (var rowNode in rowNodes)
            {
                var cellCount = 0;
                var cells = rowNode.SelectNodes("./td|./th");
                if (cells != null)
                {
                    foreach (var cellNode in cells)
                    {
                        var cs = cellNode.GetAttributeValue("colspan", "1");
                        cellCount += int.TryParse(cs, out var csVal) ? csVal : 1;
                    }
                }
                maxCols = Math.Max(maxCols, cellCount);
            }
        }

        // Siatka tabeli
        if (maxCols > 0)
        {
            var grid = new TableGrid();
            for (int i = 0; i < maxCols; i++)
                grid.Append(new GridColumn());
            table.Append(grid);
        }

        // Przetwórz wiersze
        if (rowNodes != null)
        {
            foreach (var rowNode in rowNodes)
            {
                var row = new TableRow();
                
                // Wysokość wiersza
                var rowStyle = rowNode.GetAttributeValue("style", "");
                var rowHeightMatch = Regex.Match(rowStyle, @"(?:min-)?height:\s*([\d.]+)px");
                if (rowHeightMatch.Success)
                {
                    var heightPx = (int)Math.Round(double.Parse(rowHeightMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
                    var heightTwips = PxToTwips(heightPx);
                    var rowProps = new TableRowProperties(
                        new TableRowHeight { Val = (uint)heightTwips, HeightType = HeightRuleValues.AtLeast }
                    );
                    row.Append(rowProps);
                }
                
                var cells = rowNode.SelectNodes("./td|./th");
                if (cells != null)
                {
                    foreach (var cellNode in cells)
                    {
                        var cell = new TableCell();
                        var cellProps = new TableCellProperties();
                        
                        // Colspan
                        var colspanAttr = cellNode.GetAttributeValue("colspan", "1");
                        if (int.TryParse(colspanAttr, out var colspan) && colspan > 1)
                            cellProps.Append(new GridSpan { Val = colspan });
                        
                        // Rowspan
                        var rowspanAttr = cellNode.GetAttributeValue("rowspan", "1");
                        if (int.TryParse(rowspanAttr, out var rowspan) && rowspan > 1)
                            cellProps.Append(new VerticalMerge { Val = MergedCellValues.Restart });
                        
                        // Parsuj style komórki
                        var cellStyle = cellNode.GetAttributeValue("style", "");
                        ApplyCellStyle(cellProps, cellStyle);
                        
                        // Obramowania komórki
                        ApplyCellBorders(cellProps, cellStyle);

                        cell.Append(cellProps);

                        // Zawartość komórki
                        var hasContent = false;
                        foreach (var childNode in cellNode.ChildNodes)
                        {
                            var childTag = childNode.Name.ToLower();
                            if (childTag == "p" || childTag == "div" || childTag == "br" ||
                                childTag == "h1" || childTag == "h2" || childTag == "h3" ||
                                childTag == "h4" || childTag == "h5" || childTag == "h6" ||
                                childTag == "ul" || childTag == "ol" || childTag == "table")
                            {
                                var els = ConvertHtmlNode(childNode);
                                foreach (var el in els)
                                {
                                    if (el is Paragraph || el is Table)
                                    {
                                        cell.Append(el);
                                        hasContent = true;
                                    }
                                }
                            }
                            else if (childNode.NodeType == HtmlNodeType.Text || 
                                     IsInlineTag(childTag))
                            {
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
                        
                        if (!hasContent)
                            cell.Append(new Paragraph());

                        row.Append(cell);
                    }
                }

                table.Append(row);
            }
        }

        return table;
    }

    /// <summary>
    /// Sprawdza czy tag jest inline
    /// </summary>
    private bool IsInlineTag(string tagName) => tagName switch
    {
        "span" or "strong" or "b" or "em" or "i" or "u" or "s" or "strike" or "sub" or "sup" or "a" => true,
        _ => false
    };

    /// <summary>
    /// Aplikuje style do komórki tabeli
    /// </summary>
    private void ApplyCellStyle(TableCellProperties cellProps, string style)
    {
        if (string.IsNullOrEmpty(style)) return;
        
        // Szerokość
        var widthMatch = Regex.Match(style, @"width:\s*(\d+)(px|%)?");
        if (widthMatch.Success)
        {
            var widthVal = int.Parse(widthMatch.Groups[1].Value);
            var widthUnit = widthMatch.Groups[2].Value;
            
            if (widthUnit == "%")
                cellProps.Append(new TableCellWidth { Width = (widthVal * 50).ToString(), Type = TableWidthUnitValues.Pct });
            else
                cellProps.Append(new TableCellWidth { Width = (widthVal * 15).ToString(), Type = TableWidthUnitValues.Dxa });
        }
        else
        {
            cellProps.Append(new TableCellWidth { Type = TableWidthUnitValues.Auto });
        }
        
        // Kolor tła
        var bgColor = ExtractColor(style, @"background(?:-color)?:\s*");
        if (bgColor != null)
        {
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = bgColor });
        }
        
        // Wyrównanie pionowe
        var vAlignMatch = Regex.Match(style, @"vertical-align:\s*(top|middle|bottom)");
        if (vAlignMatch.Success)
        {
            var vAlign = vAlignMatch.Groups[1].Value switch
            {
                "middle" or "center" => TableVerticalAlignmentValues.Center,
                "bottom" => TableVerticalAlignmentValues.Bottom,
                _ => TableVerticalAlignmentValues.Top
            };
            cellProps.Append(new TableCellVerticalAlignment { Val = vAlign });
        }
        
        // Padding
        var paddingMatch = Regex.Match(style, @"padding:\s*([\d.]+)px(?:\s+([\d.]+)px)?(?:\s+([\d.]+)px)?(?:\s+([\d.]+)px)?");
        if (paddingMatch.Success)
        {
            var top = int.Parse(paddingMatch.Groups[1].Value);
            var right = paddingMatch.Groups[2].Success ? int.Parse(paddingMatch.Groups[2].Value) : top;
            var bottom = paddingMatch.Groups[3].Success ? int.Parse(paddingMatch.Groups[3].Value) : top;
            var left = paddingMatch.Groups[4].Success ? int.Parse(paddingMatch.Groups[4].Value) : right;
            
            cellProps.Append(new TableCellMargin(
                new TopMargin { Width = PxToTwips(top).ToString(), Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = PxToTwips(left).ToString(), Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = PxToTwips(bottom).ToString(), Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = PxToTwips(right).ToString(), Type = TableWidthUnitValues.Dxa }
            ));
        }

        // NoWrap
        if (style.Contains("white-space:nowrap") || style.Contains("white-space: nowrap"))
        {
            cellProps.Append(new NoWrap());
        }

        // Writing mode
        if (style.Contains("writing-mode:vertical-rl"))
        {
            cellProps.Append(new TextDirection { Val = TextDirectionValues.TopToBottomRightToLeft });
        }
        else if (style.Contains("writing-mode:vertical-lr"))
        {
            cellProps.Append(new TextDirection { Val = TextDirectionValues.BottomToTopLeftToRight });
        }
    }

    /// <summary>
    /// Aplikuje obramowania do komórki
    /// </summary>
    private void ApplyCellBorders(TableCellProperties cellProps, string style)
    {
        if (string.IsNullOrEmpty(style)) return;
        
        // Parsuj poszczególne strony
        var borders = new TableCellBorders();
        bool hasBorders = false;
        
        var sides = new[] { ("border-top", typeof(TopBorder)), ("border-bottom", typeof(BottomBorder)), 
                            ("border-left", typeof(LeftBorder)), ("border-right", typeof(RightBorder)) };
        
        foreach (var (prefix, borderType) in sides)
        {
            var match = Regex.Match(style, $@"{Regex.Escape(prefix)}:\s*([\d.]+)px\s+(\w+)\s+#?([a-fA-F0-9]{{3,6}})");
            if (match.Success)
            {
                var size = (uint)(double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 8);
                var bStyle = ParseBorderStyle(match.Groups[2].Value);
                var color = NormalizeColor(match.Groups[3].Value);
                
                var border = (BorderType)Activator.CreateInstance(borderType)!;
                border.Val = bStyle;
                border.Size = size;
                border.Color = color;
                borders.Append(border);
                hasBorders = true;
            }
            else if (style.Contains($"{prefix}:none") || style.Contains($"{prefix}: none"))
            {
                var border = (BorderType)Activator.CreateInstance(borderType)!;
                border.Val = BorderValues.None;
                border.Size = 0;
                borders.Append(border);
                hasBorders = true;
            }
        }
        
        // Parsuj border shorthand
        if (!hasBorders)
        {
            var borderAll = Regex.Match(style, @"(?<![a-z-])border:\s*([\d.]+)px\s+(\w+)\s+#?([a-fA-F0-9]{3,6})");
            if (borderAll.Success)
            {
                var size = (uint)(double.Parse(borderAll.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 8);
                var bStyle = ParseBorderStyle(borderAll.Groups[2].Value);
                var color = NormalizeColor(borderAll.Groups[3].Value);
                
                borders.Append(new TopBorder { Val = bStyle, Size = size, Color = color });
                borders.Append(new BottomBorder { Val = bStyle, Size = size, Color = color });
                borders.Append(new LeftBorder { Val = bStyle, Size = size, Color = color });
                borders.Append(new RightBorder { Val = bStyle, Size = size, Color = color });
                hasBorders = true;
            }
        }
        
        if (hasBorders)
            cellProps.Append(borders);
    }

    /// <summary>
    /// Parsuje styl obramowania CSS na wartość OpenXML
    /// </summary>
    private BorderValues ParseBorderStyle(string cssStyle) => cssStyle.ToLower() switch
    {
        "solid" => BorderValues.Single,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "none" => BorderValues.None,
        _ => BorderValues.Single
    };

    /// <summary>
    /// Normalizuje kolor (3 znaki na 6)
    /// </summary>
    private string NormalizeColor(string color)
    {
        if (color.Length == 3)
            return $"{color[0]}{color[0]}{color[1]}{color[1]}{color[2]}{color[2]}";
        return color;
    }

    /// <summary>
    /// Wyciąga kolor z CSS (obsługuje hex i rgb)
    /// </summary>
    private string? ExtractColor(string style, string prefix)
    {
        // Najpierw hex
        var hexMatch = Regex.Match(style, $@"{prefix}#?([a-fA-F0-9]{{3,6}})");
        if (hexMatch.Success)
            return NormalizeColor(hexMatch.Groups[1].Value);
        
        // Potem rgb()
        var rgbMatch = Regex.Match(style, $@"{prefix}rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)");
        if (rgbMatch.Success)
        {
            var r = int.Parse(rgbMatch.Groups[1].Value);
            var g = int.Parse(rgbMatch.Groups[2].Value);
            var b = int.Parse(rgbMatch.Groups[3].Value);
            return $"{r:X2}{g:X2}{b:X2}";
        }
        
        // rgba()
        var rgbaMatch = Regex.Match(style, $@"{prefix}rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*[\d.]+\s*\)");
        if (rgbaMatch.Success)
        {
            var r = int.Parse(rgbaMatch.Groups[1].Value);
            var g = int.Parse(rgbaMatch.Groups[2].Value);
            var b = int.Parse(rgbaMatch.Groups[3].Value);
            return $"{r:X2}{g:X2}{b:X2}";
        }
        
        return null;
    }

    /// <summary>
    /// Konwertuje obraz na Paragraph z obrazem - z dokładnym odwzorowaniem wymiarów
    /// </summary>
    private Paragraph? ConvertImageElement(HtmlNode node)
    {
        var src = node.GetAttributeValue("src", "");
        if (string.IsNullOrEmpty(src)) return null;

        if (!src.StartsWith("data:")) return null;
        
        var match = Regex.Match(src, @"data:([^;]+);base64,(.+)");
        if (!match.Success) return null;
        
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

    /// <summary>
    /// Tworzy paragraf z obrazem z zachowaniem oryginalnych wymiarów EMU
    /// </summary>
    private Paragraph CreateImageParagraph(byte[] imageBytes, string contentType, HtmlNode node)
    {
        var imagePart = contentType switch
        {
            "image/png" => _mainPart!.AddImagePart(ImagePartType.Png),
            "image/gif" => _mainPart!.AddImagePart(ImagePartType.Gif),
            "image/bmp" => _mainPart!.AddImagePart(ImagePartType.Bmp),
            "image/svg+xml" => _mainPart!.AddImagePart(ImagePartType.Svg),
            _ => _mainPart!.AddImagePart(ImagePartType.Jpeg)
        };

        using var stream = new MemoryStream(imageBytes);
        imagePart.FeedData(stream);

        var relationshipId = _mainPart!.GetIdOfPart(imagePart);

        // Próbuj najpierw użyć oryginalnych wymiarów EMU (zachowanych z DOCX)
        long widthEmu, heightEmu;
        
        var emuWidthAttr = node.GetAttributeValue("data-width-emu", "");
        var emuHeightAttr = node.GetAttributeValue("data-height-emu", "");
        
        if (!string.IsNullOrEmpty(emuWidthAttr) && !string.IsNullOrEmpty(emuHeightAttr) &&
            long.TryParse(emuWidthAttr, out var origWidthEmu) && long.TryParse(emuHeightAttr, out var origHeightEmu))
        {
            widthEmu = origWidthEmu;
            heightEmu = origHeightEmu;
        }
        else
        {
            // Fallback: parsuj ze stylu CSS
            var style = node.GetAttributeValue("style", "");
            var widthMatch = Regex.Match(style, @"width:\s*(\d+)px");
            var heightMatch = Regex.Match(style, @"height:\s*(\d+)px");
            
            var width = widthMatch.Success ? int.Parse(widthMatch.Groups[1].Value) : 200;
            var height = heightMatch.Success ? int.Parse(heightMatch.Groups[1].Value) : (int)(width * 0.75);
            
            widthEmu = width * 9525;
            heightEmu = height * 9525;
        }

        // Ogranicz do max 15cm szerokości (typowa strona A4 minus marginesy)
        var maxWidthEmu = 5400000L; // ~15cm
        if (widthEmu > maxWidthEmu)
        {
            var scale = (double)maxWidthEmu / widthEmu;
            widthEmu = maxWidthEmu;
            heightEmu = (long)(heightEmu * scale);
        }

        _imageCounter++;

        var element = new Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmu, Cy = heightEmu },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = (uint)_imageCounter, Name = $"Image{_imageCounter}" },
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
    /// Konwertuje link na Paragraph z hiperłączem
    /// </summary>
    private Paragraph ConvertAnchorElement(HtmlNode node)
    {
        var para = new Paragraph();
        var href = node.GetAttributeValue("href", "#");

        try
        {
            var relationshipId = _mainPart!.AddHyperlinkRelationship(new Uri(href, UriKind.RelativeOrAbsolute), true).Id;
            
            var hyperlink = new Hyperlink { Id = relationshipId };
            
            foreach (var child in node.ChildNodes)
            {
                var runs = CreateRunsFromNode(child);
                foreach (var run in runs)
                {
                    run.RunProperties ??= new RunProperties();
                    if (!run.RunProperties.Elements<Color>().Any())
                        run.RunProperties.Append(new Color { Val = "0563C1" });
                    if (!run.RunProperties.Elements<Underline>().Any())
                        run.RunProperties.Append(new Underline { Val = UnderlineValues.Single });
                    run.RunProperties.Append(new RunStyle { Val = "Hyperlink" });
                    hyperlink.Append(run);
                }
            }

            para.Append(hyperlink);
        }
        catch
        {
            // Jeśli URI jest nieprawidłowy, dodaj jako zwykły tekst
            AppendInlineContent(para, node);
        }
        
        return para;
    }

    private Paragraph ConvertInlineElement(HtmlNode node)
    {
        var para = new Paragraph();
        AppendInlineContent(para, node);
        return para;
    }

    private Paragraph ConvertBlockquoteElement(HtmlNode node)
    {
        var para = new Paragraph();
        var props = new ParagraphProperties();
        props.Append(new Indentation { Left = "720" });
        props.Append(new ParagraphBorders(
            new LeftBorder { Val = BorderValues.Single, Size = 24, Color = "CCCCCC", Space = 4 }
        ));
        para.Append(props);
        AppendInlineContent(para, node);
        return para;
    }

    /// <summary>
    /// Tworzy paragraf stylizowany na linię horyzontalną
    /// </summary>
    private Paragraph CreateHorizontalRule()
    {
        var para = new Paragraph();
        var props = new ParagraphProperties();
        props.Append(new ParagraphBorders(
            new BottomBorder { Val = BorderValues.Single, Size = 12, Color = "000000", Space = 1 }
        ));
        props.Append(new SpacingBetweenLines { Before = "120", After = "120" });
        para.Append(props);
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

        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any())
        {
            paragraph.Append(new Run(new Text("") { Space = SpaceProcessingModeValues.Preserve }));
        }
    }

    /// <summary>
    /// Tworzy Run z węzła HTML z pełnym odwzorowaniem formatowania
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
                        run.Append(inheritedProps.CloneNode(true));
                    run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    runs.Add(run);
                }
                break;

            case HtmlNodeType.Element:
                var tagName = node.Name.ToLower();
                var newProps = (inheritedProps?.CloneNode(true) as RunProperties) ?? new RunProperties();
                
                switch (tagName)
                {
                    case "strong": case "b":
                        if (!newProps.Elements<Bold>().Any())
                            newProps.Append(new Bold());
                        break;
                    case "em": case "i":
                        if (!newProps.Elements<Italic>().Any())
                            newProps.Append(new Italic());
                        break;
                    case "u":
                        if (!newProps.Elements<Underline>().Any())
                            newProps.Append(new Underline { Val = UnderlineValues.Single });
                        break;
                    case "s": case "strike":
                        if (!newProps.Elements<Strike>().Any())
                            newProps.Append(new Strike());
                        break;
                    case "sub":
                        if (!newProps.Elements<VerticalTextAlignment>().Any())
                            newProps.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
                        break;
                    case "sup":
                        if (!newProps.Elements<VerticalTextAlignment>().Any())
                            newProps.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
                        break;
                    case "br":
                        runs.Add(new Run(new Break()));
                        return runs;
                    case "img":
                        var imgPara = ConvertImageElement(node);
                        if (imgPara != null)
                        {
                            // Wyciągnij run z obrazem
                            foreach (var r in imgPara.Elements<Run>())
                            {
                                runs.Add((Run)r.CloneNode(true));
                            }
                        }
                        return runs;
                    case "a":
                        // Hiperłącze w inline context
                        var href = node.GetAttributeValue("href", "#");
                        try
                        {
                            var relId = _mainPart!.AddHyperlinkRelationship(new Uri(href, UriKind.RelativeOrAbsolute), true).Id;
                            // Dodaj runy z linkiem - w inline kontekście nie możemy dodać Hyperlink do Run
                            // więc po prostu stylizujemy tekst
                            var linkProps = (newProps.CloneNode(true) as RunProperties) ?? new RunProperties();
                            if (!linkProps.Elements<Color>().Any())
                                linkProps.Append(new Color { Val = "0563C1" });
                            if (!linkProps.Elements<Underline>().Any())
                                linkProps.Append(new Underline { Val = UnderlineValues.Single });
                            
                            foreach (var child in node.ChildNodes)
                                runs.AddRange(CreateRunsFromNode(child, linkProps));
                        }
                        catch
                        {
                            foreach (var child in node.ChildNodes)
                                runs.AddRange(CreateRunsFromNode(child, newProps));
                        }
                        return runs;
                }

                // Parsuj style inline
                var style = node.GetAttributeValue("style", "");
                ApplyRunStyle(newProps, style);

                foreach (var child in node.ChildNodes)
                {
                    runs.AddRange(CreateRunsFromNode(child, newProps));
                }
                break;
        }

        return runs;
    }

    /// <summary>
    /// Aplikuje styl CSS do ParagraphProperties z pełnym parsowaniem
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

        // Wcięcia
        var indentation = new Indentation();
        bool hasIndent = false;
        
        var marginLeftMatch = Regex.Match(style, @"margin-left:\s*(\d+)px");
        if (marginLeftMatch.Success)
        {
            indentation.Left = PxToTwips(int.Parse(marginLeftMatch.Groups[1].Value)).ToString();
            hasIndent = true;
        }
        
        var marginRightMatch = Regex.Match(style, @"margin-right:\s*(\d+)px");
        if (marginRightMatch.Success)
        {
            indentation.Right = PxToTwips(int.Parse(marginRightMatch.Groups[1].Value)).ToString();
            hasIndent = true;
        }
        
        var textIndentMatch = Regex.Match(style, @"text-indent:\s*(-?\d+)px");
        if (textIndentMatch.Success)
        {
            var indent = int.Parse(textIndentMatch.Groups[1].Value);
            if (indent < 0)
            {
                indentation.Hanging = PxToTwips(Math.Abs(indent)).ToString();
            }
            else
            {
                indentation.FirstLine = PxToTwips(indent).ToString();
            }
            hasIndent = true;
        }
        
        if (hasIndent)
            props.Append(indentation);

        // Odstępy
        var spacing = new SpacingBetweenLines();
        bool hasSpacing = false;
        
        var marginTopMatch = Regex.Match(style, @"margin-top:\s*([\d.]+)(px|pt)");
        if (marginTopMatch.Success)
        {
            var val = double.Parse(marginTopMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = marginTopMatch.Groups[2].Value;
            if (unit == "px") val = val * 0.75; // px to pt approx
            spacing.Before = ((int)(val * 20)).ToString(); // pt to twips
            hasSpacing = true;
        }
        
        var marginBottomMatch = Regex.Match(style, @"margin-bottom:\s*([\d.]+)(px|pt)");
        if (marginBottomMatch.Success)
        {
            var val = double.Parse(marginBottomMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = marginBottomMatch.Groups[2].Value;
            if (unit == "px") val = val * 0.75;
            spacing.After = ((int)(val * 20)).ToString();
            hasSpacing = true;
        }
        
        var lineHeightMatch = Regex.Match(style, @"line-height:\s*([\d.]+)(pt)?");
        if (lineHeightMatch.Success)
        {
            var val = double.Parse(lineHeightMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = lineHeightMatch.Groups[2].Value;
            if (unit == "pt")
            {
                // Dokładna wartość w pt
                spacing.Line = ((int)(val * 20)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }
            else
            {
                // Mnożnik
                spacing.Line = ((int)(val * 240)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            hasSpacing = true;
        }
        
        if (hasSpacing)
            props.Append(spacing);

        // Kolor tła paragrafu
        var bgColor = ExtractColor(style, @"background(?:-color)?:\s*");
        if (bgColor != null)
        {
            props.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = bgColor });
        }

        // Obramowania paragrafu
        ApplyParagraphBorders(props, style);
    }

    /// <summary>
    /// Aplikuje dodatkowe style do paragrafu (bez nadpisywania StyleId)
    /// </summary>
    private void ApplyParagraphStyleExtras(ParagraphProperties props, string style)
    {
        if (string.IsNullOrEmpty(style)) return;

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
    }

    /// <summary>
    /// Aplikuje obramowania paragrafu z CSS
    /// </summary>
    private void ApplyParagraphBorders(ParagraphProperties props, string style)
    {
        var borders = new ParagraphBorders();
        bool hasBorders = false;

        var borderPatterns = new[]
        {
            ("border-top", new Func<BorderType>(() => new TopBorder())),
            ("border-bottom", new Func<BorderType>(() => new BottomBorder())),
            ("border-left", new Func<BorderType>(() => new LeftBorder())),
            ("border-right", new Func<BorderType>(() => new RightBorder())),
        };

        foreach (var (prefix, createBorder) in borderPatterns)
        {
            var match = Regex.Match(style, $@"{Regex.Escape(prefix)}:\s*([\d.]+)px\s+(\w+)\s+#?([a-fA-F0-9]{{3,6}})");
            if (match.Success)
            {
                var border = createBorder();
                border.Val = ParseBorderStyle(match.Groups[2].Value);
                border.Size = (uint)(double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 8);
                border.Color = NormalizeColor(match.Groups[3].Value);
                border.Space = 4;
                borders.Append(border);
                hasBorders = true;
            }
        }

        if (hasBorders)
            props.Append(borders);
    }

    /// <summary>
    /// Aplikuje styl CSS do RunProperties z pełnym parsowaniem
    /// </summary>
    private void ApplyRunStyle(RunProperties props, string style)
    {
        if (string.IsNullOrEmpty(style)) return;

        // Font-weight
        if (Regex.IsMatch(style, @"font-weight:\s*(bold|[7-9]\d{2})"))
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

        // Text-decoration: obsługa wielu wartości
        var textDecMatch = Regex.Match(style, @"text-decoration:\s*([^;]+)");
        if (textDecMatch.Success)
        {
            var decValue = textDecMatch.Groups[1].Value.ToLower();
            if (decValue.Contains("underline") && !props.Elements<Underline>().Any())
                props.Append(new Underline { Val = UnderlineValues.Single });
            if (decValue.Contains("line-through") && !props.Elements<Strike>().Any())
                props.Append(new Strike());
        }

        // Font-size (obsługa pt, px, em, rem)
        var fontSizeMatch = Regex.Match(style, @"font-size:\s*([\d.]+)(pt|px|em|rem)");
        if (fontSizeMatch.Success)
        {
            var size = double.Parse(fontSizeMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = fontSizeMatch.Groups[2].Value;
            
            double ptSize = unit switch
            {
                "px" => size * 0.75,
                "em" => size * 11, // Assume base 11pt
                "rem" => size * 11,
                _ => size // pt
            };
            
            var halfPoints = ((int)(ptSize * 2)).ToString();
            if (!props.Elements<FontSize>().Any())
                props.Append(new FontSize { Val = halfPoints });
        }
        // font-size: smaller/larger
        else if (style.Contains("font-size:smaller") || style.Contains("font-size: smaller"))
        {
            if (!props.Elements<FontSize>().Any())
                props.Append(new FontSize { Val = "18" }); // ~9pt
        }

        // Font-family
        var fontFamilyMatch = Regex.Match(style, @"font-family:\s*'?([^',;]+)'?");
        if (fontFamilyMatch.Success)
        {
            var fontName = fontFamilyMatch.Groups[1].Value.Trim();
            if (!props.Elements<RunFonts>().Any())
                props.Append(new RunFonts { Ascii = fontName, HighAnsi = fontName });
        }

        // Color (obsługa hex, rgb, rgba)
        var colorVal = ExtractColor(style, @"(?<!background-)color:\s*");
        if (colorVal != null && !props.Elements<Color>().Any())
        {
            props.Append(new Color { Val = colorVal });
        }

        // Background-color
        var bgColorVal = ExtractColor(style, @"background-color:\s*");
        if (bgColorVal != null && !props.Elements<Shading>().Any())
        {
            props.Append(new Shading { Fill = bgColorVal, Val = ShadingPatternValues.Clear });
        }

        // Vertical align
        if (style.Contains("vertical-align:super") && !props.Elements<VerticalTextAlignment>().Any())
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        if (style.Contains("vertical-align:sub") && !props.Elements<VerticalTextAlignment>().Any())
            props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });

        // Letter spacing
        var letterSpacingMatch = Regex.Match(style, @"letter-spacing:\s*([\d.]+)(pt|px)");
        if (letterSpacingMatch.Success && !props.Elements<Spacing>().Any())
        {
            var ls = double.Parse(letterSpacingMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var lsUnit = letterSpacingMatch.Groups[2].Value;
            if (lsUnit == "px") ls = ls * 0.75;
            props.Append(new Spacing { Val = (int)(ls * 20) });
        }

        // Text-transform
        if (style.Contains("text-transform:uppercase") || style.Contains("text-transform: uppercase"))
        {
            if (!props.Elements<Caps>().Any())
                props.Append(new Caps());
        }
        
        // font-variant
        if (style.Contains("font-variant:small-caps") || style.Contains("font-variant: small-caps"))
        {
            if (!props.Elements<SmallCaps>().Any())
                props.Append(new SmallCaps());
        }
    }

    private Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private Paragraph CreatePageBreak()
    {
        return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
    }

    private void SetDocumentMetadata(WordprocessingDocument document, DocumentMetadata metadata)
    {
        // Core Properties (OPC package properties)
        var props = document.PackageProperties;
        
        if (!string.IsNullOrEmpty(metadata.Title))
            props.Title = metadata.Title;
        if (!string.IsNullOrEmpty(metadata.Author))
            props.Creator = metadata.Author;
        if (!string.IsNullOrEmpty(metadata.Subject))
            props.Subject = metadata.Subject;
        if (!string.IsNullOrEmpty(metadata.Keywords))
            props.Keywords = metadata.Keywords;
        if (!string.IsNullOrEmpty(metadata.Description))
            props.Description = metadata.Description;
        if (!string.IsNullOrEmpty(metadata.Category))
            props.Category = metadata.Category;
        if (!string.IsNullOrEmpty(metadata.ContentStatus))
            props.ContentStatus = metadata.ContentStatus;
        if (!string.IsNullOrEmpty(metadata.LastModifiedBy))
            props.LastModifiedBy = metadata.LastModifiedBy;
        if (!string.IsNullOrEmpty(metadata.Revision))
            props.Revision = metadata.Revision;
        if (!string.IsNullOrEmpty(metadata.Version))
            props.Version = metadata.Version;
        
        props.Created = metadata.Created ?? DateTime.UtcNow;
        props.Modified = DateTime.UtcNow;

        // Extended Properties (app.xml — Company, Manager)
        var extPropsPart = document.AddExtendedFilePropertiesPart();
        extPropsPart.Properties = new Properties();

        if (!string.IsNullOrEmpty(metadata.Company))
            extPropsPart.Properties.Company = new Company(metadata.Company);
        if (!string.IsNullOrEmpty(metadata.Manager))
            extPropsPart.Properties.Manager = new Manager(metadata.Manager);

        extPropsPart.Properties.Application = new DocumentFormat.OpenXml.ExtendedProperties.Application("Doc2 Importer");
        extPropsPart.Properties.Save();
    }

    /// <summary>
    /// Dodaje ustawienia strony z dokładnymi marginesami
    /// </summary>
    private void AddPageSettings(Body body, HeaderFooterContent? header = null, HeaderFooterContent? footer = null)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }
        
        // A4
        if (!sectionProps.Elements<PageSize>().Any())
        {
            sectionProps.Append(new PageSize 
            { 
                Width = 11906,
                Height = 16838
            });
        }
        
        var headerHeightTwips = header != null ? (int)(header.Height * 1440 / 2.54) : 720;
        var footerHeightTwips = footer != null ? (int)(footer.Height * 1440 / 2.54) : 720;
        
        var topMargin = Math.Max(1440, headerHeightTwips + 720);
        var bottomMargin = Math.Max(1440, footerHeightTwips + 720);
        
        if (!sectionProps.Elements<PageMargin>().Any())
        {
            sectionProps.Append(new PageMargin
            {
                Top = topMargin,
                Right = 1440,
                Bottom = bottomMargin,
                Left = 1440,
                Header = 720,
                Footer = 720
            });
        }
    }

    private int PxToTwips(int px) => (int)(px / 96.0 * 1440);
}

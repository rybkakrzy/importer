using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImporterApi.Models;

namespace ImporterApi.Services;

/// <summary>
/// Serwis do konwersji dokumentów DOCX na HTML
/// Własna implementacja parsera OpenXML z wysoką dokładnością odwzorowania stylów
/// </summary>
public class DocxToHtmlConverter
{
    private readonly Dictionary<string, DocumentImage> _images = new();
    private readonly Dictionary<string, string> _styles = new();
    private readonly Dictionary<string, Style> _rawStyles = new();
    private readonly List<DocumentStyle> _documentStyles = new();
    private int _imageCounter = 0;
    private NumberingDefinitionsPart? _numberingPart;
    private ThemePart? _themePart;

    /// <summary>
    /// Konwertuje plik DOCX na HTML
    /// </summary>
    public DocumentContent Convert(Stream docxStream)
    {
        _images.Clear();
        _styles.Clear();
        _rawStyles.Clear();
        _documentStyles.Clear();
        _imageCounter = 0;
        _numberingPart = null;
        _themePart = null;

        using var document = WordprocessingDocument.Open(docxStream, false);
        
        // Załaduj części pomocnicze
        _numberingPart = document.MainDocumentPart?.NumberingDefinitionsPart;
        _themePart = document.MainDocumentPart?.ThemePart;
        
        // Załaduj style dokumentu
        var stylesLoaded = ExtractDocumentStyles(document);
        
        var content = new DocumentContent
        {
            Metadata = ExtractMetadata(document),
            Html = ConvertBodyToHtml(document),
            Images = _images.Values.ToList(),
            Styles = stylesLoaded.Count > 0 ? stylesLoaded : DefaultWordStyles.GetDefaultStyles(),
            Header = ExtractHeader(document),
            Footer = ExtractFooter(document)
        };

        return content;
    }

    /// <summary>
    /// Wyciąga nagłówek z dokumentu
    /// </summary>
    private HeaderFooterContent? ExtractHeader(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return null;

        var headerPart = mainPart.HeaderParts.FirstOrDefault();
        if (headerPart?.Header == null) return null;

        foreach (var imagePart in headerPart.ImageParts)
        {
            LoadImageFromPart(headerPart, imagePart);
        }

        var html = ConvertHeaderFooterToHtml(headerPart.Header, headerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        var sectionProps = mainPart.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();
        double headerHeight = 1.5;
        
        var pgMar = sectionProps?.Elements<PageMargin>().FirstOrDefault();
        if (pgMar != null)
        {
            int topMargin = 0;
            int headerMargin = 720;
            
            if (pgMar.Top?.Value != null)
            {
                topMargin = Math.Abs(pgMar.Top.Value);
            }
            if (pgMar.Header?.Value != null)
            {
                headerMargin = (int)pgMar.Header.Value;
            }
            
            if (topMargin > headerMargin)
            {
                headerHeight = (topMargin - headerMargin) / 1440.0 * 2.54;
            }
            else
            {
                headerHeight = topMargin / 1440.0 * 2.54;
            }
        }

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, headerHeight))
        };
    }

    /// <summary>
    /// Wyciąga stopkę z dokumentu
    /// </summary>
    private HeaderFooterContent? ExtractFooter(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return null;

        var footerPart = mainPart.FooterParts.FirstOrDefault();
        if (footerPart?.Footer == null) return null;

        foreach (var imagePart in footerPart.ImageParts)
        {
            LoadImageFromPart(footerPart, imagePart);
        }

        var html = ConvertHeaderFooterToHtml(footerPart.Footer, footerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        var sectionProps = mainPart.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();
        double footerHeight = 1.5;
        
        var pgMar = sectionProps?.Elements<PageMargin>().FirstOrDefault();
        if (pgMar != null)
        {
            int bottomMargin = 0;
            int footerMargin = 720;
            
            if (pgMar.Bottom?.Value != null)
            {
                bottomMargin = Math.Abs(pgMar.Bottom.Value);
            }
            if (pgMar.Footer?.Value != null)
            {
                footerMargin = (int)pgMar.Footer.Value;
            }
            
            if (bottomMargin > footerMargin)
            {
                footerHeight = (bottomMargin - footerMargin) / 1440.0 * 2.54;
            }
            else
            {
                footerHeight = bottomMargin / 1440.0 * 2.54;
            }
        }

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, footerHeight))
        };
    }

    /// <summary>
    /// Konwertuje zawartość nagłówka/stopki na HTML
    /// </summary>
    private string ConvertHeaderFooterToHtml(OpenXmlCompositeElement headerFooter, OpenXmlPart part, WordprocessingDocument document)
    {
        var html = new StringBuilder();

        foreach (var element in headerFooter.Elements())
        {
            if (element is Paragraph para)
            {
                html.Append(ConvertParagraphToHtml(para, document, part));
            }
            else if (element is Table table)
            {
                html.Append(ConvertTableToHtml(table, document, part));
            }
        }

        return html.ToString();
    }

    /// <summary>
    /// Wyciąga metadane z dokumentu
    /// </summary>
    private DocumentMetadata ExtractMetadata(WordprocessingDocument document)
    {
        var metadata = new DocumentMetadata();

        var coreProps = document.PackageProperties;
        if (coreProps != null)
        {
            metadata.Title = coreProps.Title;
            metadata.Author = coreProps.Creator;
            metadata.Subject = coreProps.Subject;
            metadata.Created = coreProps.Created;
            metadata.Modified = coreProps.Modified;
        }

        var body = document.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            var text = body.InnerText;
            metadata.WordCount = text.Split(new[] { ' ', '\t', '\n', '\r' }, 
                StringSplitOptions.RemoveEmptyEntries).Length;
        }

        return metadata;
    }

    /// <summary>
    /// Konwertuje ciało dokumentu na HTML z prawidłowym grupowaniem list
    /// </summary>
    private string ConvertBodyToHtml(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return "<div></div>";

        LoadDocumentStyles(document);
        LoadDocumentImages(document);

        var html = new StringBuilder();
        html.Append("<div class=\"document-content\">");

        var elements = body.Elements().ToList();
        int i = 0;
        while (i < elements.Count)
        {
            var element = elements[i];
            
            if (element is Paragraph p && IsListParagraph(p))
            {
                // Zbierz kolejne elementy listy i owijaj w <ul>/<ol>
                html.Append(ConvertConsecutiveListItems(elements, ref i, document));
            }
            else
            {
                html.Append(ConvertElementToHtml(element, document));
                i++;
            }
        }

        html.Append("</div>");
        return html.ToString();
    }

    /// <summary>
    /// Sprawdza czy paragraf jest elementem listy
    /// </summary>
    private bool IsListParagraph(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        return numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0;
    }

    /// <summary>
    /// Konwertuje kolejne elementy listy na prawidłowy HTML z zagnieżdżaniem
    /// </summary>
    private string ConvertConsecutiveListItems(List<OpenXmlElement> elements, ref int index, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        
        var firstPara = (Paragraph)elements[index];
        var listType = GetListType(firstPara.ParagraphProperties?.NumberingProperties, document);
        var firstLevel = GetListLevel(firstPara);
        
        html.Append(listType == "ol" ? "<ol style=\"margin:0;\">" : "<ul style=\"margin:0;\">");
        
        while (index < elements.Count)
        {
            if (elements[index] is not Paragraph p || !IsListParagraph(p))
                break;
            
            var currentLevel = GetListLevel(p);
            
            if (currentLevel > firstLevel)
            {
                // Zagnieżdżona lista
                var lastLi = "</li>";
                html.Length -= lastLi.Length; // Usuń ostatnie </li> - zagnieżdżona lista idzie wewnątrz
                html.Append(ConvertConsecutiveListItems(elements, ref index, document));
                html.Append("</li>");
            }
            else if (currentLevel < firstLevel)
            {
                break;
            }
            else
            {
                var cssStyle = GetParagraphStyle(p.ParagraphProperties);
                var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && _styles.TryGetValue(styleId, out var styleCss))
                {
                    cssStyle = styleCss + cssStyle;
                }
                
                html.Append($"<li style=\"{cssStyle}\">");
                
                foreach (var child in p.Elements())
                {
                    switch (child)
                    {
                        case Run run:
                            html.Append(ConvertRunToHtml(run, document));
                            break;
                        case Hyperlink hyperlink:
                            html.Append(ConvertHyperlinkToHtml(hyperlink, document));
                            break;
                        case SimpleField simpleField:
                            html.Append(ConvertSimpleFieldToHtml(simpleField));
                            break;
                    }
                }
                
                if (!p.Elements<Run>().Any() && !p.Elements<Hyperlink>().Any())
                {
                    html.Append("&nbsp;");
                }
                
                html.Append("</li>");
                index++;
            }
        }
        
        html.Append(listType == "ol" ? "</ol>" : "</ul>");
        return html.ToString();
    }

    /// <summary>
    /// Pobiera poziom zagnieżdżenia listy
    /// </summary>
    private int GetListLevel(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        return numPr?.NumberingLevelReference?.Val?.Value ?? 0;
    }

    /// <summary>
    /// Ładuje style z dokumentu z rozwiązywaniem dziedziczenia
    /// </summary>
    private void LoadDocumentStyles(WordprocessingDocument document)
    {
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return;

        // Załaduj surowe style
        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value != null)
            {
                _rawStyles[style.StyleId.Value] = style;
            }
        }

        // Konwertuj na CSS z rozwiązywaniem dziedziczenia (BasedOn)
        foreach (var kvp in _rawStyles)
        {
            var css = ConvertStyleToCssWithInheritance(kvp.Value);
            _styles[kvp.Key] = css;
        }
    }

    /// <summary>
    /// Konwertuje styl na CSS z rozwiązywaniem dziedziczenia (BasedOn)
    /// </summary>
    private string ConvertStyleToCssWithInheritance(Style style, HashSet<string>? visited = null)
    {
        visited ??= new HashSet<string>();
        
        var styleId = style.StyleId?.Value;
        if (styleId != null && visited.Contains(styleId))
            return string.Empty;
        
        if (styleId != null)
            visited.Add(styleId);

        var css = new StringBuilder();
        
        // Najpierw zastosuj styl bazowy
        var basedOn = style.BasedOn?.Val?.Value;
        if (basedOn != null && _rawStyles.TryGetValue(basedOn, out var baseStyle))
        {
            css.Append(ConvertStyleToCssWithInheritance(baseStyle, visited));
        }
        
        // Nadpisz właściwościami tego stylu
        var runProps = style.StyleRunProperties;
        if (runProps != null)
        {
            css.Append(ConvertRunPropertiesToCss(runProps));
        }

        var paraProps = style.StyleParagraphProperties;
        if (paraProps != null)
        {
            css.Append(ConvertParagraphPropertiesToCss(paraProps));
        }

        return css.ToString();
    }

    /// <summary>
    /// Wyciąga style dokumentu do modelu DocumentStyle z pełnym dziedziczeniem
    /// </summary>
    private List<DocumentStyle> ExtractDocumentStyles(WordprocessingDocument document)
    {
        var result = new List<DocumentStyle>();
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return result;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.Type?.Value != StyleValues.Paragraph) continue;
            if (style.StyleId?.Value == null) continue;

            var styleName = style.StyleName?.Val?.Value ?? style.StyleId.Value;
            
            var semiHidden = style.SemiHidden;
            if (semiHidden != null && semiHidden.Val == null)
                continue;

            var docStyle = new DocumentStyle
            {
                Id = style.StyleId.Value,
                Name = TranslateStyleName(styleName),
                Type = "paragraph",
                BasedOn = style.BasedOn?.Val?.Value,
                NextStyle = style.NextParagraphStyle?.Val?.Value
            };

            var runProps = style.StyleRunProperties;
            if (runProps != null)
            {
                var font = runProps.RunFonts;
                if (font != null)
                {
                    docStyle.FontFamily = font.Ascii?.Value ?? 
                                          font.HighAnsi?.Value ?? 
                                          font.ComplexScript?.Value;
                }

                if (runProps.FontSize?.Val?.Value != null &&
                    double.TryParse(runProps.FontSize.Val.Value, out var fontSize))
                {
                    docStyle.FontSize = fontSize / 2;
                }

                var color = runProps.Color?.Val?.Value;
                if (!string.IsNullOrEmpty(color) && color != "auto")
                {
                    docStyle.Color = "#" + color;
                }
                else if (runProps.Color?.ThemeColor?.Value != null)
                {
                    var themeColor = ResolveThemeColor(runProps.Color.ThemeColor.Value);
                    if (themeColor != null)
                        docStyle.Color = themeColor;
                }

                docStyle.IsBold = runProps.Bold != null && 
                                  (runProps.Bold.Val == null || runProps.Bold.Val.Value);
                docStyle.IsItalic = runProps.Italic != null && 
                                    (runProps.Italic.Val == null || runProps.Italic.Val.Value);
                docStyle.IsUnderline = runProps.Underline != null && 
                                       runProps.Underline.Val?.Value != UnderlineValues.None;
            }

            var paraProps = style.StyleParagraphProperties;
            if (paraProps != null)
            {
                var justification = paraProps.Justification?.Val;
                if (justification != null && justification.HasValue)
                {
                    var justVal = justification.Value;
                    if (justVal == JustificationValues.Left) docStyle.Alignment = "left";
                    else if (justVal == JustificationValues.Center) docStyle.Alignment = "center";
                    else if (justVal == JustificationValues.Right) docStyle.Alignment = "right";
                    else if (justVal == JustificationValues.Both) docStyle.Alignment = "justify";
                }

                var spacing = paraProps.SpacingBetweenLines;
                if (spacing != null)
                {
                    if (spacing.Before?.Value != null &&
                        int.TryParse(spacing.Before.Value, out var before))
                    {
                        docStyle.SpaceBefore = before / 20.0;
                    }
                    if (spacing.After?.Value != null &&
                        int.TryParse(spacing.After.Value, out var after))
                    {
                        docStyle.SpaceAfter = after / 20.0;
                    }
                    if (spacing.Line?.Value != null &&
                        int.TryParse(spacing.Line.Value, out var lineVal))
                    {
                        docStyle.LineSpacing = lineVal / 240.0;
                    }
                }

                var indent = paraProps.Indentation;
                if (indent != null)
                {
                    if (indent.Left?.Value != null &&
                        int.TryParse(indent.Left.Value, out var left))
                    {
                        docStyle.LeftIndent = TwipsToCm(left);
                    }
                    if (indent.Right?.Value != null &&
                        int.TryParse(indent.Right.Value, out var right))
                    {
                        docStyle.RightIndent = TwipsToCm(right);
                    }
                    if (indent.FirstLine?.Value != null &&
                        int.TryParse(indent.FirstLine.Value, out var firstLine))
                    {
                        docStyle.FirstLineIndent = TwipsToCm(firstLine);
                    }
                }

                if (paraProps.OutlineLevel?.Val?.Value != null)
                {
                    docStyle.OutlineLevel = paraProps.OutlineLevel.Val.Value + 1;
                }
            }

            result.Add(docStyle);
        }

        return result.Count > 0 ? result : DefaultWordStyles.GetDefaultStyles();
    }

    /// <summary>
    /// Tłumaczy nazwę stylu na polski
    /// </summary>
    private string TranslateStyleName(string name)
    {
        return name.ToLower() switch
        {
            "normal" => "Normalny",
            "heading 1" or "heading1" => "Nagłówek 1",
            "heading 2" or "heading2" => "Nagłówek 2",
            "heading 3" or "heading3" => "Nagłówek 3",
            "heading 4" or "heading4" => "Nagłówek 4",
            "heading 5" or "heading5" => "Nagłówek 5",
            "heading 6" or "heading6" => "Nagłówek 6",
            "title" => "Tytuł",
            "subtitle" => "Podtytuł",
            "quote" => "Cytat",
            "intense quote" => "Cytat intensywny",
            "list paragraph" => "Akapit listy",
            "no spacing" => "Bez odstępów",
            "toc heading" => "Nagłówek spisu treści",
            _ => name
        };
    }

    private double TwipsToCm(int twips) => Math.Round(twips / 567.0, 2);

    /// <summary>
    /// Ładuje obrazy z dokumentu
    /// </summary>
    private void LoadDocumentImages(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return;

        foreach (var imagePart in mainPart.ImageParts)
            LoadImageFromPart(mainPart, imagePart);

        foreach (var headerPart in mainPart.HeaderParts)
            foreach (var imagePart in headerPart.ImageParts)
                LoadImageFromPart(headerPart, imagePart);

        foreach (var footerPart in mainPart.FooterParts)
            foreach (var imagePart in footerPart.ImageParts)
                LoadImageFromPart(footerPart, imagePart);
    }

    private void LoadImageFromPart(OpenXmlPart part, ImagePart imagePart)
    {
        var relationshipId = part.GetIdOfPart(imagePart);
        if (_images.ContainsKey(relationshipId)) return;
        
        using var stream = imagePart.GetStream();
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        
        _images[relationshipId] = new DocumentImage
        {
            Id = relationshipId,
            ContentType = imagePart.ContentType,
            Base64Data = System.Convert.ToBase64String(memoryStream.ToArray())
        };
    }

    /// <summary>
    /// Konwertuje element OpenXML na HTML
    /// </summary>
    private string ConvertElementToHtml(OpenXmlElement element, WordprocessingDocument document)
    {
        return element switch
        {
            Paragraph p => ConvertParagraphToHtml(p, document),
            Table t => ConvertTableToHtml(t, document),
            SdtBlock sdt => ConvertSdtBlockToHtml(sdt, document),
            _ => string.Empty
        };
    }

    /// <summary>
    /// Konwertuje paragraf na HTML z pełnym odwzorowaniem stylów
    /// </summary>
    private string ConvertParagraphToHtml(Paragraph paragraph, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var html = new StringBuilder();
        var paraProps = paragraph.ParagraphProperties;
        
        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
        var headingLevel = GetHeadingLevel(styleId);
        
        // Listy powinny być obsługiwane przez ConvertConsecutiveListItems
        var numPr = paraProps?.NumberingProperties;
        var isListItem = numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0;
        
        var tag = headingLevel > 0 ? $"h{headingLevel}" : "p";
        
        // Buduj CSS: najpierw styl z definicji (z dziedziczeniem), potem inline
        var cssBuilder = new StringBuilder();
        if (styleId != null && _styles.TryGetValue(styleId, out var styleCss))
        {
            cssBuilder.Append(styleCss);
        }
        cssBuilder.Append(GetParagraphStyle(paraProps));
        
        // Obramowanie paragrafu
        var borderCss = GetParagraphBorderCss(paraProps);
        if (!string.IsNullOrEmpty(borderCss))
        {
            cssBuilder.Append(borderCss);
        }
        
        var cssStyle = cssBuilder.ToString();

        if (isListItem)
        {
            html.Append($"<li style=\"{cssStyle}\">");
        }
        else
        {
            html.Append($"<{tag} style=\"{cssStyle}\">");
        }

        // Obsługa złożonych pól (FieldChar Begin/Separate/End)
        var hasComplexField = paragraph.Descendants<FieldChar>().Any();
        if (hasComplexField)
        {
            html.Append(ConvertComplexFieldParagraphContent(paragraph, document, sourcePart));
        }
        else
        {
            foreach (var child in paragraph.Elements())
            {
                switch (child)
                {
                    case Run run:
                        html.Append(ConvertRunToHtml(run, document, sourcePart));
                        break;
                    case Hyperlink hyperlink:
                        html.Append(ConvertHyperlinkToHtml(hyperlink, document));
                        break;
                    case SimpleField simpleField:
                        html.Append(ConvertSimpleFieldToHtml(simpleField));
                        break;
                }
            }
        }

        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any() && !paragraph.Elements<SimpleField>().Any())
        {
            html.Append("&nbsp;");
        }

        html.Append(isListItem ? "</li>" : $"</{tag}>");
        return html.ToString();
    }

    /// <summary>
    /// Konwertuje zawartość paragrafu ze złożonymi kodami pól
    /// </summary>
    private string ConvertComplexFieldParagraphContent(Paragraph paragraph, WordprocessingDocument document, OpenXmlPart? sourcePart)
    {
        var html = new StringBuilder();
        bool inField = false;
        string fieldInstruction = "";
        bool fieldSeparated = false;

        foreach (var child in paragraph.Elements())
        {
            if (child is not Run run)
            {
                if (child is Hyperlink hyperlink)
                    html.Append(ConvertHyperlinkToHtml(hyperlink, document));
                else if (child is SimpleField simpleField)
                    html.Append(ConvertSimpleFieldToHtml(simpleField));
                continue;
            }

            var fieldChar = run.GetFirstChild<FieldChar>();
            if (fieldChar != null)
            {
                var fctVal = fieldChar.FieldCharType?.Value;
                if (fctVal == FieldCharValues.Begin)
                {
                    inField = true;
                    fieldInstruction = "";
                    fieldSeparated = false;
                }
                else if (fctVal == FieldCharValues.Separate)
                {
                    fieldSeparated = true;
                    // Emit field placeholder based on instruction
                    var instr = fieldInstruction.Trim().ToUpperInvariant();
                    if (instr.Contains("PAGE") && !instr.Contains("NUMPAGES") && !instr.Contains("SECTIONPAGES"))
                    {
                        html.Append("<span class=\"field-page\">{page}</span>");
                    }
                    else if (instr.Contains("NUMPAGES") || instr.Contains("SECTIONPAGES"))
                    {
                        html.Append("<span class=\"field-numpages\">{pages}</span>");
                    }
                    else if (instr.Contains("DATE") || instr.Contains("TIME"))
                    {
                        html.Append($"<span class=\"field-date\">{DateTime.Now:dd.MM.yyyy}</span>");
                    }
                }
                else if (fctVal == FieldCharValues.End)
                {
                    if (!fieldSeparated)
                    {
                        // Pole bez separatora - spróbuj zinterpretować
                        var instrEnd = fieldInstruction.Trim().ToUpperInvariant();
                        if (instrEnd.Contains("PAGE") && !instrEnd.Contains("NUMPAGES"))
                        {
                            html.Append("<span class=\"field-page\">{page}</span>");
                        }
                        else if (instrEnd.Contains("NUMPAGES") || instrEnd.Contains("SECTIONPAGES"))
                        {
                            html.Append("<span class=\"field-numpages\">{pages}</span>");
                        }
                    }
                    inField = false;
                    fieldInstruction = "";
                    fieldSeparated = false;
                }
                continue;
            }

            if (inField)
            {
                var fieldCode = run.GetFirstChild<FieldCode>();
                if (fieldCode != null)
                {
                    fieldInstruction += fieldCode.Text;
                    continue;
                }
                
                // Po separatorze - to jest wyświetlana wartość pola, pomijamy
                if (fieldSeparated) continue;
            }

            // Normalny run
            html.Append(ConvertRunToHtml(run, document, sourcePart));
        }

        return html.ToString();
    }

    /// <summary>
    /// Pobiera CSS obramowania paragrafu
    /// </summary>
    private string GetParagraphBorderCss(ParagraphProperties? props)
    {
        if (props == null) return string.Empty;
        
        var borders = props.ParagraphBorders;
        if (borders == null) return string.Empty;
        
        var css = new StringBuilder();
        
        if (borders.TopBorder?.Val != null && borders.TopBorder.Val.Value != BorderValues.None && borders.TopBorder.Val.Value != BorderValues.Nil)
            css.Append($"border-top:{GetBorderCss(borders.TopBorder)};");
        
        if (borders.BottomBorder?.Val != null && borders.BottomBorder.Val.Value != BorderValues.None && borders.BottomBorder.Val.Value != BorderValues.Nil)
            css.Append($"border-bottom:{GetBorderCss(borders.BottomBorder)};");
        
        if (borders.LeftBorder?.Val != null && borders.LeftBorder.Val.Value != BorderValues.None && borders.LeftBorder.Val.Value != BorderValues.Nil)
            css.Append($"border-left:{GetBorderCss(borders.LeftBorder)};");
        
        if (borders.RightBorder?.Val != null && borders.RightBorder.Val.Value != BorderValues.None && borders.RightBorder.Val.Value != BorderValues.Nil)
            css.Append($"border-right:{GetBorderCss(borders.RightBorder)};");
        
        if (css.Length > 0)
            css.Append("padding:4px 8px;");
        
        return css.ToString();
    }

    /// <summary>
    /// Konwertuje obramowanie OpenXML na CSS border string
    /// </summary>
    private string GetBorderCss(BorderType border)
    {
        var size = border.Size?.Value ?? 4;
        var sizePx = Math.Max(1, size / 8.0);
        var color = border.Color?.Value ?? "000000";
        if (color == "auto") color = "000000";
        
        var borderVal = border.Val?.Value;
        string style = "solid";
        if (borderVal == BorderValues.Single) style = "solid";
        else if (borderVal == BorderValues.Double) style = "double";
        else if (borderVal == BorderValues.Dotted) style = "dotted";
        else if (borderVal == BorderValues.Dashed) style = "dashed";
        else if (borderVal == BorderValues.DashSmallGap) style = "dashed";
        else if (borderVal == BorderValues.DotDash) style = "dashed";
        else if (borderVal == BorderValues.Triple) style = "double";
        else if (borderVal == BorderValues.Thick) style = "solid";
        else if (borderVal == BorderValues.ThickThinSmallGap) style = "double";
        else if (borderVal == BorderValues.ThinThickSmallGap) style = "double";
        
        return $"{sizePx:F1}px {style} #{color}";
    }

    private int GetHeadingLevel(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return 0;
        var match = Regex.Match(styleId, @"Heading(\d)|Nagwek(\d)", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var level = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
            return int.Parse(level);
        }
        return 0;
    }

    /// <summary>
    /// Pobiera typ listy (ol/ul) na podstawie definicji numeracji w dokumencie
    /// </summary>
    private string GetListType(NumberingProperties? numPr, WordprocessingDocument document)
    {
        if (numPr == null || _numberingPart?.Numbering == null) return "ul";
        
        var numId = numPr.NumberingId?.Val?.Value;
        if (numId == null) return "ul";
        
        var level = numPr.NumberingLevelReference?.Val?.Value ?? 0;
        
        var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return "ul";
        
        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return "ul";
        
        var abstractNum = _numberingPart.Numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null) return "ul";
        
        var levelDef = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == level);
        if (levelDef == null) return "ul";
        
        var numFmt = levelDef.NumberingFormat?.Val?.Value;
        
        if (numFmt == NumberFormatValues.Decimal) return "ol";
        if (numFmt == NumberFormatValues.UpperLetter) return "ol";
        if (numFmt == NumberFormatValues.LowerLetter) return "ol";
        if (numFmt == NumberFormatValues.UpperRoman) return "ol";
        if (numFmt == NumberFormatValues.LowerRoman) return "ol";
        if (numFmt == NumberFormatValues.Bullet) return "ul";
        return "ul";
    }

    /// <summary>
    /// Pobiera styl CSS dla paragrafu (tylko właściwości inline)
    /// </summary>
    private string GetParagraphStyle(ParagraphProperties? props)
    {
        if (props == null) return string.Empty;
        return ConvertParagraphPropertiesToCss(props);
    }

    /// <summary>
    /// Konwertuje właściwości paragrafu na CSS z dokładnymi jednostkami
    /// </summary>
    private string ConvertParagraphPropertiesToCss(OpenXmlElement props)
    {
        var css = new StringBuilder();

        var justification = props.Descendants<Justification>().FirstOrDefault();
        if (justification?.Val != null)
        {
            css.Append($"text-align:{GetJustificationAlignment(justification.Val.Value)};");
        }

        var indentation = props.Descendants<Indentation>().FirstOrDefault();
        if (indentation != null)
        {
            if (indentation.Left?.Value != null && int.TryParse(indentation.Left.Value, out var leftVal))
                css.Append($"margin-left:{TwipsToPx(leftVal)}px;");
            if (indentation.Right?.Value != null && int.TryParse(indentation.Right.Value, out var rightVal))
                css.Append($"margin-right:{TwipsToPx(rightVal)}px;");
            if (indentation.FirstLine?.Value != null && int.TryParse(indentation.FirstLine.Value, out var firstLineVal))
                css.Append($"text-indent:{TwipsToPx(firstLineVal)}px;");
            if (indentation.Hanging?.Value != null && int.TryParse(indentation.Hanging.Value, out var hangingVal))
            {
                var hangingPx = TwipsToPx(hangingVal);
                css.Append($"text-indent:-{hangingPx}px;padding-left:{hangingPx}px;");
            }
        }

        var spacing = props.Descendants<SpacingBetweenLines>().FirstOrDefault();
        if (spacing != null)
        {
            if (spacing.Before?.Value != null && int.TryParse(spacing.Before.Value, out var beforeVal))
                css.Append($"margin-top:{(beforeVal / 20.0):F1}pt;");
            if (spacing.After?.Value != null && int.TryParse(spacing.After.Value, out var afterVal))
                css.Append($"margin-bottom:{(afterVal / 20.0):F1}pt;");
            if (spacing.Line?.Value != null && int.TryParse(spacing.Line.Value, out var lineVal))
            {
                var lineRule = spacing.LineRule?.Value;
                if (lineRule == LineSpacingRuleValues.Exact || lineRule == LineSpacingRuleValues.AtLeast)
                {
                    css.Append($"line-height:{(lineVal / 20.0):F1}pt;");
                }
                else
                {
                    css.Append($"line-height:{(lineVal / 240.0):F2};");
                }
            }
        }

        // Kolor tła paragrafu
        var shading = props.Descendants<Shading>().FirstOrDefault();
        if (shading?.Fill?.Value != null && shading.Fill.Value != "auto" && shading.Fill.Value.ToUpper() != "FFFFFF")
        {
            css.Append($"background-color:#{shading.Fill.Value};");
        }

        return css.ToString();
    }

    /// <summary>
    /// Konwertuje Run (fragment tekstu) na HTML z semantycznymi tagami
    /// </summary>
    private string ConvertRunToHtml(Run run, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var html = new StringBuilder();
        var runProps = run.RunProperties;

        // Pobierz CSS bez formatowania obsługiwanego przez tagi HTML
        bool needsBold = false, needsItalic = false, needsUnderline = false, needsStrike = false, needsSup = false, needsSub = false;
        if (runProps != null)
        {
            needsBold = runProps.Bold != null && (runProps.Bold.Val == null || runProps.Bold.Val.Value);
            needsItalic = runProps.Italic != null && (runProps.Italic.Val == null || runProps.Italic.Val.Value);
            needsUnderline = runProps.Underline != null && runProps.Underline.Val?.Value != UnderlineValues.None;
            needsStrike = (runProps.Strike != null && (runProps.Strike.Val == null || runProps.Strike.Val.Value)) ||
                          (runProps.DoubleStrike != null && (runProps.DoubleStrike.Val == null || runProps.DoubleStrike.Val.Value));
            var vertAlign = runProps.VerticalTextAlignment;
            if (vertAlign?.Val != null)
            {
                needsSup = vertAlign.Val.Value == VerticalPositionValues.Superscript;
                needsSub = vertAlign.Val.Value == VerticalPositionValues.Subscript;
            }
        }

        var cleanCss = GetRunStyleClean(runProps);

        html.Append($"<span style=\"{cleanCss}\">");
        if (needsBold) html.Append("<strong>");
        if (needsItalic) html.Append("<em>");
        if (needsUnderline) html.Append("<u>");
        if (needsStrike) html.Append("<s>");
        if (needsSup) html.Append("<sup>");
        if (needsSub) html.Append("<sub>");

        foreach (var child in run.Elements())
        {
            switch (child)
            {
                case Text text:
                    html.Append(EscapeHtml(text.Text));
                    break;
                case Break br:
                    html.Append(br.Type?.Value == BreakValues.Page ? "<div class=\"page-break\"></div>" : "<br/>");
                    break;
                case TabChar _:
                    html.Append("<span style=\"display:inline-block;min-width:2em;\">\t</span>");
                    break;
                case Drawing drawing:
                    html.Append(ConvertDrawingToHtml(drawing, document, sourcePart));
                    break;
                case Picture picture:
                    html.Append(ConvertPictureToHtml(picture, document, sourcePart));
                    break;
                case NoBreakHyphen _:
                    html.Append("&#8209;");
                    break;
                case SoftHyphen _:
                    html.Append("&shy;");
                    break;
                case SymbolChar sym:
                    if (sym.Char?.Value != null)
                    {
                        try { html.Append($"&#x{sym.Char.Value};"); } catch { }
                    }
                    break;
                case LastRenderedPageBreak _:
                    break;
            }
        }

        if (needsSub) html.Append("</sub>");
        if (needsSup) html.Append("</sup>");
        if (needsStrike) html.Append("</s>");
        if (needsUnderline) html.Append("</u>");
        if (needsItalic) html.Append("</em>");
        if (needsBold) html.Append("</strong>");
        html.Append("</span>");
        
        return html.ToString();
    }

    /// <summary>
    /// Pobiera CSS dla Run bez właściwości obsługiwanych przez semantyczne tagi HTML
    /// </summary>
    private string GetRunStyleClean(RunProperties? props)
    {
        if (props == null) return string.Empty;
        var css = new StringBuilder();

        // Rozmiar czcionki
        var fontSize = props.Descendants<FontSize>().FirstOrDefault();
        if (fontSize?.Val != null)
        {
            var size = double.Parse(fontSize.Val.Value, System.Globalization.CultureInfo.InvariantCulture) / 2;
            css.Append($"font-size:{size}pt;");
        }

        // Rodzina czcionki
        var fontFamily = props.Descendants<RunFonts>().FirstOrDefault();
        if (fontFamily != null)
        {
            var name = fontFamily.Ascii?.Value ?? fontFamily.HighAnsi?.Value ?? fontFamily.ComplexScript?.Value;
            if (name != null)
                css.Append($"font-family:'{name}',sans-serif;");
        }

        // Kolor tekstu (z obsługą kolorów motywu)
        var color = props.Descendants<Color>().FirstOrDefault();
        if (color?.Val != null && color.Val.Value != "auto")
        {
            css.Append($"color:#{color.Val.Value};");
        }
        else if (color?.ThemeColor?.Value != null)
        {
            var themeColor = ResolveThemeColor(color.ThemeColor.Value);
            if (themeColor != null) css.Append($"color:{themeColor};");
        }

        // Podświetlenie
        var highlight = props.Descendants<Highlight>().FirstOrDefault();
        if (highlight?.Val != null)
            css.Append($"background-color:{GetHighlightColor(highlight.Val.Value)};");

        // Shading
        var shading = props.Descendants<Shading>().FirstOrDefault();
        if (shading?.Fill?.Value != null && shading.Fill.Value != "auto")
            css.Append($"background-color:#{shading.Fill.Value};");

        // Rozstrzelenie liter
        var spacing = props.Descendants<Spacing>().FirstOrDefault();
        if (spacing?.Val != null)
            css.Append($"letter-spacing:{(spacing.Val.Value / 20.0):F1}pt;");

        // Caps / SmallCaps
        var caps = props.Descendants<Caps>().FirstOrDefault();
        if (caps != null && (caps.Val == null || caps.Val.Value))
            css.Append("text-transform:uppercase;");
        var smallCaps = props.Descendants<SmallCaps>().FirstOrDefault();
        if (smallCaps != null && (smallCaps.Val == null || smallCaps.Val.Value))
            css.Append("font-variant:small-caps;");

        return css.ToString();
    }

    /// <summary>
    /// Konwertuje właściwości Run na CSS (pełne, dla użycia w stylach)
    /// </summary>
    private string ConvertRunPropertiesToCss(OpenXmlElement props)
    {
        var css = new StringBuilder();

        var bold = props.Descendants<Bold>().FirstOrDefault();
        if (bold != null && (bold.Val == null || bold.Val.Value))
            css.Append("font-weight:bold;");

        var italic = props.Descendants<Italic>().FirstOrDefault();
        if (italic != null && (italic.Val == null || italic.Val.Value))
            css.Append("font-style:italic;");

        // text-decoration z obsługą wielu wartości
        var decorations = new List<string>();
        var underline = props.Descendants<Underline>().FirstOrDefault();
        if (underline?.Val != null && underline.Val.Value != UnderlineValues.None)
            decorations.Add("underline");
        var strike = props.Descendants<Strike>().FirstOrDefault();
        if (strike != null && (strike.Val == null || strike.Val.Value))
            decorations.Add("line-through");
        var dStrike = props.Descendants<DoubleStrike>().FirstOrDefault();
        if (dStrike != null && (dStrike.Val == null || dStrike.Val.Value))
            decorations.Add("line-through");
        if (decorations.Count > 0)
            css.Append($"text-decoration:{string.Join(" ", decorations)};");

        var fontSize = props.Descendants<FontSize>().FirstOrDefault();
        if (fontSize?.Val != null)
        {
            var size = double.Parse(fontSize.Val.Value, System.Globalization.CultureInfo.InvariantCulture) / 2;
            css.Append($"font-size:{size}pt;");
        }

        var fontFamily = props.Descendants<RunFonts>().FirstOrDefault();
        if (fontFamily != null)
        {
            var name = fontFamily.Ascii?.Value ?? fontFamily.HighAnsi?.Value ?? fontFamily.ComplexScript?.Value;
            if (name != null)
                css.Append($"font-family:'{name}',sans-serif;");
        }

        var color = props.Descendants<Color>().FirstOrDefault();
        if (color?.Val != null && color.Val.Value != "auto")
            css.Append($"color:#{color.Val.Value};");
        else if (color?.ThemeColor?.Value != null)
        {
            var themeColor = ResolveThemeColor(color.ThemeColor.Value);
            if (themeColor != null) css.Append($"color:{themeColor};");
        }

        var highlight = props.Descendants<Highlight>().FirstOrDefault();
        if (highlight?.Val != null)
            css.Append($"background-color:{GetHighlightColor(highlight.Val.Value)};");

        var shading = props.Descendants<Shading>().FirstOrDefault();
        if (shading?.Fill?.Value != null && shading.Fill.Value != "auto")
            css.Append($"background-color:#{shading.Fill.Value};");

        var vertAlign = props.Descendants<VerticalTextAlignment>().FirstOrDefault();
        if (vertAlign?.Val != null)
        {
            if (vertAlign.Val.Value == VerticalPositionValues.Superscript)
                css.Append("vertical-align:super;font-size:smaller;");
            else if (vertAlign.Val.Value == VerticalPositionValues.Subscript)
                css.Append("vertical-align:sub;font-size:smaller;");
        }

        var spacing = props.Descendants<Spacing>().FirstOrDefault();
        if (spacing?.Val != null)
            css.Append($"letter-spacing:{(spacing.Val.Value / 20.0):F1}pt;");

        var capsProp = props.Descendants<Caps>().FirstOrDefault();
        if (capsProp != null && (capsProp.Val == null || capsProp.Val.Value))
            css.Append("text-transform:uppercase;");
        var smallCapsProp = props.Descendants<SmallCaps>().FirstOrDefault();
        if (smallCapsProp != null && (smallCapsProp.Val == null || smallCapsProp.Val.Value))
            css.Append("font-variant:small-caps;");

        return css.ToString();
    }

    /// <summary>
    /// Rozwiązuje kolor motywu na wartość hex
    /// </summary>
    private string? ResolveThemeColor(ThemeColorValues themeColor)
    {
        if (_themePart?.Theme?.ThemeElements?.ColorScheme == null) return null;
        
        var cs = _themePart.Theme.ThemeElements.ColorScheme;
        
        DocumentFormat.OpenXml.Drawing.Color2Type? c2 = null;
        if (themeColor == ThemeColorValues.Dark1) c2 = cs.Dark1Color;
        else if (themeColor == ThemeColorValues.Light1) c2 = cs.Light1Color;
        else if (themeColor == ThemeColorValues.Dark2) c2 = cs.Dark2Color;
        else if (themeColor == ThemeColorValues.Light2) c2 = cs.Light2Color;
        else if (themeColor == ThemeColorValues.Accent1) c2 = cs.Accent1Color;
        else if (themeColor == ThemeColorValues.Accent2) c2 = cs.Accent2Color;
        else if (themeColor == ThemeColorValues.Accent3) c2 = cs.Accent3Color;
        else if (themeColor == ThemeColorValues.Accent4) c2 = cs.Accent4Color;
        else if (themeColor == ThemeColorValues.Accent5) c2 = cs.Accent5Color;
        else if (themeColor == ThemeColorValues.Accent6) c2 = cs.Accent6Color;
        else if (themeColor == ThemeColorValues.Hyperlink) c2 = cs.Hyperlink;
        else if (themeColor == ThemeColorValues.FollowedHyperlink) c2 = cs.FollowedHyperlinkColor;
        
        if (c2 == null) return null;
        
        var srgb = c2.GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>();
        if (srgb?.Val?.Value != null) return "#" + srgb.Val.Value;
        
        var sysColor = c2.GetFirstChild<DocumentFormat.OpenXml.Drawing.SystemColor>();
        if (sysColor?.LastColor?.Value != null) return "#" + sysColor.LastColor.Value;
        
        return null;
    }

    private string GetHighlightColor(HighlightColorValues value)
    {
        if (value == HighlightColorValues.Yellow) return "#ffff00";
        if (value == HighlightColorValues.Green) return "#00ff00";
        if (value == HighlightColorValues.Cyan) return "#00ffff";
        if (value == HighlightColorValues.Magenta) return "#ff00ff";
        if (value == HighlightColorValues.Blue) return "#0000ff";
        if (value == HighlightColorValues.Red) return "#ff0000";
        if (value == HighlightColorValues.DarkBlue) return "#000080";
        if (value == HighlightColorValues.DarkCyan) return "#008080";
        if (value == HighlightColorValues.DarkGreen) return "#008000";
        if (value == HighlightColorValues.DarkMagenta) return "#800080";
        if (value == HighlightColorValues.DarkRed) return "#800000";
        if (value == HighlightColorValues.DarkYellow) return "#808000";
        if (value == HighlightColorValues.DarkGray) return "#808080";
        if (value == HighlightColorValues.LightGray) return "#c0c0c0";
        if (value == HighlightColorValues.Black) return "#000000";
        return "transparent";
    }

    /// <summary>
    /// Konwertuje hiperłącze na HTML
    /// </summary>
    private string ConvertHyperlinkToHtml(Hyperlink hyperlink, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        
        var relationshipId = hyperlink.Id?.Value;
        string? url = null;

        if (relationshipId != null)
        {
            var rel = document.MainDocumentPart?.HyperlinkRelationships
                .FirstOrDefault(r => r.Id == relationshipId);
            url = rel?.Uri?.ToString();
        }

        html.Append($"<a href=\"{EscapeHtml(url ?? "#")}\" target=\"_blank\" style=\"color:#0563C1;text-decoration:underline;\">");
        
        foreach (var run in hyperlink.Elements<Run>())
        {
            var runProps = run.RunProperties;
            var cssStyle = GetRunStyleClean(runProps);
            
            html.Append($"<span style=\"{cssStyle}\">");
            foreach (var child in run.Elements())
            {
                if (child is Text text)
                    html.Append(EscapeHtml(text.Text));
            }
            html.Append("</span>");
        }

        html.Append("</a>");
        return html.ToString();
    }

    private string ConvertSimpleFieldToHtml(SimpleField simpleField)
    {
        var instruction = simpleField.Instruction?.Value?.Trim().ToUpperInvariant() ?? "";
        
        if (instruction.Contains("PAGE") && !instruction.Contains("NUMPAGES") && !instruction.Contains("SECTIONPAGES"))
            return "<span class=\"field-page\">{page}</span>";
        if (instruction.Contains("NUMPAGES") || instruction.Contains("SECTIONPAGES"))
            return "<span class=\"field-numpages\">{pages}</span>";
        if (instruction.Contains("DATE") || instruction.Contains("TIME"))
            return $"<span class=\"field-date\">{DateTime.Now:dd.MM.yyyy}</span>";
        
        var text = string.Join("", simpleField.Descendants<Text>().Select(t => t.Text));
        return !string.IsNullOrEmpty(text) ? EscapeHtml(text) : "";
    }

    /// <summary>
    /// Konwertuje Drawing (obraz) na HTML z zachowaniem wymiarów i danych EMU
    /// </summary>
    private string ConvertDrawingToHtml(Drawing drawing, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return string.Empty;

        var relationshipId = blip.Embed.Value;
        
        var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        var width = extent?.Cx != null ? EmuToPx(extent.Cx.Value) : 200;
        var height = extent?.Cy != null ? EmuToPx(extent.Cy.Value) : 200;
        var widthEmu = extent?.Cx?.Value ?? (long)(width * 9525);
        var heightEmu = extent?.Cy?.Value ?? (long)(height * 9525);

        string? base64Data = null;
        string? contentType = null;

        if (_images.TryGetValue(relationshipId, out var image))
        {
            base64Data = image.Base64Data;
            contentType = image.ContentType;
        }
        else if (sourcePart != null)
        {
            try
            {
                var imagePart = sourcePart.GetPartById(relationshipId) as ImagePart;
                if (imagePart != null)
                {
                    using var stream = imagePart.GetStream();
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    
                    base64Data = System.Convert.ToBase64String(memoryStream.ToArray());
                    contentType = imagePart.ContentType;

                    _images[relationshipId] = new DocumentImage
                    {
                        Id = relationshipId,
                        ContentType = contentType,
                        Base64Data = base64Data
                    };
                }
            }
            catch { }
        }

        if (base64Data == null || contentType == null) return string.Empty;

        return $"<img src=\"data:{contentType};base64,{base64Data}\" " +
               $"style=\"max-width:100%;width:{width}px;height:{height}px;\" " +
               $"data-image-id=\"{relationshipId}\" " +
               $"data-width-emu=\"{widthEmu}\" data-height-emu=\"{heightEmu}\" />";
    }

    /// <summary>
    /// Konwertuje Picture (stary format VML) na HTML z odczytem wymiarów
    /// </summary>
    private string ConvertPictureToHtml(Picture picture, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var imageData = picture.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
        if (imageData?.RelationshipId?.Value == null) return string.Empty;

        var relationshipId = imageData.RelationshipId.Value;

        // Spróbuj pobrać wymiary z VML shape
        var shape = picture.Descendants<DocumentFormat.OpenXml.Vml.Shape>().FirstOrDefault();
        var styleAttr = "";
        try { styleAttr = shape?.GetAttribute("style", "").Value ?? ""; } catch { }
        
        int vmlWidth = 200, vmlHeight = 150;
        var wm = Regex.Match(styleAttr, @"width:\s*([\d.]+)pt");
        var hm = Regex.Match(styleAttr, @"height:\s*([\d.]+)pt");
        if (wm.Success)
            vmlWidth = (int)(double.Parse(wm.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 96 / 72);
        if (hm.Success)
            vmlHeight = (int)(double.Parse(hm.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) * 96 / 72);

        string? base64Data = null;
        string? contentType = null;

        if (_images.TryGetValue(relationshipId, out var image))
        {
            base64Data = image.Base64Data;
            contentType = image.ContentType;
        }
        else if (sourcePart != null)
        {
            try
            {
                var imagePart = sourcePart.GetPartById(relationshipId) as ImagePart;
                if (imagePart != null)
                {
                    using var stream = imagePart.GetStream();
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    base64Data = System.Convert.ToBase64String(memoryStream.ToArray());
                    contentType = imagePart.ContentType;
                    _images[relationshipId] = new DocumentImage { Id = relationshipId, ContentType = contentType, Base64Data = base64Data };
                }
            }
            catch { }
        }

        if (base64Data == null || contentType == null) return string.Empty;

        return $"<img src=\"data:{contentType};base64,{base64Data}\" " +
               $"style=\"max-width:100%;width:{vmlWidth}px;height:{vmlHeight}px;\" " +
               $"data-image-id=\"{relationshipId}\" />";
    }

    /// <summary>
    /// Konwertuje tabelę na HTML z dokładnym odwzorowaniem obramowań, paddingu i stylów
    /// </summary>
    private string ConvertTableToHtml(Table table, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var html = new StringBuilder();
        var tableProps = table.GetFirstChild<TableProperties>();
        
        // Szerokość tabeli
        var tableWidth = "auto";
        if (tableProps?.TableWidth?.Width?.Value != null)
        {
            var w = tableProps.TableWidth;
            if (w.Type?.Value == TableWidthUnitValues.Pct)
                tableWidth = $"{int.Parse(w.Width.Value) / 50}%";
            else if (w.Type?.Value == TableWidthUnitValues.Dxa)
                tableWidth = $"{TwipsToPx(int.Parse(w.Width.Value))}px";
        }
        
        // Wyrównanie tabeli
        var tableAlign = "";
        if (tableProps?.TableJustification?.Val != null)
        {
            var tblAlignVal = tableProps.TableJustification.Val.Value;
            if (tblAlignVal == TableRowAlignmentValues.Center) tableAlign = "margin-left:auto;margin-right:auto;";
            else if (tblAlignVal == TableRowAlignmentValues.Right) tableAlign = "margin-left:auto;margin-right:0;";
        }

        // Wcięcie tabeli
        var tableIndent = "";
        if (tableProps?.TableIndentation?.Width?.Value != null)
        {
            tableIndent = $"margin-left:{TwipsToPx(tableProps.TableIndentation.Width.Value)}px;";
        }

        // Domyślne obramowania tabeli
        var defaultBorders = tableProps?.TableBorders;
        
        // Domyślny padding komórek
        var defaultPadding = "4px 8px";
        var tblCellMar = tableProps?.TableCellMarginDefault;
        if (tblCellMar != null)
        {
            var topPad = GetTwipsValue(tblCellMar.TopMargin) ?? 4;
            var bottomPad = GetTwipsValue(tblCellMar.BottomMargin) ?? 4;
            var leftPad = GetDxaValue(tblCellMar.TableCellLeftMargin) ?? 8;
            var rightPad = GetDxaValue(tblCellMar.TableCellRightMargin) ?? 8;
            defaultPadding = $"{TwipsToPx(topPad)}px {TwipsToPx(rightPad)}px {TwipsToPx(bottomPad)}px {TwipsToPx(leftPad)}px";
        }
        
        html.Append($"<table style=\"border-collapse:collapse;width:{tableWidth};margin:4px 0;{tableAlign}{tableIndent}\">");

        foreach (var row in table.Elements<TableRow>())
        {
            // Wysokość wiersza
            var rowStyle = "";
            var trHeight = row.TableRowProperties?.Elements<TableRowHeight>().FirstOrDefault();
            if (trHeight?.Val?.Value != null)
            {
                var hPx = TwipsToPx((int)trHeight.Val.Value);
                var rule = trHeight.HeightType?.Value == HeightRuleValues.Exact ? "height" : "min-height";
                rowStyle = $" style=\"{rule}:{hPx}px;\"";
            }
            
            html.Append($"<tr{rowStyle}>");
            
            foreach (var cell in row.Elements<TableCell>())
            {
                var cellProps = cell.TableCellProperties;
                var cellStyle = GetTableCellStyleDetailed(cell, defaultBorders, defaultPadding);
                
                var colspan = "";
                if (cellProps?.GridSpan?.Val?.Value is > 1)
                    colspan = $" colspan=\"{cellProps.GridSpan.Val.Value}\"";
                
                var rowspan = "";
                var vMerge = cellProps?.VerticalMerge;
                if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
                {
                    var rsc = CountRowSpan(table, row, cell);
                    if (rsc > 1) rowspan = $" rowspan=\"{rsc}\"";
                }
                else if (vMerge != null && vMerge.Val == null)
                {
                    continue;
                }
                
                html.Append($"<td{colspan}{rowspan} style=\"{cellStyle}\">");
                
                foreach (var para in cell.Elements<Paragraph>())
                    html.Append(ConvertParagraphToHtml(para, document, sourcePart));
                
                html.Append("</td>");
            }
            
            html.Append("</tr>");
        }

        html.Append("</table>");
        return html.ToString();
    }

    /// <summary>
    /// Pomocnicza: pobiera wartość twips z elementu TableWidthType
    /// </summary>
    private int? GetTwipsValue(TableWidthType? element)
    {
        if (element?.Width?.Value == null) return null;
        return int.TryParse(element.Width.Value, out var v) ? v : null;
    }

    /// <summary>
    /// Pomocnicza: pobiera wartość dxa z TableCellMarginWidth
    /// </summary>
    private int? GetDxaValue(TableWidthDxaNilType? element)
    {
        if (element?.Width?.Value == null) return null;
        return (int)element.Width.Value;
    }
    
    private int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = table.Elements<TableRow>().ToList();
        var startRowIndex = rows.IndexOf(startRow);
        var cellIndex = startRow.Elements<TableCell>().ToList().IndexOf(startCell);
        
        var rowSpan = 1;
        for (int i = startRowIndex + 1; i < rows.Count; i++)
        {
            var cells = rows[i].Elements<TableCell>().ToList();
            if (cellIndex >= cells.Count) break;
            
            var vMerge = cells[cellIndex].TableCellProperties?.VerticalMerge;
            if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                rowSpan++;
            else
                break;
        }
        
        return rowSpan;
    }

    /// <summary>
    /// Pobiera szczegółowy styl CSS komórki z pełnym odwzorowaniem obramowań
    /// </summary>
    private string GetTableCellStyleDetailed(TableCell cell, TableBorders? defaultBorders, string defaultPadding)
    {
        var css = new StringBuilder();
        var props = cell.TableCellProperties;
        
        // Obramowania: komórka > domyślne tabeli
        var cb = props?.TableCellBorders;
        css.Append($"border-top:{GetCellBorderCss(cb?.TopBorder, (BorderType?)defaultBorders?.TopBorder ?? (BorderType?)defaultBorders?.InsideHorizontalBorder)};");
        css.Append($"border-bottom:{GetCellBorderCss(cb?.BottomBorder, (BorderType?)defaultBorders?.BottomBorder ?? (BorderType?)defaultBorders?.InsideHorizontalBorder)};");
        css.Append($"border-left:{GetCellBorderCss(cb?.LeftBorder, (BorderType?)defaultBorders?.LeftBorder ?? (BorderType?)defaultBorders?.InsideVerticalBorder)};");
        css.Append($"border-right:{GetCellBorderCss(cb?.RightBorder, (BorderType?)defaultBorders?.RightBorder ?? (BorderType?)defaultBorders?.InsideVerticalBorder)};");

        // Padding
        var cm = props?.TableCellMargin;
        if (cm != null)
        {
            var top = GetTwipsValue(cm.TopMargin) ?? 0;
            var bottom = GetTwipsValue(cm.BottomMargin) ?? 0;
            var left = cm.LeftMargin != null && cm.LeftMargin.Width?.Value != null 
                ? int.Parse(cm.LeftMargin.Width.Value) : 0;
            var right = cm.RightMargin != null && cm.RightMargin.Width?.Value != null
                ? int.Parse(cm.RightMargin.Width.Value) : 0;
            css.Append($"padding:{TwipsToPx(top)}px {TwipsToPx(right)}px {TwipsToPx(bottom)}px {TwipsToPx(left)}px;");
        }
        else
        {
            css.Append($"padding:{defaultPadding};");
        }

        css.Append("vertical-align:top;");
        
        if (props != null)
        {
            // Szerokość
            var w = props.TableCellWidth;
            if (w?.Width?.Value != null)
            {
                if (w.Type?.Value == TableWidthUnitValues.Pct)
                    css.Append($"width:{int.Parse(w.Width.Value) / 50}%;");
                else
                    css.Append($"width:{TwipsToPx(int.Parse(w.Width.Value))}px;");
            }

            // Kolor tła
            if (props.Shading?.Fill?.Value != null && props.Shading.Fill.Value != "auto")
                css.Append($"background-color:#{props.Shading.Fill.Value};");

            // Wyrównanie pionowe
            if (props.TableCellVerticalAlignment?.Val != null)
                css.Append($"vertical-align:{GetTableVerticalAlignment(props.TableCellVerticalAlignment.Val.Value)};");
            
            // Kierunek tekstu
            if (props.TextDirection?.Val != null)
            {
                var tdVal = props.TextDirection.Val.Value;
                if (tdVal == TextDirectionValues.TopToBottomRightToLeft) css.Append("writing-mode:vertical-rl;");
                else if (tdVal == TextDirectionValues.BottomToTopLeftToRight) css.Append("writing-mode:vertical-lr;");
            }

            // NoWrap
            if (props.NoWrap != null)
                css.Append("white-space:nowrap;");
        }

        return css.ToString();
    }

    private string GetCellBorderCss(BorderType? cellBorder, BorderType? defaultBorder)
    {
        var border = cellBorder ?? defaultBorder;
        if (border == null) return "none";
        if (border.Val?.Value == BorderValues.None || border.Val?.Value == BorderValues.Nil) return "none";
        return GetBorderCss(border);
    }

    private string ConvertSdtBlockToHtml(SdtBlock sdtBlock, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        var content = sdtBlock.SdtContentBlock;
        if (content != null)
            foreach (var element in content.Elements())
                html.Append(ConvertElementToHtml(element, document));
        return html.ToString();
    }

    private int TwipsToPx(int twips) => (int)(twips / 1440.0 * 96);
    private int EmuToPx(long emu) => (int)(emu / 914400.0 * 96);
    private string EscapeHtml(string text) => System.Net.WebUtility.HtmlEncode(text);

    private static string GetJustificationAlignment(JustificationValues value)
    {
        if (value == JustificationValues.Left) return "left";
        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        if (value == JustificationValues.Both) return "justify";
        return "left";
    }

    private static string GetTableVerticalAlignment(TableVerticalAlignmentValues value)
    {
        if (value == TableVerticalAlignmentValues.Top) return "top";
        if (value == TableVerticalAlignmentValues.Center) return "middle";
        if (value == TableVerticalAlignmentValues.Bottom) return "bottom";
        return "top";
    }
}

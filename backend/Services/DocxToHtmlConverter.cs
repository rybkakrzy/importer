using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImporterApi.Models;

namespace ImporterApi.Services;

/// <summary>
/// Serwis do konwersji dokumentów DOCX na HTML
/// Własna implementacja parsera OpenXML
/// </summary>
public class DocxToHtmlConverter
{
    private readonly Dictionary<string, DocumentImage> _images = new();
    private readonly Dictionary<string, string> _styles = new();
    private readonly List<DocumentStyle> _documentStyles = new();
    private int _imageCounter = 0;

    /// <summary>
    /// Konwertuje plik DOCX na HTML
    /// </summary>
    public DocumentContent Convert(Stream docxStream)
    {
        _images.Clear();
        _styles.Clear();
        _documentStyles.Clear();
        _imageCounter = 0;

        using var document = WordprocessingDocument.Open(docxStream, false);
        
        // Załaduj style dokumentu
        var stylesLoaded = ExtractDocumentStyles(document);
        
        var content = new DocumentContent
        {
            Metadata = ExtractMetadata(document),
            Html = ConvertBodyToHtml(document),
            Images = _images.Values.ToList(),
            Styles = stylesLoaded.Count > 0 ? stylesLoaded : DefaultWordStyles.GetDefaultStyles()
        };

        return content;
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

        // Zliczanie słów
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
    /// Konwertuje ciało dokumentu na HTML
    /// </summary>
    private string ConvertBodyToHtml(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return "<div></div>";

        // Załaduj style dokumentu
        LoadDocumentStyles(document);
        
        // Załaduj obrazy
        LoadDocumentImages(document);

        var html = new StringBuilder();
        html.Append("<div class=\"document-content\">");

        foreach (var element in body.Elements())
        {
            html.Append(ConvertElementToHtml(element, document));
        }

        html.Append("</div>");
        return html.ToString();
    }

    /// <summary>
    /// Ładuje style z dokumentu
    /// </summary>
    private void LoadDocumentStyles(WordprocessingDocument document)
    {
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value != null)
            {
                var css = ConvertStyleToCss(style);
                _styles[style.StyleId.Value] = css;
            }
        }
    }

    /// <summary>
    /// Wyciąga style dokumentu do modelu DocumentStyle
    /// </summary>
    private List<DocumentStyle> ExtractDocumentStyles(WordprocessingDocument document)
    {
        var result = new List<DocumentStyle>();
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return result;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            // Tylko style paragrafów
            if (style.Type?.Value != StyleValues.Paragraph) continue;
            if (style.StyleId?.Value == null) continue;

            // Pobierz nazwę stylu
            var styleName = style.StyleName?.Val?.Value ?? style.StyleId.Value;
            
            // Pomiń ukryte style
            var semiHidden = style.SemiHidden;
            if (semiHidden != null && semiHidden.Val == null) // null val means "true" in OpenXML
                continue;

            var docStyle = new DocumentStyle
            {
                Id = style.StyleId.Value,
                Name = TranslateStyleName(styleName),
                Type = "paragraph",
                BasedOn = style.BasedOn?.Val?.Value,
                NextStyle = style.NextParagraphStyle?.Val?.Value
            };

            // Właściwości czcionki
            var runProps = style.StyleRunProperties;
            if (runProps != null)
            {
                // Czcionka
                var font = runProps.RunFonts;
                if (font != null)
                {
                    docStyle.FontFamily = font.Ascii?.Value ?? 
                                          font.HighAnsi?.Value ?? 
                                          font.ComplexScript?.Value;
                }

                // Rozmiar czcionki (half-points -> points)
                if (runProps.FontSize?.Val?.Value != null &&
                    double.TryParse(runProps.FontSize.Val.Value, out var fontSize))
                {
                    docStyle.FontSize = fontSize / 2;
                }

                // Kolor
                var color = runProps.Color?.Val?.Value;
                if (!string.IsNullOrEmpty(color) && color != "auto")
                {
                    docStyle.Color = "#" + color;
                }

                // Bold, Italic, Underline
                docStyle.IsBold = runProps.Bold != null && 
                                  (runProps.Bold.Val == null || runProps.Bold.Val.Value);
                docStyle.IsItalic = runProps.Italic != null && 
                                    (runProps.Italic.Val == null || runProps.Italic.Val.Value);
                docStyle.IsUnderline = runProps.Underline != null && 
                                       runProps.Underline.Val?.Value != UnderlineValues.None;
            }

            // Właściwości paragrafu
            var paraProps = style.StyleParagraphProperties;
            if (paraProps != null)
            {
                // Wyrównanie
                var justification = paraProps.Justification?.Val;
                if (justification != null && justification.HasValue)
                {
                    var jVal = justification.Value;
                    if (jVal == JustificationValues.Left)
                        docStyle.Alignment = "left";
                    else if (jVal == JustificationValues.Center)
                        docStyle.Alignment = "center";
                    else if (jVal == JustificationValues.Right)
                        docStyle.Alignment = "right";
                    else if (jVal == JustificationValues.Both)
                        docStyle.Alignment = "justify";
                }

                // Odstępy
                var spacing = paraProps.SpacingBetweenLines;
                if (spacing != null)
                {
                    // SpaceBefore/After w twips (1/20 punktu)
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

                    // Interlinia
                    if (spacing.Line?.Value != null &&
                        int.TryParse(spacing.Line.Value, out var lineVal))
                    {
                        // 240 = jednokrotna
                        docStyle.LineSpacing = lineVal / 240.0;
                    }
                }

                // Wcięcia (twips -> cm)
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

                // Poziom outline
                if (paraProps.OutlineLevel?.Val?.Value != null)
                {
                    docStyle.OutlineLevel = paraProps.OutlineLevel.Val.Value + 1;
                }
            }

            result.Add(docStyle);
        }

        // Jeśli nie ma stylów, zwróć domyślne
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
            _ => name // Zachowaj oryginalną nazwę jeśli nie ma tłumaczenia
        };
    }

    /// <summary>
    /// Konwertuje twips na centymetry
    /// </summary>
    private double TwipsToCm(int twips)
    {
        // 1 twip = 1/1440 cala = 1/567 cm
        return Math.Round(twips / 567.0, 2);
    }

    /// <summary>
    /// Konwertuje styl Word na CSS
    /// </summary>
    private string ConvertStyleToCss(Style style)
    {
        var css = new StringBuilder();
        
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
    /// Ładuje obrazy z dokumentu
    /// </summary>
    private void LoadDocumentImages(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return;

        foreach (var imagePart in mainPart.ImageParts)
        {
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            
            using var stream = imagePart.GetStream();
            using var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            
            var base64 = System.Convert.ToBase64String(memoryStream.ToArray());
            var contentType = imagePart.ContentType;

            _images[relationshipId] = new DocumentImage
            {
                Id = relationshipId,
                ContentType = contentType,
                Base64Data = base64
            };
        }
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
    /// Konwertuje paragraf na HTML
    /// </summary>
    private string ConvertParagraphToHtml(Paragraph paragraph, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        var paraProps = paragraph.ParagraphProperties;
        
        // Sprawdź czy to nagłówek
        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
        var headingLevel = GetHeadingLevel(styleId);
        
        // Sprawdź czy to lista
        var numPr = paraProps?.NumberingProperties;
        var isListItem = numPr != null;
        
        var tag = headingLevel > 0 ? $"h{headingLevel}" : "p";
        var cssStyle = GetParagraphStyle(paraProps);
        
        if (isListItem)
        {
            var listType = GetListType(numPr, document);
            html.Append($"<li style=\"{cssStyle}\">");
        }
        else
        {
            html.Append($"<{tag} style=\"{cssStyle}\">");
        }

        // Konwertuj zawartość paragrafu
        foreach (var child in paragraph.Elements())
        {
            switch (child)
            {
                case Run run:
                    html.Append(ConvertRunToHtml(run, document));
                    break;
                case Hyperlink hyperlink:
                    html.Append(ConvertHyperlinkToHtml(hyperlink, document));
                    break;
                case BookmarkStart _:
                case BookmarkEnd _:
                    // Ignoruj zakładki
                    break;
            }
        }

        // Jeśli paragraf jest pusty, dodaj &nbsp;
        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any())
        {
            html.Append("&nbsp;");
        }

        if (isListItem)
        {
            html.Append("</li>");
        }
        else
        {
            html.Append($"</{tag}>");
        }

        return html.ToString();
    }

    /// <summary>
    /// Pobiera poziom nagłówka ze stylu
    /// </summary>
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
    /// Pobiera typ listy
    /// </summary>
    private string GetListType(NumberingProperties? numPr, WordprocessingDocument document)
    {
        // Domyślnie lista nieuporządkowana
        return "ul";
    }

    /// <summary>
    /// Pobiera styl CSS dla paragrafu
    /// </summary>
    private string GetParagraphStyle(ParagraphProperties? props)
    {
        if (props == null) return string.Empty;

        var css = new StringBuilder();
        css.Append(ConvertParagraphPropertiesToCss(props));

        // Dodaj styl z definicji stylu
        var styleId = props.ParagraphStyleId?.Val?.Value;
        if (styleId != null && _styles.TryGetValue(styleId, out var styleCss))
        {
            css.Append(styleCss);
        }

        return css.ToString();
    }

    /// <summary>
    /// Konwertuje właściwości paragrafu na CSS
    /// </summary>
    private string ConvertParagraphPropertiesToCss(OpenXmlElement props)
    {
        var css = new StringBuilder();

        // Wyrównanie
        var justification = props.Descendants<Justification>().FirstOrDefault();
        if (justification?.Val != null)
        {
            var align = GetJustificationAlignment(justification.Val.Value);
            css.Append($"text-align:{align};");
        }

        // Wcięcia
        var indentation = props.Descendants<Indentation>().FirstOrDefault();
        if (indentation != null)
        {
            if (indentation.Left?.Value != null)
            {
                var leftPx = TwipsToPx(int.Parse(indentation.Left.Value));
                css.Append($"margin-left:{leftPx}px;");
            }
            if (indentation.Right?.Value != null)
            {
                var rightPx = TwipsToPx(int.Parse(indentation.Right.Value));
                css.Append($"margin-right:{rightPx}px;");
            }
            if (indentation.FirstLine?.Value != null)
            {
                var firstLinePx = TwipsToPx(int.Parse(indentation.FirstLine.Value));
                css.Append($"text-indent:{firstLinePx}px;");
            }
        }

        // Odstępy
        var spacing = props.Descendants<SpacingBetweenLines>().FirstOrDefault();
        if (spacing != null)
        {
            if (spacing.Before?.Value != null)
            {
                var beforePx = TwipsToPx(int.Parse(spacing.Before.Value));
                css.Append($"margin-top:{beforePx}px;");
            }
            if (spacing.After?.Value != null)
            {
                var afterPx = TwipsToPx(int.Parse(spacing.After.Value));
                css.Append($"margin-bottom:{afterPx}px;");
            }
            if (spacing.Line?.Value != null)
            {
                var lineHeight = int.Parse(spacing.Line.Value) / 240.0;
                css.Append($"line-height:{lineHeight:F2};");
            }
        }

        return css.ToString();
    }

    /// <summary>
    /// Konwertuje Run (fragment tekstu) na HTML
    /// </summary>
    private string ConvertRunToHtml(Run run, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        var runProps = run.RunProperties;
        var cssStyle = GetRunStyle(runProps);

        html.Append($"<span style=\"{cssStyle}\">");

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
                    html.Append("&emsp;");
                    break;
                case Drawing drawing:
                    html.Append(ConvertDrawingToHtml(drawing, document));
                    break;
                case Picture picture:
                    html.Append(ConvertPictureToHtml(picture, document));
                    break;
            }
        }

        html.Append("</span>");
        return html.ToString();
    }

    /// <summary>
    /// Pobiera styl CSS dla Run
    /// </summary>
    private string GetRunStyle(RunProperties? props)
    {
        if (props == null) return string.Empty;
        return ConvertRunPropertiesToCss(props);
    }

    /// <summary>
    /// Konwertuje właściwości Run na CSS
    /// </summary>
    private string ConvertRunPropertiesToCss(OpenXmlElement props)
    {
        var css = new StringBuilder();

        // Pogrubienie
        var bold = props.Descendants<Bold>().FirstOrDefault();
        if (bold != null && (bold.Val == null || bold.Val.Value))
        {
            css.Append("font-weight:bold;");
        }

        // Kursywa
        var italic = props.Descendants<Italic>().FirstOrDefault();
        if (italic != null && (italic.Val == null || italic.Val.Value))
        {
            css.Append("font-style:italic;");
        }

        // Podkreślenie
        var underline = props.Descendants<Underline>().FirstOrDefault();
        if (underline?.Val != null && underline.Val.Value != UnderlineValues.None)
        {
            css.Append("text-decoration:underline;");
        }

        // Przekreślenie
        var strike = props.Descendants<Strike>().FirstOrDefault();
        if (strike != null && (strike.Val == null || strike.Val.Value))
        {
            css.Append("text-decoration:line-through;");
        }

        // Rozmiar czcionki
        var fontSize = props.Descendants<FontSize>().FirstOrDefault();
        if (fontSize?.Val != null)
        {
            var size = double.Parse(fontSize.Val.Value) / 2; // Half-points to points
            css.Append($"font-size:{size}pt;");
        }

        // Rodzina czcionki
        var fontFamily = props.Descendants<RunFonts>().FirstOrDefault();
        if (fontFamily?.Ascii?.Value != null)
        {
            css.Append($"font-family:'{fontFamily.Ascii.Value}',sans-serif;");
        }

        // Kolor tekstu
        var color = props.Descendants<Color>().FirstOrDefault();
        if (color?.Val != null && color.Val.Value != "auto")
        {
            css.Append($"color:#{color.Val.Value};");
        }

        // Kolor tła (highlight)
        var highlight = props.Descendants<Highlight>().FirstOrDefault();
        if (highlight?.Val != null)
        {
            var bgColor = GetHighlightColor(highlight.Val.Value);
            css.Append($"background-color:{bgColor};");
        }

        // Shading (inne tło)
        var shading = props.Descendants<Shading>().FirstOrDefault();
        if (shading?.Fill?.Value != null && shading.Fill.Value != "auto")
        {
            css.Append($"background-color:#{shading.Fill.Value};");
        }

        // Indeks górny/dolny
        var vertAlign = props.Descendants<VerticalTextAlignment>().FirstOrDefault();
        if (vertAlign?.Val != null)
        {
            if (vertAlign.Val.Value == VerticalPositionValues.Superscript)
            {
                css.Append("vertical-align:super;font-size:smaller;");
            }
            else if (vertAlign.Val.Value == VerticalPositionValues.Subscript)
            {
                css.Append("vertical-align:sub;font-size:smaller;");
            }
        }

        return css.ToString();
    }

    /// <summary>
    /// Mapuje kolor podświetlenia Word na CSS
    /// </summary>
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

        html.Append($"<a href=\"{url ?? "#"}\" target=\"_blank\">");
        
        foreach (var run in hyperlink.Elements<Run>())
        {
            html.Append(ConvertRunToHtml(run, document));
        }

        html.Append("</a>");
        return html.ToString();
    }

    /// <summary>
    /// Konwertuje Drawing (obraz) na HTML
    /// </summary>
    private string ConvertDrawingToHtml(Drawing drawing, WordprocessingDocument document)
    {
        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return string.Empty;

        var relationshipId = blip.Embed.Value;
        
        if (_images.TryGetValue(relationshipId, out var image))
        {
            // Pobierz wymiary
            var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
            var width = extent?.Cx != null ? EmuToPx(extent.Cx.Value) : 200;
            var height = extent?.Cy != null ? EmuToPx(extent.Cy.Value) : 200;

            return $"<img src=\"data:{image.ContentType};base64,{image.Base64Data}\" " +
                   $"style=\"max-width:100%;width:{width}px;height:auto;\" " +
                   $"data-image-id=\"{image.Id}\" />";
        }

        return string.Empty;
    }

    /// <summary>
    /// Konwertuje Picture (stary format obrazu) na HTML
    /// </summary>
    private string ConvertPictureToHtml(Picture picture, WordprocessingDocument document)
    {
        var imageData = picture.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
        if (imageData?.RelationshipId?.Value == null) return string.Empty;

        var relationshipId = imageData.RelationshipId.Value;

        if (_images.TryGetValue(relationshipId, out var image))
        {
            return $"<img src=\"data:{image.ContentType};base64,{image.Base64Data}\" " +
                   $"style=\"max-width:100%;\" data-image-id=\"{image.Id}\" />";
        }

        return string.Empty;
    }

    /// <summary>
    /// Konwertuje tabelę na HTML
    /// </summary>
    private string ConvertTableToHtml(Table table, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        html.Append("<table style=\"border-collapse:collapse;width:100%;margin:10px 0;\">");

        foreach (var row in table.Elements<TableRow>())
        {
            html.Append("<tr>");
            
            foreach (var cell in row.Elements<TableCell>())
            {
                var cellStyle = GetTableCellStyle(cell);
                html.Append($"<td style=\"{cellStyle}\">");
                
                foreach (var para in cell.Elements<Paragraph>())
                {
                    html.Append(ConvertParagraphToHtml(para, document));
                }
                
                html.Append("</td>");
            }
            
            html.Append("</tr>");
        }

        html.Append("</table>");
        return html.ToString();
    }

    /// <summary>
    /// Pobiera styl CSS dla komórki tabeli
    /// </summary>
    private string GetTableCellStyle(TableCell cell)
    {
        var css = new StringBuilder();
        css.Append("border:1px solid #ccc;padding:8px;vertical-align:top;");

        var props = cell.TableCellProperties;
        if (props != null)
        {
            // Szerokość komórki
            var width = props.TableCellWidth;
            if (width?.Width?.Value != null)
            {
                var widthPx = TwipsToPx(int.Parse(width.Width.Value));
                css.Append($"width:{widthPx}px;");
            }

            // Kolor tła
            var shading = props.Shading;
            if (shading?.Fill?.Value != null && shading.Fill.Value != "auto")
            {
                css.Append($"background-color:#{shading.Fill.Value};");
            }

            // Wyrównanie pionowe
            var vAlign = props.TableCellVerticalAlignment;
            if (vAlign?.Val != null)
            {
                var align = GetTableVerticalAlignment(vAlign.Val.Value);
                css.Append($"vertical-align:{align};");
            }
        }

        return css.ToString();
    }

    /// <summary>
    /// Konwertuje blok SDT na HTML
    /// </summary>
    private string ConvertSdtBlockToHtml(SdtBlock sdtBlock, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        var content = sdtBlock.SdtContentBlock;
        
        if (content != null)
        {
            foreach (var element in content.Elements())
            {
                html.Append(ConvertElementToHtml(element, document));
            }
        }

        return html.ToString();
    }

    /// <summary>
    /// Konwertuje Twips na piksele
    /// </summary>
    private int TwipsToPx(int twips)
    {
        // 1 inch = 1440 twips, 1 inch = 96 px
        return (int)(twips / 1440.0 * 96);
    }

    /// <summary>
    /// Konwertuje EMU na piksele
    /// </summary>
    private int EmuToPx(long emu)
    {
        // 1 inch = 914400 EMU, 1 inch = 96 px
        return (int)(emu / 914400.0 * 96);
    }

    /// <summary>
    /// Escapuje tekst do HTML
    /// </summary>
    private string EscapeHtml(string text)
    {
        return System.Net.WebUtility.HtmlEncode(text);
    }

    /// <summary>
    /// Pobiera wyrównanie tekstu na podstawie JustificationValues
    /// </summary>
    private static string GetJustificationAlignment(JustificationValues value)
    {
        if (value == JustificationValues.Left) return "left";
        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        if (value == JustificationValues.Both) return "justify";
        return "left";
    }

    /// <summary>
    /// Pobiera wyrównanie pionowe tabeli
    /// </summary>
    private static string GetTableVerticalAlignment(TableVerticalAlignmentValues value)
    {
        if (value == TableVerticalAlignmentValues.Top) return "top";
        if (value == TableVerticalAlignmentValues.Center) return "middle";
        if (value == TableVerticalAlignmentValues.Bottom) return "bottom";
        return "top";
    }
}

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using HtmlAgilityPack;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Infrastructure.Services;

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

    // Domyślne ustawienia dokumentu (firmowa czcionka itp.). Wstrzykiwane przez DI;
    // dla benchmarków / testów konstruktor bezparametrowy używa wartości domyślnych.
    private readonly DocumentDefaultsOptions _defaults;

    public HtmlToDocxConverter()
    {
        _defaults = new DocumentDefaultsOptions();
    }

    public HtmlToDocxConverter(IOptions<DocumentDefaultsOptions> defaults)
    {
        _defaults = defaults?.Value ?? new DocumentDefaultsOptions();
    }


    /// <summary>
    /// Część (Part) do której mają być dodawane obrazki w bieżącym kontekście:
    /// MainDocumentPart dla body, HeaderPart / FooterPart dla nagłówka/stopki.
    /// Obrazki muszą być powiązane z częścią w której są używane (relationship),
    /// inaczej Word ich nie wyświetli.
    /// </summary>
    private OpenXmlPart? _currentImageContainer;

    /// <summary>
    /// Czy aktualnie konwertujemy header/footer (wpływa na limit szerokości obrazka).
    /// </summary>
    private bool _inHeaderFooter = false;

    /// <summary>
    /// Domyślny StyleId dla paragrafów w bieżącej sekcji (Header / Footer).
    /// Aplikowany na paragrafy, które nie mają własnego <c>data-style-id</c>.
    /// W body pozostaje <c>null</c> → Word użyje stylu Normal.
    /// </summary>
    private string? _currentSectionStyleId = null;

    /// <summary>
    /// Konwertuje HTML na plik DOCX
    /// </summary>
    public byte[] Convert(string html, DocumentMetadata? metadata = null, HeaderFooterContent? header = null, HeaderFooterContent? footer = null, PageMargins? margins = null)
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
            AddPageSettings(body, header, footer, margins);

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

            // {page}/{pages} placeholders w nagłówku też obsługujemy
            var headerHtml = header.Html
                .Replace("{page}", "<span class=\"field-page\"></span>")
                .Replace("{pages}", "<span class=\"field-numpages\"></span>");
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(headerHtml);

            // Obrazki muszą być dodane do HeaderPart, nie MainDocumentPart
            var prevContainer = _currentImageContainer;
            var prevInHF = _inHeaderFooter;
            var prevSection = _currentSectionStyleId;
            _currentImageContainer = headerPart;
            _inHeaderFooter = true;
            _currentSectionStyleId = "Header";
            try
            {
                ConvertHtmlToHeaderFooter(htmlDoc.DocumentNode, headerElement);
            }
            finally
            {
                _currentImageContainer = prevContainer;
                _inHeaderFooter = prevInHF;
                _currentSectionStyleId = prevSection;
            }

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

            var prevContainer = _currentImageContainer;
            var prevInHF = _inHeaderFooter;
            var prevSection = _currentSectionStyleId;
            _currentImageContainer = footerPart;
            _inHeaderFooter = true;
            _currentSectionStyleId = "Footer";
            try
            {
                ConvertHtmlToHeaderFooter(htmlDoc.DocumentNode, footerElement);
            }
            finally
            {
                _currentImageContainer = prevContainer;
                _inHeaderFooter = prevInHF;
                _currentSectionStyleId = prevSection;
            }

            footerPart.Footer = footerElement;
            footerPart.Footer.Save();

            var footerPartId = _mainPart.GetIdOfPart(footerPart);
            AddFooterReference(footerPartId);
        }
    }

    /// <summary>
    /// Konwertuje HTML na elementy nagłówka/stopki.
    /// Akceptuje wszystkie tagi obsługiwane przez body (h1-6, p, ul/ol, table, img, blockquote, hr).
    /// Treść która nie tworzy block-level (np. tekst bez paragrafu, samodzielny &lt;span&gt;) jest
    /// pakowana w domyślny &lt;p&gt;. Header w DOCX nie może zawierać lużnych Runów.
    /// </summary>
    private void ConvertHtmlToHeaderFooter(HtmlNode node, OpenXmlCompositeElement parent)
    {
        Paragraph? pendingTextParagraph = null;

        void FlushPending()
        {
            if (pendingTextParagraph != null)
            {
                if (!pendingTextParagraph.Elements<Run>().Any() && !pendingTextParagraph.Elements<Hyperlink>().Any())
                {
                    // pusty paragraf po flushy nie powinien być dodawany
                }
                else
                {
                    parent.Append(pendingTextParagraph);
                }
                pendingTextParagraph = null;
            }
        }

        foreach (var child in node.ChildNodes)
        {
            var name = child.Name.ToLower();
            switch (name)
            {
                case "#text":
                {
                    var text = child.InnerText;
                    if (!string.IsNullOrEmpty(text) && !string.IsNullOrWhiteSpace(text))
                    {
                        pendingTextParagraph ??= new Paragraph();
                        pendingTextParagraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
                    }
                    break;
                }
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
                case "a":
                {
                    // Inline-level w header/footer — pakujemy do bieżącego (lub nowego) paragrafu.
                    // Specjalne klasy field-page / field-numpages = pole liczby strony.
                    pendingTextParagraph ??= new Paragraph();
                    if (child.HasClass("field-page"))
                    {
                        pendingTextParagraph.Append(new Run(new SimpleField { Instruction = " PAGE " }));
                    }
                    else if (child.HasClass("field-numpages"))
                    {
                        pendingTextParagraph.Append(new Run(new SimpleField { Instruction = " NUMPAGES " }));
                    }
                    else if (name == "a")
                    {
                        pendingTextParagraph.Append(ConvertAnchorElement(child));
                    }
                    else
                    {
                        // Reużyj CreateRunsFromNode z dziedziczeniem stylu rodzica (span style=...)
                        var parentStyle = child.GetAttributeValue("style", "");
                        RunProperties? base_ = null;
                        if (!string.IsNullOrEmpty(parentStyle))
                        {
                            base_ = new RunProperties();
                            ApplyRunStyle(base_, parentStyle);
                            if (!base_.HasChildren) base_ = null;
                        }
                        foreach (var run in CreateRunsFromNode(child, base_))
                        {
                            pendingTextParagraph.Append(run);
                        }
                    }
                    break;
                }
                case "br":
                {
                    pendingTextParagraph ??= new Paragraph();
                    pendingTextParagraph.Append(new Run(new Break()));
                    break;
                }
                case "img":
                {
                    pendingTextParagraph ??= new Paragraph();
                    var imgRun = CreateImageRun(child);
                    if (imgRun != null) pendingTextParagraph.Append(imgRun);
                    break;
                }
                case "p":
                {
                    FlushPending();
                    parent.Append(ConvertParagraphElement(child));
                    break;
                }
                case "h1": case "h2": case "h3": case "h4": case "h5": case "h6":
                {
                    FlushPending();
                    var level = int.Parse(name[1].ToString());
                    parent.Append(ConvertHeadingElement(child, level));
                    break;
                }
                case "ul":
                case "ol":
                {
                    FlushPending();
                    foreach (var listElement in ConvertListElement(child, name == "ol"))
                        parent.Append(listElement);
                    break;
                }
                case "table":
                {
                    FlushPending();
                    parent.Append(ConvertTableElement(child));
                    break;
                }
                case "blockquote":
                {
                    FlushPending();
                    parent.Append(ConvertBlockquoteElement(child));
                    break;
                }
                case "hr":
                {
                    FlushPending();
                    parent.Append(CreateHorizontalRule());
                    break;
                }
                case "div":
                case "section":
                case "article":
                case "header":
                case "footer":
                {
                    // kontener — zejdź w dzieci
                    FlushPending();
                    ConvertHtmlToHeaderFooter(child, parent);
                    break;
                }
                default:
                {
                    // nieznany tag — traktujemy jak inline (zejdź w dzieci do bieżącego paragrafu)
                    pendingTextParagraph ??= new Paragraph();
                    foreach (var run in CreateRunsFromNode(child, null))
                    {
                        pendingTextParagraph.Append(run);
                    }
                    break;
                }
            }
        }

        FlushPending();

        // Header/Footer MUSZĄ zawierać co najmniej jeden block-level (np. Paragraph),
        // inaczej Word odmówi otwarcia dokumentu.
        if (!parent.Elements<Paragraph>().Any() && !parent.Elements<Table>().Any())
        {
            parent.Append(new Paragraph());
        }

        // Każdy paragraf w header/footer bez własnego stylu otrzymuje domyślny
        // styl sekcji ("Header"/"Footer"), tak jak robi to Word natywnie.
        // Dzięki temu czcionka i odstępy są zgodne z konwencją Worda
        // i tekst nie pojawia się jako "Normal".
        if (_currentSectionStyleId != null)
        {
            foreach (var p in parent.Elements<Paragraph>())
            {
                ApplyDefaultSectionStyle(p, _currentSectionStyleId);
            }
        }
    }

    /// <summary>
    /// Ustawia <c>ParagraphStyleId</c> dla paragrafu, o ile nie ma już ustawionego stylu.
    /// Używane dla paragrafów w nagłówku/stopce, by przejęły styl "Header"/"Footer".
    /// </summary>
    private static void ApplyDefaultSectionStyle(Paragraph paragraph, string styleId)
    {
        var props = paragraph.ParagraphProperties;
        if (props == null)
        {
            props = new ParagraphProperties();
            paragraph.InsertAt(props, 0);
        }
        if (props.ParagraphStyleId == null)
        {
            props.InsertAt(new ParagraphStyleId { Val = styleId }, 0);
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

        // Firmowa czcionka — z konfiguracji (sekcja DocumentDefaults w appsettings.json).
        var bodyFont = string.IsNullOrWhiteSpace(_defaults.FontFamily) ? "Calibri" : _defaults.FontFamily;
        var headingFont = string.IsNullOrWhiteSpace(_defaults.HeadingFontFamily) ? bodyFont : _defaults.HeadingFontFamily;
        // Rozmiar w DOCX jest podawany w pół-punktach (1pt = 2 jednostki).
        var halfPt = ((int)Math.Round(_defaults.FontSizePt * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Domyślne właściwości dokumentu
        var docDefaults = new DocDefaults(
            new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = bodyFont, HighAnsi = bodyFont, EastAsia = bodyFont, ComplexScript = bodyFont },
                    new FontSize { Val = halfPt },
                    new FontSizeComplexScript { Val = halfPt },
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
            new RunFonts { Ascii = bodyFont, HighAnsi = bodyFont },
            new FontSize { Val = halfPt }
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
                new RunFonts { Ascii = headingFont, HighAnsi = headingFont },
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

        // Styl Nagłówka (Header) — wbudowany styl Worda, używany dla treści
        // nagłówka strony. Bez niego Word renderuje paragraf nagłówka jako
        // Normal (bez tab-stopów do prawej/centerowania, bez odstępów),
        // co powoduje wizualne rozbieżności względem edytora.
        var headerStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Header"
        };
        headerStyle.Append(new StyleName { Val = "header" });
        headerStyle.Append(new BasedOn { Val = "Normal" });
        headerStyle.Append(new LinkedStyle { Val = "HeaderChar" });
        headerStyle.Append(new UIPriority { Val = 99 });
        headerStyle.Append(new UnhideWhenUsed());
        headerStyle.Append(new StyleParagraphProperties(
            new Tabs(
                new TabStop { Val = TabStopValues.Center, Position = 4536 },
                new TabStop { Val = TabStopValues.Right, Position = 9072 }
            ),
            new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
        ));
        styles.Append(headerStyle);

        var headerCharStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "HeaderChar",
            CustomStyle = true
        };
        headerCharStyle.Append(new StyleName { Val = "Nagłówek Znak" });
        headerCharStyle.Append(new BasedOn { Val = "DefaultParagraphFont" });
        headerCharStyle.Append(new LinkedStyle { Val = "Header" });
        headerCharStyle.Append(new UIPriority { Val = 99 });
        styles.Append(headerCharStyle);

        // Styl Stopki (Footer)
        var footerStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Footer"
        };
        footerStyle.Append(new StyleName { Val = "footer" });
        footerStyle.Append(new BasedOn { Val = "Normal" });
        footerStyle.Append(new LinkedStyle { Val = "FooterChar" });
        footerStyle.Append(new UIPriority { Val = 99 });
        footerStyle.Append(new UnhideWhenUsed());
        footerStyle.Append(new StyleParagraphProperties(
            new Tabs(
                new TabStop { Val = TabStopValues.Center, Position = 4536 },
                new TabStop { Val = TabStopValues.Right, Position = 9072 }
            ),
            new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
        ));
        styles.Append(footerStyle);

        var footerCharStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = "FooterChar",
            CustomStyle = true
        };
        footerCharStyle.Append(new StyleName { Val = "Stopka Znak" });
        footerCharStyle.Append(new BasedOn { Val = "DefaultParagraphFont" });
        footerCharStyle.Append(new LinkedStyle { Val = "Footer" });
        footerCharStyle.Append(new UIPriority { Val = 99 });
        styles.Append(footerCharStyle);

        // Domyślny styl Run (wymagany jako bazowy dla LinkedStyle)
        var defaultParagraphFont = new Style
        {
            Type = StyleValues.Character,
            StyleId = "DefaultParagraphFont",
            Default = true
        };
        defaultParagraphFont.Append(new StyleName { Val = "Default Paragraph Font" });
        defaultParagraphFont.Append(new UIPriority { Val = 1 });
        defaultParagraphFont.Append(new SemiHidden());
        defaultParagraphFont.Append(new UnhideWhenUsed());
        styles.Append(defaultParagraphFont);

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
                else if (node.HasClass("sdt-block"))
                {
                    elements.Add(BuildSdtBlockFromHtml(node));
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
    private List<OpenXmlElement> ConvertListElement(HtmlNode node, bool ordered, int level = 0, int? parentNumId = null)
    {
        var elements = new List<OpenXmlElement>();
        
        int numId;
        if (parentNumId.HasValue)
        {
            // Zagnieżdżona lista — współdziel numId z rodzicem
            numId = parentNumId.Value;
        }
        else
        {
            // Lista najwyższego poziomu — utwórz nową definicję numeracji
            EnsureNumberingPart();
            
            // Przeskanuj strukturę listy aby określić format dla każdego poziomu
            var levelFormats = new Dictionary<int, bool>();
            ScanListLevels(node, ordered, level, levelFormats);
            
            var abstractNumId = CreateAbstractNumbering(levelFormats);
            numId = CreateNumberingInstance(abstractNumId);
        }

        foreach (var child in node.ChildNodes)
        {
            if (child.Name.ToLower() != "li") continue;
            
            // Sprawdź czy li zawiera zagnieżdżone listy
            var nestedLists = child.SelectNodes("./ul|./ol");
            
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

            // Buduj base RunProperties ze stylu <li> (dziedziczenie do span/text wewnątrz)
            RunProperties? liBaseProps = null;
            if (!string.IsNullOrEmpty(liStyle))
            {
                liBaseProps = new RunProperties();
                ApplyRunStyle(liBaseProps, liStyle);
                if (!liBaseProps.HasChildren) liBaseProps = null;
            }

            // Dodaj zawartość (bez zagnieżdżonej listy)
            foreach (var liChild in child.ChildNodes)
            {
                if (liChild.Name.ToLower() == "ul" || liChild.Name.ToLower() == "ol")
                    continue; // Zagnieżdżona lista będzie obsłużona osobno
                    
                var runs = CreateRunsFromNode(liChild, liBaseProps);
                foreach (var run in runs)
                    para.Append(run);
            }

            if (!para.Elements<Run>().Any() && !para.Elements<Hyperlink>().Any())
            {
                para.Append(new Run(new Text("") { Space = SpaceProcessingModeValues.Preserve }));
            }

            elements.Add(para);

            // Obsłuż zagnieżdżone listy — współdziel numId
            if (nestedLists != null)
            {
                foreach (var nestedList in nestedLists)
                {
                    var isOrdered = nestedList.Name.ToLower() == "ol";
                    elements.AddRange(ConvertListElement(nestedList, isOrdered, level + 1, numId));
                }
            }
        }

        return elements;
    }

    /// <summary>
    /// Skanuje strukturę HTML listy aby określić format (ordered/unordered) dla każdego poziomu zagnieżdżenia
    /// </summary>
    private void ScanListLevels(HtmlNode node, bool ordered, int level, Dictionary<int, bool> levelFormats)
    {
        if (!levelFormats.ContainsKey(level))
            levelFormats[level] = ordered;

        foreach (var child in node.ChildNodes)
        {
            if (child.Name.ToLower() != "li") continue;
            var nested = child.SelectNodes("./ul|./ol");
            if (nested == null) continue;
            foreach (var nestedList in nested)
            {
                var isOrdered = nestedList.Name.ToLower() == "ol";
                ScanListLevels(nestedList, isOrdered, level + 1, levelFormats);
            }
        }
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
    private int CreateAbstractNumbering(Dictionary<int, bool> levelFormats)
    {
        var abstractNumId = _numberingId++;
        
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
        
        // Dodaj wymagany identyfikator Nsid dla poprawnej obsługi numeracji w Word
        var nsidValue = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpper();
        abstractNum.Append(new Nsid { Val = nsidValue });
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
        
        // Zdefiniuj 9 poziomów — każdy poziom ma format zgodny ze strukturą HTML
        for (int lvl = 0; lvl < 9; lvl++)
        {
            // Określ format dla danego poziomu (domyślnie jak poziom najwyższy)
            var isOrdered = levelFormats.TryGetValue(lvl, out var fmt)
                ? fmt
                : (levelFormats.TryGetValue(0, out var defaultFmt) && defaultFmt);
            
            var levelDef = new Level { LevelIndex = lvl };
            levelDef.Append(new StartNumberingValue { Val = 1 });
            
            if (isOrdered)
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
                
                // Standardowe definicje bullet Word z prawidłowymi czcionkami:
                var bulletType = lvl % 3;
                switch (bulletType)
                {
                    case 0: // Wypełnione kółko (Symbol)
                        levelDef.Append(new LevelText { Val = "\uF0B7" });
                        levelDef.Append(new NumberingSymbolRunProperties(
                            new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }
                        ));
                        break;
                    case 1: // Puste kółko (Courier New)
                        levelDef.Append(new LevelText { Val = "o" });
                        levelDef.Append(new NumberingSymbolRunProperties(
                            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New", Hint = FontTypeHintValues.Default }
                        ));
                        break;
                    case 2: // Wypełniony kwadrat (Wingdings)
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

        // Domyślne obramowania — None (żadne linie), chyba że CSS tabeli jawnie definiuje `border:`.
        // Wcześniej wymuszaliśmy solid-black jako default, co powodowało fałszywe czarne linie
        // w tabelach, które w oryginalnym DOCX miały `w:tblBorders` z val=nil/none lub w ogóle bez
        // definicji (bordery per-komórka są w pełni opisane przez ApplyCellBorders).
        var defaultBorders = new TableBorders(
            new TopBorder { Val = BorderValues.None, Size = 0 },
            new BottomBorder { Val = BorderValues.None, Size = 0 },
            new LeftBorder { Val = BorderValues.None, Size = 0 },
            new RightBorder { Val = BorderValues.None, Size = 0 },
            new InsideHorizontalBorder { Val = BorderValues.None, Size = 0 },
            new InsideVerticalBorder { Val = BorderValues.None, Size = 0 }
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
    /// Tworzy paragraf z obrazem z zachowaniem oryginalnych wymiarów EMU.
    /// </summary>
    private Paragraph CreateImageParagraph(byte[] imageBytes, string contentType, HtmlNode node)
    {
        var drawing = BuildImageDrawing(imageBytes, contentType, node);
        return drawing != null ? new Paragraph(new Run(drawing)) : new Paragraph();
    }

    /// <summary>
    /// Tworzy inline Run z obrazem (do osadzania w paragrafie header/footer/body).
    /// </summary>
    private Run? CreateImageRun(HtmlNode node)
    {
        var src = node.GetAttributeValue("src", "");
        if (string.IsNullOrEmpty(src) || !src.StartsWith("data:")) return null;
        var m = Regex.Match(src, @"data:([^;]+);base64,(.+)");
        if (!m.Success) return null;
        try
        {
            var bytes = System.Convert.FromBase64String(m.Groups[2].Value);
            var drawing = BuildImageDrawing(bytes, m.Groups[1].Value, node);
            return drawing != null ? new Run(drawing) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Buduje Drawing dla obrazka, dodaje ImagePart do aktualnej części (body/header/footer).
    /// </summary>
    private Drawing? BuildImageDrawing(byte[] imageBytes, string contentType, HtmlNode node)
    {
        var container = _currentImageContainer ?? (OpenXmlPart?)_mainPart;
        if (container == null) return null;

        var imagePartType = contentType switch
        {
            "image/png" => ImagePartType.Png,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/svg+xml" => ImagePartType.Svg,
            _ => ImagePartType.Jpeg
        };

        ImagePart imagePart = container switch
        {
            MainDocumentPart m => m.AddImagePart(imagePartType),
            HeaderPart h => h.AddImagePart(imagePartType),
            FooterPart f => f.AddImagePart(imagePartType),
            _ => _mainPart!.AddImagePart(imagePartType)
        };

        using (var stream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = container.GetIdOfPart(imagePart);

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
            // Fallback: parsuj ze stylu CSS (obsługuje liczby zmiennoprzecinkowe)
            var style = node.GetAttributeValue("style", "");
            var widthMatch = Regex.Match(style, @"width:\s*([\d.]+)px");
            var heightMatch = Regex.Match(style, @"height:\s*([\d.]+)px");

            var ci = System.Globalization.CultureInfo.InvariantCulture;
            var width = widthMatch.Success ? double.Parse(widthMatch.Groups[1].Value, ci) : 200;
            var height = heightMatch.Success ? double.Parse(heightMatch.Groups[1].Value, ci) : width * 0.75;

            // Atrybuty width/height (HTML)
            if (!widthMatch.Success)
            {
                var wAttr = node.GetAttributeValue("width", "");
                if (!string.IsNullOrEmpty(wAttr) && double.TryParse(wAttr, System.Globalization.NumberStyles.Float, ci, out var wp))
                    width = wp;
            }
            if (!heightMatch.Success)
            {
                var hAttr = node.GetAttributeValue("height", "");
                if (!string.IsNullOrEmpty(hAttr) && double.TryParse(hAttr, System.Globalization.NumberStyles.Float, ci, out var hp))
                    height = hp;
            }

            widthEmu = (long)(width * 9525);
            heightEmu = (long)(height * 9525);
        }

        // Limit szerokości:
        //  - body: ~15 cm (5 400 000 EMU)
        //  - header/footer: ~17 cm (6 120 000 EMU) – w nagłówku obrazki są zwykle szersze
        var maxWidthEmu = _inHeaderFooter ? 6_120_000L : 5_400_000L;
        if (widthEmu > maxWidthEmu)
        {
            var scale = (double)maxWidthEmu / widthEmu;
            widthEmu = maxWidthEmu;
            heightEmu = (long)(heightEmu * scale);
        }
        if (widthEmu < 9525) widthEmu = 9525;   // min 1 px
        if (heightEmu < 9525) heightEmu = 9525;

        _imageCounter++;

        return new Drawing(
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
    /// Dodaje inline content do paragrafu.
    /// Style font-* / color / font-weight / font-style ustawione na rodzicu (<p>/<h1>/<li>/...)
    /// dziedziczą się w HTML kaskadowo na dzieci. W DOCX run NIE dziedziczy automatycznie,
    /// więc budujemy bazowe `RunProperties` ze stylu rodzica i przekazujemy je
    /// do `CreateRunsFromNode` jako `inheritedProps`.
    /// </summary>
    private void AppendInlineContent(Paragraph paragraph, HtmlNode node)
    {
        RunProperties? baseRunProps = null;
        var parentStyle = node.GetAttributeValue("style", "");
        if (!string.IsNullOrEmpty(parentStyle))
        {
            baseRunProps = new RunProperties();
            ApplyRunStyle(baseRunProps, parentStyle);
            // jeżeli ApplyRunStyle nic nie dodał, traktuj jako brak
            if (!baseRunProps.HasChildren)
                baseRunProps = null;
        }

        foreach (var child in node.ChildNodes)
        {
            // Inline content control (formant) zachowany z odczytu DOCX — owijamy ponownie w SdtRun.
            if (child.NodeType == HtmlNodeType.Element
                && child.Name.Equals("span", StringComparison.OrdinalIgnoreCase)
                && child.HasClass("sdt-inline"))
            {
                var sdtRun = BuildSdtRunFromHtml(child, baseRunProps);
                if (sdtRun != null)
                    paragraph.Append(sdtRun);
                continue;
            }

            var runs = CreateRunsFromNode(child, baseRunProps);
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
        var inv = System.Globalization.CultureInfo.InvariantCulture;

        var marginTopMatch = Regex.Match(style, @"margin-top:\s*([\d.,]+)(px|pt)");
        if (marginTopMatch.Success)
        {
            var val = double.Parse(marginTopMatch.Groups[1].Value.Replace(',', '.'), inv);
            var unit = marginTopMatch.Groups[2].Value;
            if (unit == "px") val = val * 0.75; // px to pt approx
            spacing.Before = ((int)Math.Round(val * 20)).ToString(); // pt to twips
            hasSpacing = true;
        }
        
        var marginBottomMatch = Regex.Match(style, @"margin-bottom:\s*([\d.,]+)(px|pt)");
        if (marginBottomMatch.Success)
        {
            var val = double.Parse(marginBottomMatch.Groups[1].Value.Replace(',', '.'), inv);
            var unit = marginBottomMatch.Groups[2].Value;
            if (unit == "px") val = val * 0.75;
            spacing.After = ((int)Math.Round(val * 20)).ToString();
            hasSpacing = true;
        }
        
        var lineHeightMatch = Regex.Match(style, @"line-height:\s*([\d.,]+)(pt)?");
        if (lineHeightMatch.Success)
        {
            var val = double.Parse(lineHeightMatch.Groups[1].Value.Replace(',', '.'), inv);
            var unit = lineHeightMatch.Groups[2].Value;
            if (unit == "pt")
            {
                // Dokładna wartość w pt
                spacing.Line = ((int)Math.Round(val * 20)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }
            else
            {
                // Mnożnik
                spacing.Line = ((int)Math.Round(val * 240)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            hasSpacing = true;
        }
        
        if (hasSpacing)
            props.Append(spacing);

        // w:contextualSpacing (znosi odstępy między paragrafami tego samego stylu) —
        // oznaczony w CSS jako --w-contextual-spacing:1
        if (Regex.IsMatch(style, @"--w-contextual-spacing\s*:\s*1"))
        {
            props.Append(new ContextualSpacing());
        }

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
        var fontSizeMatch = Regex.Match(style, @"font-size:\s*([\d.,]+)(pt|px|em|rem)");
        if (fontSizeMatch.Success)
        {
            var size = double.Parse(fontSizeMatch.Groups[1].Value.Replace(',', '.'),
                System.Globalization.CultureInfo.InvariantCulture);
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
        var letterSpacingMatch = Regex.Match(style, @"letter-spacing:\s*([\d.,]+)(pt|px)");
        if (letterSpacingMatch.Success && !props.Elements<Spacing>().Any())
        {
            var ls = double.Parse(letterSpacingMatch.Groups[1].Value.Replace(',', '.'),
                System.Globalization.CultureInfo.InvariantCulture);
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

        extPropsPart.Properties.Application = new DocumentFormat.OpenXml.ExtendedProperties.Application("Doc2 D2Tools");
        extPropsPart.Properties.Save();
    }

    /// <summary>
    /// Dodaje ustawienia strony z dokładnymi marginesami
    /// </summary>
    private void AddPageSettings(Body body, HeaderFooterContent? header = null, HeaderFooterContent? footer = null, PageMargins? margins = null)
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
        
        // Przelicz marginesy na twipsy (1 cm = 567 twips)
        const double cmToTwips = 567.0;
        int leftTwips  = margins != null ? (int)Math.Round(margins.Left  * cmToTwips) : 1440;
        int rightTwips = margins != null ? (int)Math.Round(margins.Right * cmToTwips) : 1440;

        var headerHeightTwips = header != null ? (int)(header.Height * 1440 / 2.54) : 720;
        var footerHeightTwips = footer != null ? (int)(footer.Height * 1440 / 2.54) : 720;

        int topTwips    = margins != null ? (int)Math.Round(margins.Top    * cmToTwips) : 1440;
        int bottomTwips = margins != null ? (int)Math.Round(margins.Bottom * cmToTwips) : 1440;

        // Marginesy góra/dół muszą pomieścić nagłówek/stopkę
        var topMargin    = Math.Max(topTwips,    headerHeightTwips + 720);
        var bottomMargin = Math.Max(bottomTwips, footerHeightTwips + 720);
        
        if (!sectionProps.Elements<PageMargin>().Any())
        {
            sectionProps.Append(new PageMargin
            {
                Top    = topMargin,
                Right  = (uint)rightTwips,
                Bottom = bottomMargin,
                Left   = (uint)leftTwips,
                Header = 720,
                Footer = 720
            });
        }
    }

    private int PxToTwips(int px) => (int)(px / 96.0 * 1440);

    /// <summary>
    /// Buduje SdtProperties (Tag/Alias) na podstawie atrybutów data-sdt-* z elementu HTML.
    /// </summary>
    private static SdtProperties BuildSdtProperties(HtmlNode node)
    {
        var props = new SdtProperties();
        var tag = node.GetAttributeValue("data-sdt-tag", "");
        var alias = node.GetAttributeValue("data-sdt-alias", "");
        if (!string.IsNullOrEmpty(alias))
            props.Append(new SdtAlias { Val = System.Net.WebUtility.HtmlDecode(alias) });
        if (!string.IsNullOrEmpty(tag))
            props.Append(new Tag { Val = System.Net.WebUtility.HtmlDecode(tag) });
        return props;
    }

    /// <summary>
    /// Buduje SdtBlock z wrappera &lt;div class="sdt-block"&gt; ze znacznikami data-sdt-*.
    /// </summary>
    private SdtBlock BuildSdtBlockFromHtml(HtmlNode node)
    {
        var sdt = new SdtBlock();
        sdt.Append(BuildSdtProperties(node));

        var content = new SdtContentBlock();
        foreach (var child in node.ChildNodes)
        {
            foreach (var el in ConvertHtmlNode(child))
                content.Append(el);
        }

        // SdtContentBlock musi mieć przynajmniej jeden Paragraph, inaczej Word
        // potraktuje SDT jako uszkodzony.
        if (!content.Elements<Paragraph>().Any() && !content.Elements<Table>().Any())
            content.Append(new Paragraph());

        sdt.Append(content);
        return sdt;
    }

    /// <summary>
    /// Buduje SdtRun z &lt;span class="sdt-inline"&gt; ze znacznikami data-sdt-*.
    /// </summary>
    private SdtRun? BuildSdtRunFromHtml(HtmlNode node, RunProperties? inheritedProps)
    {
        var sdt = new SdtRun();
        sdt.Append(BuildSdtProperties(node));

        var content = new SdtContentRun();
        // Złóż run-y z dzieci span-a w tym samym kontekście co AppendInlineContent.
        foreach (var child in node.ChildNodes)
        {
            foreach (var run in CreateRunsFromNode(child, inheritedProps))
                content.Append(run);
        }

        if (!content.Elements<Run>().Any())
            content.Append(new Run(new Text("") { Space = SpaceProcessingModeValues.Preserve }));

        sdt.Append(content);
        return sdt;
    }
}

using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Conversion;
using D2ViewerEditor.Infrastructure.DocxModel;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Serwis do konwersji dokumentów DOCX na HTML
/// Własna implementacja parsera OpenXML z wysoką dokładnością odwzorowania stylów
/// </summary>
public class DocxToHtmlConverter : IDocxToHtmlConverter
{
    private readonly Dictionary<string, DocumentImage> _images = new();
    private readonly Dictionary<string, string> _styles = new();
    private readonly Dictionary<string, Style> _rawStyles = new();
    private readonly List<DocumentStyle> _documentStyles = new();
    // Cache: numPicBulletId -> data URI obrazka punktatora (z części numbering)
    private readonly Dictionary<int, string> _picBulletDataUris = new();
    private int _imageCounter = 0;
    private NumberingDefinitionsPart? _numberingPart;
    private ThemePart? _themePart;
    // Domyślne wartości z docDefaults/rPrDefault (stosowane, gdy run/style ich nie nadpisują)
    private string? _defaultFontFamily;
    private double? _defaultFontSizePt;
    // Cache dla fontów motywu: major/minor -> nazwa kroju
    private string? _themeMajorLatin;
    private string? _themeMinorLatin;
    private string? _themeMajorEastAsia;
    private string? _themeMinorEastAsia;
    private string? _themeMajorComplexScript;
    private string? _themeMinorComplexScript;
    // When rendering a paragraph that uses center/right tab stops (the classic
    // left/center/right header-footer layout), tabs become flex-grow spacers.
    private bool _flexTabs;

    // Firmowa czcionka — używana, gdy dokument nie definiuje własnej w docDefaults.
    private readonly DocumentDefaultsOptions _defaults;

    public DocxToHtmlConverter()
    {
        _defaults = new DocumentDefaultsOptions();
    }

    public DocxToHtmlConverter(IOptions<DocumentDefaultsOptions> defaults)
    {
        _defaults = defaults?.Value ?? new DocumentDefaultsOptions();
    }


    /// <summary>
    /// Konwertuje plik DOCX na HTML
    /// </summary>
    public DocumentContent Convert(Stream docxStream)
    {
        _images.Clear();
        _styles.Clear();
        _rawStyles.Clear();
        _documentStyles.Clear();
        _picBulletDataUris.Clear();
        _imageCounter = 0;
        _numberingPart = null;
        _themePart = null;
        _defaultFontFamily = null;
        _defaultFontSizePt = null;
        _themeMajorLatin = _themeMinorLatin = null;
        _themeMajorEastAsia = _themeMinorEastAsia = null;
        _themeMajorComplexScript = _themeMinorComplexScript = null;

        using var document = WordprocessingDocument.Open(docxStream, false);
        
        // Załaduj części pomocnicze
        _numberingPart = document.MainDocumentPart?.NumberingDefinitionsPart;
        _themePart = document.MainDocumentPart?.ThemePart;
        LoadThemeFonts();
        LoadNumberingPictureBullets();
        
        // Załaduj style dokumentu
        var stylesLoaded = ExtractDocumentStyles(document);
        
        var content = new DocumentContent
        {
            Metadata = ExtractMetadata(document),
            Html = ConvertBodyToHtml(document),
            Images = _images.Values.ToList(),
            Styles = stylesLoaded.Count > 0 ? stylesLoaded : DefaultWordStyles.GetDefaultStyles(),
            Header = ExtractHeader(document),
            Footer = ExtractFooter(document),
            Margins = ExtractPageMargins(document),
            PageSize = ExtractPageSize(document)
        };

        return content;
    }

    /// <summary>
    /// Page size + orientation (cm) from the first section. Null when the section
    /// declares no w:pgSz (caller falls back to its own default).
    /// </summary>
    private static Domain.Models.PageSize? ExtractPageSize(WordprocessingDocument document)
    {
        var sectionProps = document.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();
        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        if (page.PageWidthTwips is not { } width || page.PageHeightTwips is not { } height)
            return null;

        return new Domain.Models.PageSize
        {
            WidthCm = Math.Round(OoxmlUnits.TwipsToCm(width), 2),
            HeightCm = Math.Round(OoxmlUnits.TwipsToCm(height), 2),
            Orientation = page.Orientation == PageOrientation.Landscape ? "landscape" : "portrait"
        };
    }

    /// <summary>
    /// Wyciąga marginesy strony z dokumentu (w cm)
    /// </summary>
    private static PageMargins? ExtractPageMargins(WordprocessingDocument document)
    {
        var sectionProps = document.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();
        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        if (!page.HasPageMargin) return null;

        // Top/Bottom may be negative (mirror/overlap margins) — take the magnitude as Word
        // does for the printable band; Left/Right are kept as authored. Per-side default 2.5 cm.
        return new PageMargins
        {
            Top    = page.TopMarginTwips    is { } t ? Math.Round(OoxmlUnits.TwipsToCm(Math.Abs(t)), 2) : 2.5,
            Bottom = page.BottomMarginTwips is { } b ? Math.Round(OoxmlUnits.TwipsToCm(Math.Abs(b)), 2) : 2.5,
            Left   = page.LeftMarginTwips   is { } l ? Math.Round(OoxmlUnits.TwipsToCm(l),            2) : 2.5,
            Right  = page.RightMarginTwips  is { } r ? Math.Round(OoxmlUnits.TwipsToCm(r),            2) : 2.5,
        };
    }

    /// <summary>
    /// Height (cm) of the header/footer band = printable margin minus the header/footer
    /// distance, mirroring Word's geometry. Defaults: margin 0, distance 720 twips (0.5").
    /// Callers apply their own fallback when the section declares no page margin.
    /// </summary>
    private static double ComputeBandHeightCm(int? marginTwips, int? distanceTwips)
    {
        var margin = marginTwips is { } m ? Math.Abs(m) : 0;
        var distance = distanceTwips ?? 720;
        var band = margin > distance ? margin - distance : margin;
        return OoxmlUnits.TwipsToCm(band);
    }

    /// <summary>
    /// Wyciąga nagłówek z dokumentu
    /// </summary>
    private HeaderFooterContent? ExtractHeader(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return null;

        var sectionProps = mainPart.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();

        // Render the section's DEFAULT header (the one Word shows on ordinary pages),
        // resolved via sectPr/headerReference — NOT HeaderParts.FirstOrDefault(), whose
        // order is undefined and may return an empty even/first part. Fall back to the
        // first available part only when the section declares no references.
        var headerPart = ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.Default)
                         ?? mainPart.HeaderParts.FirstOrDefault();
        if (headerPart?.Header == null) return null;

        var html = ConvertHeaderPartToHtml(headerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        // First-page header is honoured only when the section opts in via titlePg.
        string? firstPageHtml = null;
        var differentFirstPage = false;
        if (HasTitlePage(sectionProps))
        {
            var firstPart = ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.First);
            if (firstPart?.Header != null)
            {
                var fph = ConvertHeaderPartToHtml(firstPart, document);
                if (!string.IsNullOrWhiteSpace(fph))
                {
                    firstPageHtml = fph;
                    differentFirstPage = true;
                }
            }
        }

        // Even-page header is honoured only when the document opts in via evenAndOddHeaders.
        string? evenHtml = null;
        var differentOddEven = false;
        if (HasEvenAndOddHeaders(mainPart))
        {
            var evenPart = ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.Even);
            if (evenPart?.Header != null)
            {
                var eh = ConvertHeaderPartToHtml(evenPart, document);
                if (!string.IsNullOrWhiteSpace(eh))
                {
                    evenHtml = eh;
                    differentOddEven = true;
                }
            }
        }

        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        double headerHeight = page.HasPageMargin
            ? ComputeBandHeightCm(page.TopMarginTwips, page.HeaderDistanceTwips)
            : 1.5;

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, headerHeight)),
            DifferentFirstPage = differentFirstPage,
            FirstPageHtml = firstPageHtml,
            DifferentOddEven = differentOddEven,
            EvenHtml = evenHtml
        };
    }

    /// <summary>
    /// Wyciąga stopkę z dokumentu
    /// </summary>
    private HeaderFooterContent? ExtractFooter(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return null;

        var sectionProps = mainPart.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();

        // See ExtractHeader: resolve the DEFAULT footer via sectPr/footerReference rather
        // than FooterParts.FirstOrDefault(), which can return an empty even/first part.
        var footerPart = ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.Default)
                         ?? mainPart.FooterParts.FirstOrDefault();
        if (footerPart?.Footer == null) return null;

        var html = ConvertFooterPartToHtml(footerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        string? firstPageHtml = null;
        var differentFirstPage = false;
        if (HasTitlePage(sectionProps))
        {
            var firstPart = ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.First);
            if (firstPart?.Footer != null)
            {
                var fph = ConvertFooterPartToHtml(firstPart, document);
                if (!string.IsNullOrWhiteSpace(fph))
                {
                    firstPageHtml = fph;
                    differentFirstPage = true;
                }
            }
        }

        string? evenHtml = null;
        var differentOddEven = false;
        if (HasEvenAndOddHeaders(mainPart))
        {
            var evenPart = ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.Even);
            if (evenPart?.Footer != null)
            {
                var eh = ConvertFooterPartToHtml(evenPart, document);
                if (!string.IsNullOrWhiteSpace(eh))
                {
                    evenHtml = eh;
                    differentOddEven = true;
                }
            }
        }

        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        double footerHeight = page.HasPageMargin
            ? ComputeBandHeightCm(page.BottomMarginTwips, page.FooterDistanceTwips)
            : 1.5;

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, footerHeight)),
            DifferentFirstPage = differentFirstPage,
            FirstPageHtml = firstPageHtml,
            DifferentOddEven = differentOddEven,
            EvenHtml = evenHtml
        };
    }

    /// <summary>
    /// Resolves the header part referenced by the section for the given type
    /// (default / first / even). Returns null when the section has no such reference.
    /// </summary>
    private static HeaderPart? ResolveHeaderPart(MainDocumentPart mainPart, SectionProperties? sectionProps, HeaderFooterValues type)
    {
        var reference = sectionProps?.Elements<HeaderReference>()
            .FirstOrDefault(r => r.Type != null && r.Type.Value == type);
        if (reference?.Id?.Value == null) return null;
        return mainPart.GetPartById(reference.Id.Value) as HeaderPart;
    }

    private static FooterPart? ResolveFooterPart(MainDocumentPart mainPart, SectionProperties? sectionProps, HeaderFooterValues type)
    {
        var reference = sectionProps?.Elements<FooterReference>()
            .FirstOrDefault(r => r.Type != null && r.Type.Value == type);
        if (reference?.Id?.Value == null) return null;
        return mainPart.GetPartById(reference.Id.Value) as FooterPart;
    }

    /// <summary>
    /// titlePg present and not explicitly disabled — the section uses a distinct first-page
    /// header/footer (an empty w:val omitted means "on", matching Word's behaviour).
    /// </summary>
    private static bool HasTitlePage(SectionProperties? sectionProps)
    {
        var titlePg = sectionProps?.GetFirstChild<TitlePage>();
        return titlePg != null && (titlePg.Val == null || titlePg.Val.Value);
    }

    /// <summary>
    /// Document-level w:evenAndOddHeaders — when present (and not disabled) the section's
    /// even header/footer reference is shown on even pages. Stored in settings.xml.
    /// </summary>
    private static bool HasEvenAndOddHeaders(MainDocumentPart mainPart)
    {
        var setting = mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>();
        return setting != null && (setting.Val == null || setting.Val.Value);
    }

    private string ConvertHeaderPartToHtml(HeaderPart part, WordprocessingDocument document)
    {
        foreach (var imagePart in part.ImageParts)
        {
            LoadImageFromPart(part, imagePart);
        }
        return ConvertHeaderFooterToHtml(part.Header, part, document);
    }

    private string ConvertFooterPartToHtml(FooterPart part, WordprocessingDocument document)
    {
        foreach (var imagePart in part.ImageParts)
        {
            LoadImageFromPart(part, imagePart);
        }
        return ConvertHeaderFooterToHtml(part.Footer, part, document);
    }

    /// <summary>
    /// Konwertuje zawartość nagłówka/stopki na HTML
    /// </summary>
    private string ConvertHeaderFooterToHtml(OpenXmlCompositeElement headerFooter, OpenXmlPart part, WordprocessingDocument document)
    {
        var inner = new StringBuilder();

        foreach (var element in headerFooter.Elements())
        {
            if (element is Paragraph para)
            {
                inner.Append(ConvertParagraphToHtml(para, document, part));
            }
            else if (element is Table table)
            {
                inner.Append(ConvertTableToHtml(table, document, part));
            }
        }

        if (inner.Length == 0) return string.Empty;

        // Owijamy treść w kontener z domyślnym krojem/rozmiarem czcionki z docDefaults —
        // analogicznie do body (.document-content). Bez tego runy nagłówka/stopki bez
        // własnego w:sz / w:rFonts dziedziczyłyby DOMYŚLNY ROZMIAR EDYTORA (zbyt duży),
        // a nie rozmiar dokumentu Word. To naprawia „za duży tekst" w nagłówku/stopce.
        var css = BuildDefaultContainerCss();
        var openTag = css.Length > 0
            ? $"<div class=\"header-footer-content\" style=\"{css}\">"
            : "<div class=\"header-footer-content\">";
        return openTag + inner + "</div>";
    }

    /// <summary>
    /// Buduje CSS kontenera z efektywnym domyślnym krojem i rozmiarem czcionki
    /// (z w:docDefaults/rPrDefault; gdy brak — z konfiguracji DocumentDefaults).
    /// Wspólne dla body oraz nagłówka/stopki, żeby runy bez własnego rPr dziedziczyły
    /// rozmiar dokumentu, a nie domyślny rozmiar edytora.
    /// </summary>
    private string BuildDefaultContainerCss()
    {
        var css = new StringBuilder();
        var effectiveFontFamily = !string.IsNullOrEmpty(_defaultFontFamily)
            ? _defaultFontFamily
            : (!string.IsNullOrWhiteSpace(_defaults.FontFamily) ? _defaults.FontFamily : null);
        if (!string.IsNullOrEmpty(effectiveFontFamily))
            css.Append(FontFamilyCss(effectiveFontFamily));
        var effectiveFontSizePt = _defaultFontSizePt ?? (_defaults.FontSizePt > 0 ? _defaults.FontSizePt : (double?)null);
        if (effectiveFontSizePt.HasValue)
            css.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "font-size:{0:0.##}pt;", effectiveFontSizePt.Value));
        return css.ToString();
    }

    /// <summary>
    /// Wyciąga metadane z dokumentu (core + extended properties)
    /// </summary>
    private DocumentMetadata ExtractMetadata(WordprocessingDocument document)
    {
        var metadata = new DocumentMetadata();

        // Core Properties (OPC)
        var coreProps = document.PackageProperties;
        if (coreProps != null)
        {
            metadata.Title = coreProps.Title;
            metadata.Author = coreProps.Creator;
            metadata.Subject = coreProps.Subject;
            metadata.Keywords = coreProps.Keywords;
            metadata.Description = coreProps.Description;
            metadata.Category = coreProps.Category;
            metadata.ContentStatus = coreProps.ContentStatus;
            metadata.LastModifiedBy = coreProps.LastModifiedBy;
            metadata.Revision = coreProps.Revision;
            metadata.Version = coreProps.Version;
            metadata.Created = coreProps.Created;
            metadata.Modified = coreProps.Modified;
        }

        // Extended Properties (app.xml)
        var extPropsPart = document.ExtendedFilePropertiesPart;
        if (extPropsPart?.Properties != null)
        {
            metadata.Company = extPropsPart.Properties.Company?.Text;
            metadata.Manager = extPropsPart.Properties.Manager?.Text;
        }

        // Word count
        var body = document.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            var text = body.InnerText;
            metadata.WordCount = text.Split(new[] { ' ', '\t', '\n', '\r' }, 
                StringSplitOptions.RemoveEmptyEntries).Length;
        }

        // Podpisy cyfrowe
        metadata.Signatures = ExtractSignatures(document);

        return metadata;
    }

    /// <summary>
    /// Wyciąga informacje o podpisach cyfrowych z Custom XML Parts
    /// </summary>
    private List<DigitalSignatureInfo> ExtractSignatures(WordprocessingDocument document)
    {
        var signatures = new List<DigitalSignatureInfo>();
        var sigNs = XNamespace.Get("http://schemas.D2ViewerEditor.app/digitalsignatures");

        if (document.MainDocumentPart == null) return signatures;

        foreach (var xmlPart in document.MainDocumentPart.CustomXmlParts)
        {
            try
            {
                using var stream = xmlPart.GetStream(FileMode.Open, FileAccess.Read);
                var doc = XDocument.Load(stream);

                if (doc.Root?.Name.Namespace == sigNs && doc.Root.Name.LocalName == "DigitalSignatures")
                {
                    foreach (var sigEl in doc.Root.Elements(sigNs + "Signature"))
                    {
                        signatures.Add(new DigitalSignatureInfo
                        {
                            SignerName = sigEl.Element(sigNs + "SignerName")?.Value ?? "",
                            SignerTitle = sigEl.Element(sigNs + "SignerTitle")?.Value,
                            SignerEmail = sigEl.Element(sigNs + "SignerEmail")?.Value,
                            Reason = sigEl.Element(sigNs + "Reason")?.Value,
                            CertificateSubject = sigEl.Element(sigNs + "CertificateSubject")?.Value ?? "",
                            CertificateIssuer = sigEl.Element(sigNs + "CertificateIssuer")?.Value ?? "",
                            CertificateSerialNumber = sigEl.Element(sigNs + "CertificateSerial")?.Value ?? "",
                            SignedAt = DateTime.TryParse(sigEl.Element(sigNs + "SignedAt")?.Value, out var signedAt) ? signedAt : DateTime.MinValue,
                            CertificateValidFrom = DateTime.TryParse(sigEl.Element(sigNs + "CertificateValidFrom")?.Value, out var from) ? from : DateTime.MinValue,
                            CertificateValidTo = DateTime.TryParse(sigEl.Element(sigNs + "CertificateValidTo")?.Value, out var to) ? to : DateTime.MaxValue,
                            IsValid = true,
                            ValidationMessage = "Podpis odczytany — pełna weryfikacja wymaga dedykowanego zapytania."
                        });
                    }
                }
            }
            catch { /* Ignoruj nieprawidłowe XML parts */ }
        }

        return signatures;
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
        var containerCss = BuildDefaultContainerCss();

        if (containerCss.Length > 0)
            html.Append($"<div class=\"document-content\" style=\"{containerCss}\">");
        else
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
    /// Sprawdza czy paragraf jest elementem listy (inline lub odziedziczone ze stylu)
    /// </summary>
    private bool IsListParagraph(Paragraph paragraph)
    {
        // Sprawdź bezpośrednie NumberingProperties na paragrafie
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0)
            return true;

        // Jeśli numeracja jest jawnie wyłączona (numId = 0), to nie jest lista
        if (numPr?.NumberingId?.Val?.Value == 0)
            return false;

        // Sprawdź numerację odziedziczoną ze stylu paragrafu
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var styleNumPr = ResolveStyleNumbering(styleId);
        return styleNumPr != null;
    }

    /// <summary>
    /// Rozwiązuje NumberingProperties z definicji stylu (z dziedziczeniem BasedOn)
    /// </summary>
    private NumberingProperties? ResolveStyleNumbering(string? styleId, HashSet<string>? visited = null)
    {
        if (styleId == null) return null;
        visited ??= new HashSet<string>();
        if (visited.Contains(styleId)) return null;
        visited.Add(styleId);

        if (!_rawStyles.TryGetValue(styleId, out var style)) return null;

        var spProps = style.StyleParagraphProperties;
        if (spProps != null)
        {
            var numPr = spProps.GetFirstChild<NumberingProperties>();
            if (numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0)
                return numPr;
        }

        var basedOn = style.BasedOn?.Val?.Value;
        if (basedOn != null)
            return ResolveStyleNumbering(basedOn, visited);

        return null;
    }

    /// <summary>
    /// Pobiera efektywne informacje o numeracji (numId, ilvl) z paragrafu lub jego stylu
    /// </summary>
    private (int numId, int level) GetEffectiveNumberingInfo(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0)
        {
            return (numPr.NumberingId.Val.Value, numPr.NumberingLevelReference?.Val?.Value ?? 0);
        }

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var styleNumPr = ResolveStyleNumbering(styleId);
        if (styleNumPr != null)
        {
            return (
                styleNumPr.NumberingId?.Val?.Value ?? 0,
                styleNumPr.NumberingLevelReference?.Val?.Value ?? 0
            );
        }

        return (0, 0);
    }

    /// <summary>
    /// Pobiera efektywne NumberingProperties z paragrafu (inline lub ze stylu)
    /// </summary>
    private NumberingProperties? GetEffectiveNumberingProps(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr?.NumberingId?.Val?.Value != null && numPr.NumberingId.Val.Value > 0)
            return numPr;

        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return ResolveStyleNumbering(styleId);
    }

    /// <summary>
    /// Pobiera wcięcia z definicji numeracji dla danego poziomu (w px)
    /// </summary>
    private (int leftPx, int hangingPx) GetNumberingLevelIndentation(NumberingProperties? numPr, int levelOverride = -1)
    {
        if (numPr == null || _numberingPart?.Numbering == null) return (0, 0);
        
        var numId = numPr.NumberingId?.Val?.Value;
        if (numId == null) return (0, 0);
        
        var level = levelOverride >= 0 ? levelOverride : (numPr.NumberingLevelReference?.Val?.Value ?? 0);
        
        var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return (0, 0);
        
        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return (0, 0);
        
        var abstractNum = _numberingPart.Numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null) return (0, 0);
        
        var levelDef = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == level);
        if (levelDef == null) return (0, 0);
        
        var prevParaProps = levelDef.GetFirstChild<PreviousParagraphProperties>();
        var indent = prevParaProps?.GetFirstChild<Indentation>();
        
        int leftTwips = 0, hangingTwips = 0;
        if (indent?.Left?.Value != null) int.TryParse(indent.Left.Value, out leftTwips);
        if (indent?.Hanging?.Value != null) int.TryParse(indent.Hanging.Value, out hangingTwips);
        
        return (TwipsToPx(leftTwips), TwipsToPx(hangingTwips));
    }

    /// <summary>
    /// Usuwa właściwości CSS związane z wcięciami z ciągu stylów (margin-left, text-indent, padding-left)
    /// Używane dla elementów <li> gdzie wcięcia obsługuje kontener <ul>/<ol>
    /// </summary>
    private static string StripIndentationCss(string css)
    {
        if (string.IsNullOrEmpty(css)) return css;
        
        css = Regex.Replace(css, @"margin-left:\s*[^;]+;?", "");
        css = Regex.Replace(css, @"text-indent:\s*[^;]+;?", "");
        css = Regex.Replace(css, @"padding-left:\s*[^;]+;?", "");
        
        return css;
    }

    /// <summary>
    /// Konwertuje kolejne elementy listy na prawidłowy HTML z zagnieżdżaniem
    /// </summary>
    private string ConvertConsecutiveListItems(List<OpenXmlElement> elements, ref int index, WordprocessingDocument document, int parentIndentPx = 0)
    {
        var html = new StringBuilder();
        
        var firstPara = (Paragraph)elements[index];
        var firstNumProps = GetEffectiveNumberingProps(firstPara);
        var (firstNumId, _) = GetEffectiveNumberingInfo(firstPara);
        var firstLevel = GetListLevel(firstPara);
        var firstInfo = GetListLevelInfo(firstNumProps, firstLevel);
        var listType = firstInfo.Tag;
        
        // Pobierz wcięcie z definicji numeracji i wylicz padding dla kontenera listy
        var (levelIndentPx, _) = GetNumberingLevelIndentation(firstNumProps, firstLevel);
        var listPadding = levelIndentPx > parentIndentPx 
            ? levelIndentPx - parentIndentPx 
            : (levelIndentPx > 0 ? levelIndentPx : 36);
        
        var listStyleCss = $"margin:0;padding-left:{listPadding}px;list-style-type:{firstInfo.ListStyleType};";
        var startAttr = (listType == "ol" && firstInfo.Start > 1) ? $" start=\"{firstInfo.Start}\"" : "";
        html.Append($"<{listType}{startAttr} style=\"{listStyleCss}\">");
        
        while (index < elements.Count)
        {
            if (elements[index] is not Paragraph p || !IsListParagraph(p))
                break;
            
            // Sprawdź czy numId się zmienił (inna lista) — ale tylko gdy też zmienia się typ formatu,
            // żeby luźne numId-y tej samej listy nie resetowały numeracji.
            var (currentNumId, _) = GetEffectiveNumberingInfo(p);
            var currentLevel = GetListLevel(p);
            
            if (currentNumId != firstNumId && currentLevel <= firstLevel)
            {
                var currentProps = GetEffectiveNumberingProps(p);
                var currentInfo = GetListLevelInfo(currentProps, currentLevel);
                if (currentInfo.Tag != firstInfo.Tag || currentInfo.ListStyleType != firstInfo.ListStyleType)
                    break;
            }
            
            if (currentLevel > firstLevel)
            {
                // Zagnieżdżona lista — przekaż aktualne wcięcie jako rodzica
                var lastLi = "</li>";
                html.Length -= lastLi.Length;
                html.Append(ConvertConsecutiveListItems(elements, ref index, document, levelIndentPx));
                html.Append("</li>");
            }
            else if (currentLevel < firstLevel)
            {
                break;
            }
            else
            {
                // Buduj CSS dla <li> BEZ wcięć — wcięcia obsługuje kontener <ul>/<ol>
                var cssStyle = GetParagraphStyle(p.ParagraphProperties);
                var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && _styles.TryGetValue(styleId, out var styleCss))
                {
                    cssStyle = styleCss + cssStyle;
                }
                cssStyle = StripIndentationCss(cssStyle);
                
                html.Append($"<li style=\"{cssStyle}\">");
                
                // Niestandardowy punktator (obrazek, checkbox z Wingdings, emoji) — wstaw własny marker
                if (firstInfo.BulletImageDataUri != null)
                {
                    html.Append($"<span class=\"list-marker\" style=\"display:inline-block;min-width:1.2em;margin-right:0.4em;\"><img src=\"{firstInfo.BulletImageDataUri}\" alt=\"\" style=\"height:1em;vertical-align:-0.125em;\"/></span>");
                }
                else if (firstInfo.BulletChar != null)
                {
                    var fontCss = !string.IsNullOrEmpty(firstInfo.BulletFont) && 
                        !firstInfo.BulletFont.ToLowerInvariant().Contains("wingdings") &&
                        !firstInfo.BulletFont.ToLowerInvariant().Contains("symbol")
                            ? $"font-family:'{firstInfo.BulletFont}';"
                            : "";
                    html.Append($"<span class=\"list-marker\" style=\"display:inline-block;min-width:1.2em;margin-right:0.4em;{fontCss}\">{System.Net.WebUtility.HtmlEncode(firstInfo.BulletChar)}</span>");
                }
                
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
                        case SdtRun sdtRun:
                            html.Append(ConvertSdtRunToHtml(sdtRun, document));
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
        var (_, level) = GetEffectiveNumberingInfo(paragraph);
        return level;
    }

    /// <summary>
    /// Ładuje style z dokumentu z rozwiązywaniem dziedziczenia
    /// </summary>
    private void LoadDocumentStyles(WordprocessingDocument document)
    {
        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return;

        // Odczytaj docDefaults/rPrDefault — domyślna czcionka i rozmiar dla całego dokumentu
        LoadDocDefaults(stylesPart);

        // Załaduj surowe style
        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value != null)
            {
                _rawStyles[style.StyleId.Value] = style;
            }
        }

        // The default paragraph style (w:default="1") overrides docDefaults for body text — e.g.
        // "Normalny" with rFonts ascii="Times New Roman" beats docDefaults asciiTheme=minorHAnsi
        // (Cambria). Body paragraphs without an explicit font use it, so the document container must
        // carry it; otherwise the editor falls back to its own default (Calibri) and the document
        // font is lost on screen.
        ApplyDefaultParagraphStyleFont();

        // Konwertuj na CSS z rozwiązywaniem dziedziczenia (BasedOn)
        foreach (var kvp in _rawStyles)
        {
            var css = ConvertStyleToCssWithInheritance(kvp.Value);
            _styles[kvp.Key] = css;
        }
    }

    private void ApplyDefaultParagraphStyleFont()
    {
        var defaultStyle = _rawStyles.Values.FirstOrDefault(s =>
            s.Type?.Value == StyleValues.Paragraph && s.Default?.Value == true);
        if (defaultStyle?.StyleRunProperties == null) return;

        var name = GetFontName(defaultStyle.StyleRunProperties.GetFirstChild<RunFonts>());
        if (!string.IsNullOrEmpty(name))
            _defaultFontFamily = name;

        var size = defaultStyle.StyleRunProperties.GetFirstChild<FontSize>();
        if (size?.Val?.Value != null &&
            double.TryParse(size.Val.Value, System.Globalization.CultureInfo.InvariantCulture, out var sz))
            _defaultFontSizePt = OoxmlUnits.HalfPointsToPoints(sz);
    }

    /// <summary>
    /// Ładuje domyślny krój i rozmiar czcionki z w:docDefaults/w:rPrDefault.
    /// Te wartości są stosowane na kontenerze dokumentu, aby każdy run dziedziczył je,
    /// gdy ani własne rPr, ani style nie definiują fontu.
    /// </summary>
    private void LoadDocDefaults(StyleDefinitionsPart stylesPart)
    {
        var docDefaults = stylesPart.Styles?.DocDefaults;
        var rPrDefault = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (rPrDefault == null) return;

        var fonts = rPrDefault.GetFirstChild<RunFonts>();
        var name = GetFontName(fonts);
        if (!string.IsNullOrEmpty(name))
            _defaultFontFamily = name;

        var size = rPrDefault.GetFirstChild<FontSize>();
        if (size?.Val?.Value != null &&
            double.TryParse(size.Val.Value, System.Globalization.CultureInfo.InvariantCulture, out var sz))
        {
            _defaultFontSizePt = OoxmlUnits.HalfPointsToPoints(sz);
        }
    }

    /// <summary>
    /// Konwertuje styl na CSS z rozwiązywaniem dziedziczenia (BasedOn)
    /// Duplikaty właściwości CSS są deduplikowane — zachowywana jest wartość z bardziej szczegółowego stylu.
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

        // Deduplikuj właściwości CSS: jeśli ta sama właściwość pojawia się wielokrotnie
        // (np. margin-bottom z Normal i margin-bottom z Heading1), zachowaj ostatnią (overridującą).
        return DeduplicateCss(css.ToString());
    }

    /// <summary>
    /// Deduplikuje właściwości CSS, zachowując ostatnie wystąpienie każdej właściwości.
    /// Zapobiega problemom z regex-parsowaniem w HtmlToDocxConverter gdy dziedziczenie
    /// powoduje duplikaty jak "margin-bottom:8pt; ... margin-bottom:0pt;".
    /// </summary>
    private static string DeduplicateCss(string css)
    {
        if (string.IsNullOrEmpty(css)) return css;

        var order = new List<string>();
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (Match m in Regex.Matches(css, @"([\w-]+)\s*:\s*([^;]+);"))
        {
            var name = m.Groups[1].Value.ToLowerInvariant();
            var val  = m.Groups[2].Value.Trim();
            if (!props.ContainsKey(name))
                order.Add(name);
            props[name] = val; // ostatnia wartość wygrywa
        }

        return string.Concat(order.Select(p => $"{p}:{props[p]};"));
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
                    docStyle.FontFamily = GetFontName(font);
                }

                if (runProps.FontSize?.Val?.Value != null &&
                    double.TryParse(runProps.FontSize.Val.Value, out var fontSize))
                {
                    docStyle.FontSize = OoxmlUnits.HalfPointsToPoints(fontSize);
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
                        docStyle.SpaceBefore = OoxmlUnits.TwipsToPoints(before);
                    }
                    if (spacing.After?.Value != null &&
                        int.TryParse(spacing.After.Value, out var after))
                    {
                        docStyle.SpaceAfter = OoxmlUnits.TwipsToPoints(after);
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

    private double TwipsToCm(int twips) => Math.Round(OoxmlUnits.TwipsToCm(twips), 2);

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

        var rawBytes = memoryStream.ToArray();
        var contentType = imagePart.ContentType;

        // EMF/WMF to wektorowe formaty Windows — żadna przeglądarka ich nie renderuje.
        // Próbujemy przekonwertować do PNG (działa na Windows; na innych OS zostaje fallback).
        if (IsMetafileContentType(contentType))
        {
            var png = TryConvertMetafileToPng(rawBytes);
            if (png != null)
            {
                rawBytes = png;
                contentType = "image/png";
            }
        }

        _images[relationshipId] = new DocumentImage
        {
            Id = relationshipId,
            ContentType = contentType,
            Base64Data = System.Convert.ToBase64String(rawBytes)
        };
    }

    private static bool IsMetafileContentType(string? contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return false;
        var ct = contentType.ToLowerInvariant();
        return ct.Contains("emf") || ct.Contains("wmf") || ct.Contains("metafile");
    }

    /// <summary>
    /// Konwertuje EMF/WMF do PNG. Najpierw próbuje LibreOffice headless (cross-platform —
    /// wystarczy zainstalowany w obrazie kontenera), potem System.Drawing na Windows.
    /// Zwraca null gdy żadna metoda nie jest dostępna albo zawiedzie.
    /// </summary>
    private static byte[]? TryConvertMetafileToPng(byte[] metafileBytes)
    {
        // 1) LibreOffice działa wszędzie (Linux/Windows/macOS) — preferowany.
        var viaSoffice = TryConvertWithLibreOffice(metafileBytes);
        if (viaSoffice != null) return viaSoffice;

        // 2) Fallback Windows-only.
        if (OperatingSystem.IsWindows())
        {
            try { return ConvertMetafileToPngWindows(metafileBytes); }
            catch { /* swallow */ }
        }

        return null;
    }

    private static readonly object _sofficeProbeLock = new();
    private static string? _sofficeBinaryCache;
    private static bool _sofficeProbed;

    private static string? ResolveSofficeBinary()
    {
        // Pamiętaj wynik wyszukiwania w obrębie procesu, żeby nie startować procesu
        // sprawdzającego dla każdego obrazka.
        if (_sofficeProbed) return _sofficeBinaryCache;
        lock (_sofficeProbeLock)
        {
            if (_sofficeProbed) return _sofficeBinaryCache;
            _sofficeProbed = true;

            // Pozwól wskazać binarkę przez zmienną środowiskową.
            var fromEnv = Environment.GetEnvironmentVariable("SOFFICE_BIN");
            if (!string.IsNullOrWhiteSpace(fromEnv) && ProbeBinary(fromEnv))
            {
                _sofficeBinaryCache = fromEnv;
                return _sofficeBinaryCache;
            }

            string[] candidates = OperatingSystem.IsWindows()
                ? new[]
                {
                    @"C:\Program Files\LibreOffice\program\soffice.exe",
                    @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                    "soffice.exe",
                    "soffice"
                }
                : new[]
                {
                    "/usr/bin/soffice",
                    "/usr/bin/libreoffice",
                    "soffice",
                    "libreoffice"
                };

            foreach (var c in candidates)
            {
                if (ProbeBinary(c))
                {
                    _sofficeBinaryCache = c;
                    return _sofficeBinaryCache;
                }
            }
            return null;
        }
    }

    private static bool ProbeBinary(string path)
    {
        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = path,
                Arguments = "--version",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var p = System.Diagnostics.Process.Start(psi);
            if (p == null) return false;
            if (!p.WaitForExit(3000)) { try { p.Kill(); } catch { } return false; }
            return p.ExitCode == 0;
        }
        catch { return false; }
    }

    private static byte[]? TryConvertWithLibreOffice(byte[] metafileBytes)
    {
        var soffice = ResolveSofficeBinary();
        if (soffice == null) return null;

        // Wybór rozszerzenia źródła ma znaczenie — LibreOffice wnioskuje filtr z rozszerzenia.
        var ext = LooksLikeWmf(metafileBytes) ? "wmf" : "emf";
        var tempDir = Path.Combine(Path.GetTempPath(), "metaconv_" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tempDir);
            var inputPath = Path.Combine(tempDir, $"in.{ext}");
            File.WriteAllBytes(inputPath, metafileBytes);

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = soffice,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                // Osobny katalog user-profile, żeby równoległe wywołania nie biły się o lock.
                Arguments = $"--headless -env:UserInstallation=file://{tempDir.Replace('\\', '/')}/profile " +
                            $"--convert-to png --outdir \"{tempDir}\" \"{inputPath}\""
            };
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc == null) return null;
            if (!proc.WaitForExit(30_000))
            {
                try { proc.Kill(true); } catch { }
                return null;
            }
            if (proc.ExitCode != 0) return null;

            var outputPath = Path.Combine(tempDir, "in.png");
            if (!File.Exists(outputPath)) return null;
            return File.ReadAllBytes(outputPath);
        }
        catch
        {
            return null;
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
        }
    }

    private static bool LooksLikeWmf(byte[] bytes)
    {
        // EMF zaczyna się od 0x01 0x00 0x00 0x00 (EMR_HEADER record type).
        // WMF placeable header: 0xD7 0xCD 0xC6 0x9A; standardowy WMF: 0x01 0x00 0x09 0x00 lub 0x02 0x00 0x09 0x00.
        if (bytes.Length < 4) return false;
        if (bytes[0] == 0xD7 && bytes[1] == 0xCD && bytes[2] == 0xC6 && bytes[3] == 0x9A) return true;
        if (bytes[0] == 0x01 && bytes[1] == 0x00 && bytes[2] == 0x09 && bytes[3] == 0x00) return true;
        if (bytes[0] == 0x02 && bytes[1] == 0x00 && bytes[2] == 0x09 && bytes[3] == 0x00) return true;
        return false;
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static byte[] ConvertMetafileToPngWindows(byte[] metafileBytes)
    {
        using var src = new MemoryStream(metafileBytes);
        using var img = System.Drawing.Image.FromStream(src);

        // EMF jest wektorowy — wybieramy sensowny DPI, by zachować ostrość.
        const float targetDpi = 144f;
        var widthPx = Math.Max(1, (int)Math.Ceiling(img.Width * targetDpi / Math.Max(1f, img.HorizontalResolution)));
        var heightPx = Math.Max(1, (int)Math.Ceiling(img.Height * targetDpi / Math.Max(1f, img.VerticalResolution)));

        // Bezpieczne ograniczenie, by nie wyprodukować ogromnego bitmapa.
        const int maxPx = 4096;
        if (widthPx > maxPx || heightPx > maxPx)
        {
            var scale = Math.Min(maxPx / (double)widthPx, maxPx / (double)heightPx);
            widthPx = Math.Max(1, (int)(widthPx * scale));
            heightPx = Math.Max(1, (int)(heightPx * scale));
        }

        using var bmp = new System.Drawing.Bitmap(widthPx, heightPx, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        bmp.SetResolution(targetDpi, targetDpi);
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.Clear(System.Drawing.Color.Transparent);
            g.DrawImage(img, 0, 0, widthPx, heightPx);
        }

        using var outMs = new MemoryStream();
        bmp.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
        return outMs.ToArray();
    }

    /// <summary>
    /// Ładuje obrazy punktatorów (w:numPicBullet) z części NumberingDefinitions.
    /// Każdy &lt;w:numPicBullet w:numPicBulletId="N"&gt; zawiera referencję do obrazka
    /// (VML lub DrawingML), który zapisujemy jako data URI gotowy do osadzenia w HTML.
    /// </summary>
    private void LoadNumberingPictureBullets()
    {
        if (_numberingPart?.Numbering == null) return;

        // Namespace dla atrybutów w XML (w: i r:).
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var picBullet in _numberingPart.Numbering.Elements<NumberingPictureBullet>())
        {
            // Czytamy id i relationship id bezpośrednio z XML — różne wersje SDK
            // OpenXml mają różne nazwy property na NumberingPictureBullet.
            int id;
            string? relId = null;
            try
            {
                var xml = XElement.Parse(picBullet.OuterXml);
                var idAttr = xml.Attribute(w + "numPicBulletId")?.Value;
                if (!int.TryParse(idAttr, out id)) continue;

                relId = xml.Descendants()
                    .Select(e => e.Attribute(r + "id")?.Value
                                 ?? e.Attribute(r + "embed")?.Value
                                 ?? e.Attribute(r + "link")?.Value)
                    .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            }
            catch
            {
                // Jeśli XML jest uszkodzony, pomijamy ten punktator.
                continue;
            }

            if (string.IsNullOrEmpty(relId)) continue;

            try
            {
                if (_numberingPart.GetPartById(relId) is ImagePart imagePart)
                {
                    using var stream = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    var bytes = ms.ToArray();
                    var contentType = imagePart.ContentType;
                    if (IsMetafileContentType(contentType))
                    {
                        var png = TryConvertMetafileToPng(bytes);
                        if (png != null) { bytes = png; contentType = "image/png"; }
                    }
                    var b64 = System.Convert.ToBase64String(bytes);
                    _picBulletDataUris[id] = $"data:{contentType};base64,{b64}";
                }
            }
            catch
            {
                // Relacja może nie istnieć albo nie wskazywać na ImagePart — ignorujemy.
            }
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
    /// Konwertuje paragraf na HTML z pełnym odwzorowaniem stylów
    /// </summary>
    private string ConvertParagraphToHtml(Paragraph paragraph, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        // A standalone page-break paragraph (only <w:br w:type="page"/>, no text/image) is emitted
        // as a TOP-LEVEL <div class="page-break"> block. Nesting it inside <p><span> broke the
        // editor's page splitter (regex split cut the <p> in half) and height-pagination ignored it,
        // so e.g. "PROTOKÓŁ…" did not start on a new page. The writer maps it back to w:br type=page.
        if (IsPageBreakOnlyParagraph(paragraph))
            return "<div class=\"page-break\"></div>";

        var html = new StringBuilder();
        var paraProps = paragraph.ParagraphProperties;

        var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
        var headingLevel = GetHeadingLevel(styleId);
        
        // Listy powinny być obsługiwane przez ConvertConsecutiveListItems
        var isListItem = IsListParagraph(paragraph);
        
        var tag = headingLevel > 0 ? $"h{headingLevel}" : "p";

        // Rozpoznaj specjalne style Worda (Title/Subtitle) — oznaczamy klasą, by CSS
        // mógł je potraktować tak samo jak nagłówki (prawdziwy bold zamiast cienkiej Calibri Light).
        var docClass = GetDocStyleClass(styleId);

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
        
        // Left/center/right one-line layout: a paragraph with center or right/end tab stops
        // becomes a flex row so its tab-separated segments spread across the width instead of
        // collapsing into fixed gaps. Tab characters are preserved (round-trip stays intact).
        var useFlexTabs = ParagraphHasAlignmentTab(paraProps);
        if (useFlexTabs)
            cssBuilder.Append("display:flex;align-items:baseline;width:100%;");

        var cssStyle = cssBuilder.ToString();
        var classAttr = docClass != null ? $" class=\"{docClass}\"" : string.Empty;
        // data-style-id pozwala eksporterowi HTML→DOCX odtworzyć oryginalny styleId (np. Title, Subtitle),
        // nawet jeśli wizualny tag to <p>.
        var dataStyleAttr = !string.IsNullOrEmpty(styleId) && docClass != null
            ? $" data-style-id=\"{System.Net.WebUtility.HtmlEncode(styleId)}\""
            : string.Empty;

        if (isListItem)
        {
            html.Append($"<li{classAttr}{dataStyleAttr} style=\"{cssStyle}\">");
        }
        else
        {
            html.Append($"<{tag}{classAttr}{dataStyleAttr} style=\"{cssStyle}\">");
        }

        var prevFlexTabs = _flexTabs;
        _flexTabs = useFlexTabs;

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
                    case SdtRun sdtRun:
                        html.Append(ConvertSdtRunToHtml(sdtRun, document, sourcePart));
                        break;
                }
            }
        }

        _flexTabs = prevFlexTabs;

        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any() && !paragraph.Elements<SimpleField>().Any())
        {
            html.Append("&nbsp;");
        }

        html.Append(isListItem ? "</li>" : $"</{tag}>");
        return html.ToString();
    }

    /// <summary>
    /// True when the paragraph's only meaningful content is a manual page break
    /// (<c>w:br type=page</c>) — i.e. a dedicated page-break paragraph, not text that merely
    /// happens to break. Such paragraphs render as a standalone page-break block.
    /// </summary>
    private static bool IsPageBreakOnlyParagraph(Paragraph paragraph)
    {
        var hasPageBreak = paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);
        if (!hasPageBreak) return false;

        var hasText = paragraph.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text));
        var hasGraphics = paragraph.Descendants<Drawing>().Any() || paragraph.Descendants<Picture>().Any();
        return !hasText && !hasGraphics;
    }

    /// <summary>
    /// True when the paragraph declares a center or right/end tab stop — the signature of a
    /// left/center/right one-line layout (typical Word header/footer).
    /// </summary>
    private static bool ParagraphHasAlignmentTab(OpenXmlElement? paraProps)
    {
        var tabs = paraProps?.GetFirstChild<Tabs>();
        if (tabs == null) return false;

        return tabs.Elements<TabStop>().Any(t =>
            t.Val?.Value == TabStopValues.Center ||
            t.Val?.Value == TabStopValues.Right ||
            t.Val?.Value == TabStopValues.End);
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
                        html.Append(FieldSpan("field-page", "{page}", run));
                    }
                    else if (instr.Contains("NUMPAGES") || instr.Contains("SECTIONPAGES"))
                    {
                        html.Append(FieldSpan("field-numpages", "{pages}", run));
                    }
                    else if (instr.Contains("DATE") || instr.Contains("TIME"))
                    {
                        html.Append(FieldSpan("field-date", DateTime.Now.ToString("dd.MM.yyyy"), run));
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
                            html.Append(FieldSpan("field-page", "{page}", run));
                        }
                        else if (instrEnd.Contains("NUMPAGES") || instrEnd.Contains("SECTIONPAGES"))
                        {
                            html.Append(FieldSpan("field-numpages", "{pages}", run));
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
    /// Konwertuje obramowanie OpenXML na CSS border string.
    /// Zwraca "none" gdy brak sensownej definicji borderu (brak Val lub Val=None/Nil).
    /// </summary>
    private string GetBorderCss(BorderType border)
    {
        var borderVal = border.Val?.Value;

        // Brak Val lub explicit None/Nil -> żadnej linii (zapobiega fałszywym czarnym liniom na eksporcie).
        if (borderVal == null || borderVal == BorderValues.None || borderVal == BorderValues.Nil)
            return "none";

        var size = border.Size?.Value ?? 4;
        var sizePx = Math.Max(1, size / 8.0);
        var color = border.Color?.Value ?? "000000";
        if (color == "auto") color = "000000";

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
        
        return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.#}px {1} #{2}", sizePx, style, color);
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
    /// Mapuje nazwane style Worda (Title/Subtitle) na klasy CSS dla warstwy prezentacyjnej.
    /// Pozwala frontendowi zastosować dla nich prawdziwy bold zamiast cienkiej Calibri Light.
    /// </summary>
    private static string? GetDocStyleClass(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var s = styleId.Replace(" ", string.Empty);
        if (string.Equals(s, "Title", StringComparison.OrdinalIgnoreCase)) return "doc-title";
        if (string.Equals(s, "Subtitle", StringComparison.OrdinalIgnoreCase)) return "doc-subtitle";
        return null;
    }

    /// <summary>
    /// Pobiera typ listy (ol/ul) na podstawie definicji numeracji w dokumencie
    /// </summary>
    private string GetListType(NumberingProperties? numPr, WordprocessingDocument document)
    {
        var info = GetListLevelInfo(numPr);
        return info.Tag;
    }

    /// <summary>
    /// Pełna informacja o poziomie listy: tag (ol/ul), CSS list-style-type, znak punktatora
    /// (gdy niestandardowy), czcionka punktatora oraz początkowa wartość numeracji.
    /// </summary>
    private readonly struct ListLevelInfo
    {
        public string Tag { get; init; }
        public string ListStyleType { get; init; }
        public string? BulletChar { get; init; }
        public string? BulletFont { get; init; }
        public string? BulletImageDataUri { get; init; }
        public int Start { get; init; }
    }

    private ListLevelInfo GetListLevelInfo(NumberingProperties? numPr, int levelOverride = -1)
    {
        var fallback = new ListLevelInfo { Tag = "ul", ListStyleType = "disc", Start = 1 };
        if (numPr == null || _numberingPart?.Numbering == null) return fallback;

        var numId = numPr.NumberingId?.Val?.Value;
        if (numId == null) return fallback;
        var level = levelOverride >= 0 ? levelOverride : (numPr.NumberingLevelReference?.Val?.Value ?? 0);

        var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return fallback;

        // LevelOverride wewnątrz NumberingInstance ma pierwszeństwo nad AbstractNum
        var levelOverrideElem = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(lo => lo.LevelIndex?.Value == level);
        Level? levelDef = levelOverrideElem?.GetFirstChild<Level>();

        int startOverride = levelOverrideElem?.StartOverrideNumberingValue?.Val?.Value ?? -1;

        if (levelDef == null)
        {
            var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
            if (abstractNumId == null) return fallback;
            var abstractNum = _numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
            if (abstractNum == null) return fallback;
            levelDef = abstractNum.Elements<Level>()
                .FirstOrDefault(l => l.LevelIndex?.Value == level);
        }
        if (levelDef == null) return fallback;

        var numFmt = levelDef.NumberingFormat?.Val?.Value;
        var levelText = levelDef.LevelText?.Val?.Value ?? string.Empty;
        // Word zapisuje krój punktatora w jednym z atrybutów RunFonts (Ascii/HighAnsi/Cs).
        // Bierzemy pierwszy niepusty — typowo dla Wingdings będzie to HighAnsi.
        var bulletFontRun = levelDef.NumberingSymbolRunProperties?.GetFirstChild<RunFonts>();
        var bulletFont = bulletFontRun?.Ascii?.Value
                         ?? bulletFontRun?.HighAnsi?.Value
                         ?? bulletFontRun?.ComplexScript?.Value
                         ?? bulletFontRun?.EastAsia?.Value;
        var start = startOverride > 0
            ? startOverride
            : (levelDef.StartNumberingValue?.Val?.Value ?? 1);

        // Picture bullet (w:lvlPicBulletId) — Word pozwala wstawić obrazek jako punktator.
        // Jeśli istnieje, użyjemy obrazka zamiast znaku.
        string? bulletImageDataUri = null;
        var picBulletId = levelDef.LevelPictureBulletId?.Val?.Value;
        if (picBulletId.HasValue && _picBulletDataUris.TryGetValue(picBulletId.Value, out var picUri))
        {
            bulletImageDataUri = picUri;
        }

        string tag = "ul";
        string listStyle = "disc";
        string? bulletChar = null;

        if (numFmt == NumberFormatValues.Decimal) { tag = "ol"; listStyle = "decimal"; }
        else if (numFmt == NumberFormatValues.DecimalZero) { tag = "ol"; listStyle = "decimal-leading-zero"; }
        else if (numFmt == NumberFormatValues.UpperLetter) { tag = "ol"; listStyle = "upper-alpha"; }
        else if (numFmt == NumberFormatValues.LowerLetter) { tag = "ol"; listStyle = "lower-alpha"; }
        else if (numFmt == NumberFormatValues.UpperRoman) { tag = "ol"; listStyle = "upper-roman"; }
        else if (numFmt == NumberFormatValues.LowerRoman) { tag = "ol"; listStyle = "lower-roman"; }
        else if (numFmt == NumberFormatValues.Bullet)
        {
            tag = "ul";
            // Bezpieczne pobranie pierwszego code-pointa (obsługa par surogatów dla emoji).
            int codePoint = 0;
            if (!string.IsNullOrEmpty(levelText))
            {
                codePoint = char.IsHighSurrogate(levelText[0]) && levelText.Length > 1
                    ? char.ConvertToUtf32(levelText[0], levelText[1])
                    : levelText[0];
            }

            // Word zapisuje znaki z fontów symbolicznych (Wingdings/Symbol) często w obszarze
            // Private Use Area U+F000..U+F0FF. Dla rozpoznawania punktatora interesuje nas
            // wtedy tylko młodszy bajt.
            var fontLower = (bulletFont ?? string.Empty).ToLowerInvariant();
            bool isSymbolicFont = fontLower.Contains("wingdings") || fontLower.Contains("symbol");
            int lookup = (isSymbolicFont || (codePoint >= 0xF000 && codePoint <= 0xF0FF))
                ? (codePoint & 0xFF)
                : codePoint;

            if (bulletImageDataUri != null)
            {
                // Picture bullet ma pierwszeństwo — wyłącz natywny punktator HTML.
                listStyle = "none";
            }
            else if (isSymbolicFont)
            {
                // Dla fontów symbolicznych ZAWSZE renderujemy własny marker —
                // znak źródłowy (np. Wingdings 0x6C) w przeglądarce bez Wingdings dałby tofu.
                listStyle = "none";
                bulletChar = MapBulletChar(lookup, bulletFont);
            }
            else
            {
                switch (lookup)
                {
                    case 0x2022: case 'l': listStyle = "disc"; break;
                    case 'o': case 0x25E6: listStyle = "circle"; break;
                    case 0x00A7: case 0x25AA: case 0x25FE: listStyle = "square"; break;
                    default:
                        if (codePoint != 0)
                        {
                            // Niestandardowy punktator (np. emoji) — renderujemy własny marker.
                            listStyle = "none";
                            bulletChar = MapBulletChar(codePoint, bulletFont);
                        }
                        else
                        {
                            listStyle = "disc";
                        }
                        break;
                }
            }
        }
        else
        {
            // Inne (np. NumberInDash) — fallback do decimal jeśli text wygląda na numer.
            if (!string.IsNullOrEmpty(levelText) && levelText.Contains("%"))
            { tag = "ol"; listStyle = "decimal"; }
        }

        return new ListLevelInfo
        {
            Tag = tag,
            ListStyleType = listStyle,
            BulletChar = bulletChar,
            BulletFont = bulletFont,
            BulletImageDataUri = bulletImageDataUri,
            Start = start
        };
    }

    /// <summary>
    /// Mapuje znak punktatora z Wingdings/Symbol na odpowiedni Unicode, lub zwraca znak
    /// niezmieniony gdy nie wymaga konwersji. Obsługuje pełne code-pointy (włącznie z emoji
    /// poza BMP, np. 🚀 = U+1F680).
    /// </summary>
    private static string MapBulletChar(int codePoint, string? font)
    {
        var f = (font ?? string.Empty).ToLowerInvariant();

        // Word zapisuje znaki z fontów symbolicznych (Wingdings, Symbol) często w obszarze
        // Private Use Area (PUA) U+F000..U+F0FF — to ten sam kod 0xXX, ale „przesunięty”.
        // Dla mapowania interesuje nas tylko młodszy bajt.
        int low = codePoint & 0xFF;

        if (f.Contains("wingdings"))
        {
            return low switch
            {
                0xFE => "\u2611", // ☑ zaznaczony checkbox
                0xA8 => "\u2610", // ☐ pusty checkbox
                0xFC => "\u2714", // ✔ haczyk
                0xA7 => "\u25A0", // ■ wypełniony kwadrat
                0x6C => "\u2022", // • bullet
                0xD8 => "\u2756", // ornament
                // Nieznany kod z Wingdings — bezpieczna kropka, lepsza niż „tofu”.
                _ => "\u2022"
            };
        }
        if (f.Contains("symbol"))
        {
            return low switch
            {
                0xB7 => "\u2022",
                0xA8 => "\u25E6",
                _ => "\u2022"
            };
        }
        // Zwykły Unicode (w tym emoji poza BMP) — buduj poprawny string z code-pointa.
        try { return char.ConvertFromUtf32(codePoint); }
        catch { return "\u2022"; }
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
            var inv = System.Globalization.CultureInfo.InvariantCulture;

            // w:beforeAutoSpacing="true" oznacza, że Word ignoruje wartość Before i wylicza auto.
            // W HTML nie mamy sensownego odpowiednika — pomijamy emisję, by nie wpisać niepoprawnej wartości.
            var beforeAuto = spacing.BeforeAutoSpacing?.Value == true;
            var afterAuto = spacing.AfterAutoSpacing?.Value == true;

            if (!beforeAuto)
            {
                if (spacing.Before?.Value != null && int.TryParse(spacing.Before.Value, out var beforeVal))
                    css.Append(string.Format(inv, "margin-top:{0:0.##}pt;", OoxmlUnits.TwipsToPoints(beforeVal)));
                else if (spacing.BeforeLines?.Value != null)
                {
                    // BeforeLines jest w 1/100 linii — przybliżamy do wielokrotności domyślnego rozmiaru.
                    var pt = spacing.BeforeLines.Value / 100.0 * (_defaultFontSizePt ?? 11);
                    css.Append(string.Format(inv, "margin-top:{0:0.##}pt;", pt));
                }
            }

            if (!afterAuto)
            {
                if (spacing.After?.Value != null && int.TryParse(spacing.After.Value, out var afterVal))
                    css.Append(string.Format(inv, "margin-bottom:{0:0.##}pt;", OoxmlUnits.TwipsToPoints(afterVal)));
                else if (spacing.AfterLines?.Value != null)
                {
                    var pt = spacing.AfterLines.Value / 100.0 * (_defaultFontSizePt ?? 11);
                    css.Append(string.Format(inv, "margin-bottom:{0:0.##}pt;", pt));
                }
            }

            if (spacing.Line?.Value != null && int.TryParse(spacing.Line.Value, out var lineVal))
            {
                var lineRule = spacing.LineRule?.Value;
                if (lineRule == LineSpacingRuleValues.Exact || lineRule == LineSpacingRuleValues.AtLeast)
                {
                    css.Append(string.Format(inv, "line-height:{0:0.##}pt;", OoxmlUnits.TwipsToPoints(lineVal)));
                }
                else
                {
                    // Auto (domyślne gdy brak w:lineRule) — wartość jest w 240-tych częściach linii.
                    css.Append(string.Format(inv, "line-height:{0:0.###};", lineVal / 240.0));
                }
            }
        }

        // w:contextualSpacing — gdy true, Word znosi odstępy między sąsiednimi paragrafami tego samego stylu.
        // Eksportujemy jako data-attribute, aby HtmlToDocxConverter mógł to przywrócić.
        var contextualSpacing = props.Descendants<ContextualSpacing>().FirstOrDefault();
        if (contextualSpacing != null && (contextualSpacing.Val == null || contextualSpacing.Val.Value))
        {
            // Zaznacz w CSS custom property — HtmlToDocx to rozpozna.
            css.Append("--w-contextual-spacing:1;");
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

        // Inside a flex tab-stop paragraph, a tab-only run becomes a growing spacer that is a
        // direct flex child (so it actually distributes the L/C/R segments). The tab character
        // is kept for round-trip. Mixed runs fall through to normal rendering.
        if (_flexTabs && run.Elements<TabChar>().Any()
            && !run.Elements<Text>().Any() && !run.Elements<Drawing>().Any()
            && !run.Elements<Picture>().Any() && !run.Elements<Break>().Any())
        {
            return "<span style=\"flex:1 1 0;\">\t</span>";
        }

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

        // Resolve a named character style (w:rStyle) and lay its inherited CSS *underneath*
        // the run's direct formatting (direct wins on conflicts, which come later in the
        // declaration). Without this, runs formatted only via a character style — Hyperlink,
        // Strong, Emphasis, custom — rendered with no formatting at all.
        var rStyleId = runProps?.RunStyle?.Val?.Value;
        var rStyleCss = rStyleId != null && _styles.TryGetValue(rStyleId, out var rsCss)
            ? rsCss
            : string.Empty;

        html.Append($"<span style=\"{rStyleCss}{cleanCss}\">");
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
        var inv = System.Globalization.CultureInfo.InvariantCulture;

        // Rozmiar czcionki
        var fontSize = props.Descendants<FontSize>().FirstOrDefault();
        if (fontSize?.Val != null &&
            double.TryParse(fontSize.Val.Value, System.Globalization.NumberStyles.Float, inv, out var sz))
        {
            css.Append(string.Format(inv, "font-size:{0:0.##}pt;", OoxmlUnits.HalfPointsToPoints(sz)));
        }

        // Rodzina czcionki (z obsługą theme fonts: asciiTheme/hAnsiTheme/...)
        var fontFamily = props.Descendants<RunFonts>().FirstOrDefault();
        var fontName = GetFontName(fontFamily);
        if (fontName != null)
            css.Append(FontFamilyCss(fontName));

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
            css.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "letter-spacing:{0:0.#}pt;", OoxmlUnits.TwipsToPoints(spacing.Val.Value)));

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
        if (fontSize?.Val != null &&
            double.TryParse(fontSize.Val.Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var size))
        {
            css.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "font-size:{0:0.##}pt;", OoxmlUnits.HalfPointsToPoints(size)));
        }

        var fontFamily = props.Descendants<RunFonts>().FirstOrDefault();
        var fontName = GetFontName(fontFamily);
        if (fontName != null)
            css.Append(FontFamilyCss(fontName));

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
            css.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "letter-spacing:{0:0.#}pt;", OoxmlUnits.TwipsToPoints(spacing.Val.Value)));

        var capsProp = props.Descendants<Caps>().FirstOrDefault();
        if (capsProp != null && (capsProp.Val == null || capsProp.Val.Value))
            css.Append("text-transform:uppercase;");
        var smallCapsProp = props.Descendants<SmallCaps>().FirstOrDefault();
        if (smallCapsProp != null && (smallCapsProp.Val == null || smallCapsProp.Val.Value))
            css.Append("font-variant:small-caps;");

        return css.ToString();
    }

    /// <summary>
    /// Ładuje nazwy czcionek z motywu (major/minor -> Latin/EastAsia/ComplexScript).
    /// </summary>
    private void LoadThemeFonts()
    {
        var fontScheme = _themePart?.Theme?.ThemeElements?.FontScheme;
        if (fontScheme == null) return;

        var major = fontScheme.MajorFont;
        if (major != null)
        {
            _themeMajorLatin = major.LatinFont?.Typeface?.Value;
            _themeMajorEastAsia = major.EastAsianFont?.Typeface?.Value;
            _themeMajorComplexScript = major.ComplexScriptFont?.Typeface?.Value;
        }

        var minor = fontScheme.MinorFont;
        if (minor != null)
        {
            _themeMinorLatin = minor.LatinFont?.Typeface?.Value;
            _themeMinorEastAsia = minor.EastAsianFont?.Typeface?.Value;
            _themeMinorComplexScript = minor.ComplexScriptFont?.Typeface?.Value;
        }
    }

    /// <summary>
    /// Rozwiązuje theme-font na konkretną nazwę kroju odczytaną z theme1.xml.
    /// </summary>
    private string? ResolveThemeFont(ThemeFontValues theme)
    {
        if (theme == ThemeFontValues.MajorAscii || theme == ThemeFontValues.MajorHighAnsi) return _themeMajorLatin;
        if (theme == ThemeFontValues.MinorAscii || theme == ThemeFontValues.MinorHighAnsi) return _themeMinorLatin;
        if (theme == ThemeFontValues.MajorEastAsia) return _themeMajorEastAsia;
        if (theme == ThemeFontValues.MinorEastAsia) return _themeMinorEastAsia;
        if (theme == ThemeFontValues.MajorBidi) return _themeMajorComplexScript;
        if (theme == ThemeFontValues.MinorBidi) return _themeMinorComplexScript;
        return null;
    }

    /// <summary>
    /// Wyciąga nazwę czcionki z RunFonts uwzględniając zarówno jawne atrybuty,
    /// jak i referencje do motywu (AsciiTheme, HighAnsiTheme, itd.).
    /// </summary>
    private static string FontFamilyCss(string fontName) =>
        $"font-family:'{fontName}',{GenericFontFallback(fontName)};";

    /// <summary>
    /// Picks a generic CSS fallback matching the font's family class. Critical for environments
    /// where the exact font is not installed: a missing serif (e.g. Times New Roman, Cambria) must
    /// fall back to <c>serif</c>, not <c>sans-serif</c> — otherwise body text renders like Calibri.
    /// The original font name is always kept first; this is only the after-comma fallback.
    /// </summary>
    private static string GenericFontFallback(string fontName)
    {
        var f = fontName.ToLowerInvariant();
        if (f.Contains("times") || f.Contains("cambria") || f.Contains("georgia") || f.Contains("garamond")
            || f.Contains("minion") || f.Contains("book antiqua") || f.Contains("palatino")
            || (f.Contains("serif") && !f.Contains("sans")))
            return "serif";
        if (f.Contains("courier") || f.Contains("consolas") || f.Contains("mono"))
            return "monospace";
        return "sans-serif";
    }

    private string? GetFontName(RunFonts? fonts)
    {
        if (fonts == null) return null;

        // Dla tekstu łacińskiego (przeglądarka renderuje wszystko jednym fontem)
        // Word używa wyłącznie ascii/hAnsi (i ich theme-owych odpowiedników).
        // w:cs (Complex Script — arabski, hebrajski) i w:eastAsia (CJK) są stosowane
        // tylko dla odpowiednich zakresów znaków — NIE wolno spadać na nie jako fallback,
        // bo wówczas zwykłe runy z `<w:rFonts w:cs="Arial" w:asciiTheme="minorHAnsi"/>`
        // dostają błędnie "Arial" zamiast firmowego "Calibri".

        // 1) Jawny ascii
        if (!string.IsNullOrEmpty(fonts.Ascii?.Value)) return fonts.Ascii!.Value;

        // 2) Theme dla ascii
        if (fonts.AsciiTheme?.Value != null)
        {
            var resolved = ResolveThemeFont(fonts.AsciiTheme.Value);
            if (!string.IsNullOrEmpty(resolved)) return resolved;
        }

        // 3) Jawny hAnsi (high-ANSI: znaki Latin Extended)
        if (!string.IsNullOrEmpty(fonts.HighAnsi?.Value)) return fonts.HighAnsi!.Value;

        // 4) Theme dla hAnsi
        if (fonts.HighAnsiTheme?.Value != null)
        {
            var resolved = ResolveThemeFont(fonts.HighAnsiTheme.Value);
            if (!string.IsNullOrEmpty(resolved)) return resolved;
        }

        return null;
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
        var fieldRun = simpleField.Descendants<Run>().FirstOrDefault();

        if (instruction.Contains("PAGE") && !instruction.Contains("NUMPAGES") && !instruction.Contains("SECTIONPAGES"))
            return FieldSpan("field-page", "{page}", fieldRun);
        if (instruction.Contains("NUMPAGES") || instruction.Contains("SECTIONPAGES"))
            return FieldSpan("field-numpages", "{pages}", fieldRun);
        if (instruction.Contains("DATE") || instruction.Contains("TIME"))
            return FieldSpan("field-date", DateTime.Now.ToString("dd.MM.yyyy"), fieldRun);

        var text = string.Join("", simpleField.Descendants<Text>().Select(t => t.Text));
        return !string.IsNullOrEmpty(text) ? EscapeHtml(text) : "";
    }

    /// <summary>
    /// Emits a field placeholder span carrying the field run's clean CSS (font-size/family/colour)
    /// so e.g. a PAGE number in the footer matches the surrounding footer text instead of falling
    /// back to the container/editor default size. Empty style → inherits via CSS.
    /// </summary>
    private string FieldSpan(string cssClass, string placeholder, Run? run)
    {
        var style = run?.RunProperties != null ? GetRunStyleClean(run.RunProperties) : string.Empty;
        return $"<span class=\"{cssClass}\" style=\"{style}\">{placeholder}</span>";
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
        var widthEmu = extent?.Cx?.Value ?? OoxmlUnits.PixelsToEmu(width);
        var heightEmu = extent?.Cy?.Value ?? OoxmlUnits.PixelsToEmu(height);

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

        // Word-like floating positioning: wp:anchor → emit data-pos-mode + offsets so
        // the editor restores "front"/"behind" mode and the wrapper is anchored.
        var anchor = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
        var posAttrs = string.Empty;
        if (anchor != null)
        {
            var behind = anchor.BehindDoc?.Value == true;
            var posH = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition>();
            var posV = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition>();
            long xEmu = 0, yEmu = 0;
            if (posH?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset>()?.Text is string xText
                && long.TryParse(xText, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var xParsed))
                xEmu = xParsed;
            if (posV?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset>()?.Text is string yText
                && long.TryParse(yText, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var yParsed))
                yEmu = yParsed;
            posAttrs = $" data-pos-mode=\"{(behind ? "behind" : "front")}\""
                + $" data-x-emu=\"{xEmu}\" data-y-emu=\"{yEmu}\"";
        }

        // Optional border (a:ln) — width in EMU → px, color from solidFill srgbClr, dash style.
        var borderAttrs = string.Empty;
        var outline = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
        if (outline != null && outline.Width != null && outline.Width.Value > 0)
        {
            var borderWidthPx = Math.Max(1, (int)Math.Round(OoxmlUnits.EmuToPixels(outline.Width.Value)));
            var srgb = outline.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault();
            var color = srgb?.Val?.Value;
            if (!string.IsNullOrEmpty(color) && System.Text.RegularExpressions.Regex.IsMatch(color, "^[0-9A-Fa-f]{6}$"))
            {
                var dash = outline.GetFirstChild<DocumentFormat.OpenXml.Drawing.PresetDash>();
                var style = "solid";
                if (dash?.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash) style = "dashed";
                else if (dash?.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot) style = "dotted";
                borderAttrs = $" data-border-width=\"{borderWidthPx}\" data-border-color=\"#{color}\" data-border-style=\"{style}\"";
            }
        }

        // Optional crop (a:srcRect) — l/t/r/b in 1/1000 of a percent → percent.
        var cropAttrs = string.Empty;
        var srcRect = drawing.Descendants<DocumentFormat.OpenXml.Drawing.SourceRectangle>().FirstOrDefault();
        if (srcRect != null)
        {
            var l = (srcRect.Left?.Value ?? 0) / 1000;
            var r = (srcRect.Right?.Value ?? 0) / 1000;
            var t = (srcRect.Top?.Value ?? 0) / 1000;
            var b = (srcRect.Bottom?.Value ?? 0) / 1000;
            if (l > 0 || r > 0 || t > 0 || b > 0)
            {
                cropAttrs = $" data-crop-l=\"{l}\" data-crop-r=\"{r}\" data-crop-t=\"{t}\" data-crop-b=\"{b}\"";
            }
        }

        // Alt text from wp:docPr (@descr preferred, else @title) — preserved as <img alt>.
        var docPr = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
        var alt = docPr?.Description?.Value ?? docPr?.Title?.Value;
        var altAttr = !string.IsNullOrEmpty(alt) ? $" alt=\"{EscapeHtml(alt)}\"" : string.Empty;

        return $"<img src=\"data:{contentType};base64,{base64Data}\" " +
               $"style=\"max-width:100%;width:{width}px;height:{height}px;\" " +
               $"data-image-id=\"{relationshipId}\" " +
               $"data-width-emu=\"{widthEmu}\" data-height-emu=\"{heightEmu}\"" +
               $"{altAttr}{posAttrs}{borderAttrs}{cropAttrs} />";
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
        var hasExplicitWidth = false;
        if (tableProps?.TableWidth?.Width?.Value != null)
        {
            var w = tableProps.TableWidth;
            if (w.Type?.Value == TableWidthUnitValues.Pct)
            {
                tableWidth = $"{int.Parse(w.Width.Value) / 50}%";
                hasExplicitWidth = true;
            }
            else if (w.Type?.Value == TableWidthUnitValues.Dxa)
            {
                tableWidth = $"{TwipsToPx(int.Parse(w.Width.Value))}px";
                hasExplicitWidth = true;
            }
        }

        // Authoritative column widths come from tblGrid. Emitting a <colgroup> plus a
        // fixed table layout makes the browser honour Word's column geometry instead of
        // sizing columns to content (the usual cause of "table looks nothing like Word").
        var gridColumnsPx = ReadTableGridColumnsPx(table);
        var isFixedLayout = tableProps?.TableLayout?.Type?.Value == TableLayoutValues.Fixed;
        var useFixedLayout = isFixedLayout || hasExplicitWidth;

        // When a fixed-layout table declares no explicit width, fall back to the grid sum
        // so the fixed layout has a width to distribute across the columns.
        if (useFixedLayout && tableWidth == "auto" && gridColumnsPx.Count > 0)
            tableWidth = $"{gridColumnsPx.Sum()}px";

        var layoutCss = useFixedLayout ? "table-layout:fixed;" : string.Empty;
        var colgroupHtml = BuildColgroupHtml(gridColumnsPx);

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

        // Sygnał dla HtmlToDocx: czy tabela ma jakiekolwiek zdefiniowane tblBorders.
        // Gdy wszystkie strony + inside to brak/None/Nil lub brak definicji — kod eksportu ma NIE
        // wymuszać solid-black tblBorders (co powodowałoby fałszywe czarne linie).
        var tblBordersAreEmpty = IsTableBordersEmpty(defaultBorders);
        var tblBordersMarker = tblBordersAreEmpty ? " data-no-borders=\"1\"" : "";
        
        // Domyślny padding komórek = domyślne marginesy komórki Worda (TableNormal):
        // top/bottom = 0, left/right = 108 twips. Wcześniejszy "4px 8px" dodawał 4px góra/dół
        // do KAŻDEJ komórki, przez co tabele rosły w pionie po round-tripie (każdy wiersz +~120
        // twips). Word domyślnie ma 0 góra/dół. Patrz analiza orginał_GOOD vs zapisany_BAD.
        const int wordDefaultCellMarginTwips = 108; // 0.19 cm — domyślny lewy/prawy margines komórki
        var defaultPadding = $"0px {TwipsToPx(wordDefaultCellMarginTwips)}px";
        var tblCellMar = tableProps?.TableCellMarginDefault;
        if (tblCellMar != null)
        {
            var topPad = GetTwipsValue(tblCellMar.TopMargin) ?? 0;
            var bottomPad = GetTwipsValue(tblCellMar.BottomMargin) ?? 0;
            var leftPad = GetDxaValue(tblCellMar.TableCellLeftMargin) ?? wordDefaultCellMarginTwips;
            var rightPad = GetDxaValue(tblCellMar.TableCellRightMargin) ?? wordDefaultCellMarginTwips;
            defaultPadding = $"{TwipsToPx(topPad)}px {TwipsToPx(rightPad)}px {TwipsToPx(bottomPad)}px {TwipsToPx(leftPad)}px";
        }
        
        html.Append($"<table{tblBordersMarker} style=\"border-collapse:collapse;width:{tableWidth};margin:4px 0;{layoutCss}{tableAlign}{tableIndent}\">");
        html.Append(colgroupHtml);

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
            
            // Iteruj komórki uwzględniając komórki opakowane w SDT (Content Control / formant).
            // SdtCell zawiera SdtContentCell, a w nim faktyczne TableCell — inaczej znikają dane.
            foreach (var cellLike in row.Elements())
            {
                if (cellLike is TableCell cell)
                {
                    AppendTableCellHtml(html, table, row, cell, defaultBorders, defaultPadding, document, sourcePart);
                }
                else if (cellLike is SdtCell sdtCell)
                {
                    var sdtContent = sdtCell.GetFirstChild<SdtContentCell>();
                    if (sdtContent != null)
                    {
                        foreach (var innerCell in sdtContent.Elements<TableCell>())
                            AppendTableCellHtml(html, table, row, innerCell, defaultBorders, defaultPadding, document, sourcePart);
                    }
                }
            }
            
            html.Append("</tr>");
        }

        html.Append("</table>");
        return html.ToString();
    }

    /// <summary>
    /// Reads tblGrid column widths (twips) and converts them to CSS pixels.
    /// Columns without an explicit width contribute 0 (browser distributes remainder).
    /// </summary>
    private List<int> ReadTableGridColumnsPx(Table table)
    {
        var result = new List<int>();
        var grid = table.GetFirstChild<TableGrid>();
        if (grid == null) return result;

        foreach (var col in grid.Elements<GridColumn>())
        {
            if (col.Width?.Value != null && int.TryParse(col.Width.Value, out var twips))
                result.Add(TwipsToPx(twips));
            else
                result.Add(0);
        }
        return result;
    }

    private static string BuildColgroupHtml(List<int> columnsPx)
    {
        if (columnsPx.Count == 0) return string.Empty;

        var sb = new StringBuilder("<colgroup>");
        foreach (var px in columnsPx)
            sb.Append(px > 0 ? $"<col style=\"width:{px}px;\" />" : "<col />");
        sb.Append("</colgroup>");
        return sb.ToString();
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

        // Match continuation cells by their grid-column position, not by element index:
        // a preceding gridSpan shifts the column, so index-based matching breaks merges.
        var startColumn = GetCellStartColumn(startRow, startCell);

        var rowSpan = 1;
        for (int i = startRowIndex + 1; i < rows.Count; i++)
        {
            var cell = FindCellAtColumn(rows[i], startColumn);
            var vMerge = cell?.TableCellProperties?.VerticalMerge;
            if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                rowSpan++;
            else
                break;
        }

        return rowSpan;
    }

    private static int GetGridSpan(TableCell cell)
    {
        var gs = cell.TableCellProperties?.GridSpan?.Val?.Value;
        return gs is > 0 ? gs.Value : 1;
    }

    private static int GetCellStartColumn(TableRow row, TableCell target)
    {
        var column = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            if (ReferenceEquals(cell, target)) return column;
            column += GetGridSpan(cell);
        }
        return column;
    }

    private static TableCell? FindCellAtColumn(TableRow row, int targetColumn)
    {
        var column = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            if (column == targetColumn) return cell;
            column += GetGridSpan(cell);
        }
        return null;
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
        var v = border.Val?.Value;
        if (v == null || v == BorderValues.None || v == BorderValues.Nil) return "none";
        return GetBorderCss(border);
    }

    /// <summary>
    /// Sprawdza, czy TableBorders nie zawiera żadnego widocznego borderu
    /// (brak elementu, lub wszystkie strony mają Val=None/Nil albo brak Val).
    /// </summary>
    private static bool IsTableBordersEmpty(TableBorders? tb)
    {
        if (tb == null) return true;
        bool IsBlank(BorderType? b)
        {
            if (b == null) return true;
            var v = b.Val?.Value;
            return v == null || v == BorderValues.None || v == BorderValues.Nil;
        }
        return IsBlank(tb.TopBorder)
            && IsBlank(tb.BottomBorder)
            && IsBlank(tb.LeftBorder)
            && IsBlank(tb.RightBorder)
            && IsBlank(tb.InsideHorizontalBorder)
            && IsBlank(tb.InsideVerticalBorder);
    }

    private string ConvertSdtBlockToHtml(SdtBlock sdtBlock, WordprocessingDocument document)
    {
        var html = new StringBuilder();
        var content = sdtBlock.SdtContentBlock;
        if (content != null)
        {
            // Marker pozwala odróżnić zawartość pochodzącą z formantu w HTML.
            // Atrybuty data-* przechowują podstawowe metadane (tag/alias) na potrzeby
            // ewentualnego round-trip.
            var props = sdtBlock.SdtProperties;
            var tag = props?.Elements<Tag>().FirstOrDefault()?.Val?.Value ?? "";
            var alias = props?.Elements<SdtAlias>().FirstOrDefault()?.Val?.Value ?? "";
            var dataAttrs = new StringBuilder();
            if (!string.IsNullOrEmpty(tag)) dataAttrs.Append($" data-sdt-tag=\"{System.Net.WebUtility.HtmlEncode(tag)}\"");
            if (!string.IsNullOrEmpty(alias)) dataAttrs.Append($" data-sdt-alias=\"{System.Net.WebUtility.HtmlEncode(alias)}\"");
            html.Append($"<div class=\"sdt-block\"{dataAttrs}>");

            // Zbierz elementy i obsłuż kolejne paragrafy listy tak samo jak w body
            var elems = content.Elements().ToList();
            int i = 0;
            while (i < elems.Count)
            {
                var el = elems[i];
                if (el is Paragraph p && IsListParagraph(p))
                {
                    html.Append(ConvertConsecutiveListItems(elems, ref i, document));
                }
                else
                {
                    html.Append(ConvertElementToHtml(el, document));
                    i++;
                }
            }

            html.Append("</div>");
        }
        return html.ToString();
    }

    /// <summary>
    /// Konwertuje inline-owy SDT (SdtRun — „formant” w treści paragrafu) na HTML.
    /// Bez tej obsługi cała zawartość formantu znikała przy ładowaniu dokumentu.
    /// </summary>
    private string ConvertSdtRunToHtml(SdtRun sdtRun, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var content = sdtRun.GetFirstChild<SdtContentRun>();
        if (content == null) return string.Empty;

        var props = sdtRun.SdtProperties;
        var tag = props?.Elements<Tag>().FirstOrDefault()?.Val?.Value ?? "";
        var alias = props?.Elements<SdtAlias>().FirstOrDefault()?.Val?.Value ?? "";
        var dataAttrs = new StringBuilder();
        if (!string.IsNullOrEmpty(tag)) dataAttrs.Append($" data-sdt-tag=\"{System.Net.WebUtility.HtmlEncode(tag)}\"");
        if (!string.IsNullOrEmpty(alias)) dataAttrs.Append($" data-sdt-alias=\"{System.Net.WebUtility.HtmlEncode(alias)}\"");

        var inner = new StringBuilder();
        foreach (var el in content.Elements())
        {
            switch (el)
            {
                case Run run:
                    inner.Append(ConvertRunToHtml(run, document, sourcePart));
                    break;
                case Hyperlink hl:
                    inner.Append(ConvertHyperlinkToHtml(hl, document));
                    break;
                case SimpleField sf:
                    inner.Append(ConvertSimpleFieldToHtml(sf));
                    break;
                case SdtRun nested:
                    inner.Append(ConvertSdtRunToHtml(nested, document, sourcePart));
                    break;
            }
        }

        // Pusty formant — wstaw &nbsp; żeby kursor miał się gdzie ustawić.
        if (inner.Length == 0) inner.Append("&nbsp;");

        return $"<span class=\"sdt-inline\"{dataAttrs}>{inner}</span>";
    }

    /// <summary>
    /// Pomocnicza: emituje pojedynczy &lt;td&gt; ze stylami, kolspanami i zawartością.
    /// Wydzielona, by móc ją wołać zarówno dla zwykłej TableCell jak i z SdtContentCell.
    /// </summary>
    private void AppendTableCellHtml(
        StringBuilder html,
        Table table,
        TableRow row,
        TableCell cell,
        TableBorders? defaultBorders,
        string defaultPadding,
        WordprocessingDocument document,
        OpenXmlPart? sourcePart)
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
        else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
        {
            // An omitted vMerge val defaults to "continue" (ECMA-376) — drop the merged cell.
            return;
        }

        html.Append($"<td{colspan}{rowspan} style=\"{cellStyle}\">");

        // Iteruj wszystkie dzieci komórki, by obsłużyć też SdtBlock i Table osadzone bezpośrednio.
        foreach (var inner in cell.Elements())
        {
            switch (inner)
            {
                case Paragraph para:
                    html.Append(ConvertParagraphToHtml(para, document, sourcePart));
                    break;
                case Table nestedTable:
                    html.Append(ConvertTableToHtml(nestedTable, document, sourcePart));
                    break;
                case SdtBlock sdt:
                    html.Append(ConvertSdtBlockToHtml(sdt, document));
                    break;
            }
        }

        html.Append("</td>");
    }

    private int TwipsToPx(int twips) => (int)OoxmlUnits.TwipsToPixels(twips);
    private int EmuToPx(long emu) => (int)OoxmlUnits.EmuToPixels(emu);
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

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
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
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

    // ── Liczniki numeracji list (semantyka Worda) ────────────────────────────────
    // Word utrzymuje licznik per ABSTRAKCYJNA definicja numeracji: różne w:num wskazujące ten sam
    // w:abstractNum KONTYNUUJĄ numerację (tak działa „Kontynuuj numerację"), chyba że dana instancja
    // ma w:lvlOverride/w:startOverride — wtedy licznik poziomu resetuje się przy PIERWSZYM użyciu tej
    // instancji („Rozpocznij od nowa"). Poziomy głębsze restartują po powrocie na poziom płytszy
    // (chyba że w:lvlRestart=0). Klucz: (abstractNumId, level) → ostatnio wyemitowany numer.
    private readonly Dictionary<(int abstractNumId, int level), int> _listCounters = new();
    // numId-y, których startOverride już zastosowano (reset tylko przy pierwszym użyciu instancji).
    private readonly HashSet<int> _appliedStartOverrides = new();

    // Firmowa czcionka — używana, gdy dokument nie definiuje własnej w docDefaults.
    private readonly DocumentDefaultsOptions _defaults;
    private readonly IGraphicConversionService _graphics;
    private readonly ILogger<DocxToHtmlConverter> _log;

    public DocxToHtmlConverter()
    {
        _defaults = new DocumentDefaultsOptions();
        _graphics = new GraphicConversionService();
        _log = NullLogger<DocxToHtmlConverter>.Instance;
    }

    public DocxToHtmlConverter(IOptions<DocumentDefaultsOptions> defaults, IGraphicConversionService? graphics = null,
        ILogger<DocxToHtmlConverter>? logger = null)
    {
        _defaults = defaults?.Value ?? new DocumentDefaultsOptions();
        _graphics = graphics ?? new GraphicConversionService();
        _log = logger ?? NullLogger<DocxToHtmlConverter>.Instance;
    }

    /// <summary>
    /// Gdy media part jest formatem nie-renderowalnym natywnie (EMF/WMF/TIFF), zwraca data:URL
    /// renderowalny w przeglądarce (osadzony/zdekodowany raster albo PRZEZROCZYSTY blank — nigdy
    /// widoczny placeholder) zamiast nierenderowalnego data:image/x-emf. Dla formatów web-native
    /// zwraca null (zostaje dotychczasowa ścieżka). Oryginalny part nie jest usuwany — pass-through
    /// zapewnia wierność w Word przy zapisie.
    /// </summary>
    private (string dataUrl, bool isBlank)? WebGraphicForLegacy(byte[] bytes, string? contentType, long widthEmu, long heightEmu, string? sourcePath = null)
    {
        // Unknown też przechodzi przez konwerter: pokrywa EMZ/WMZ (gzip), rastry rozpoznawalne
        // tylko przez Skia oraz parts z kłamiącym content-type. Bez tego `src` dostawał
        // nierenderowalny data:{contentType} i przeglądarka pokazywała ikonę złamanego obrazka.
        var kind = _graphics.Detect(bytes, contentType);
        if (kind is not (GraphicKind.Emf or GraphicKind.Wmf or GraphicKind.Tiff or GraphicKind.Unknown))
            return null; // web-native → zostaje dotychczasowa ścieżka
        var result = _graphics.ConvertForEditor(new GraphicSource
        {
            Data = bytes,
            ContentType = contentType,
            SourcePath = sourcePath,
            Origin = GraphicOrigin.LegacyDocxPart,
            TargetWidthEmu = widthEmu > 0 ? widthEmu : null,
            TargetHeightEmu = heightEmu > 0 ? heightEmu : null
        });
        if (result.Diagnostics.Status is GraphicConversionStatus.Fallback
            or GraphicConversionStatus.Unsupported or GraphicConversionStatus.Rejected)
        {
            _log.LogWarning(
                "Grafika bez pełnej konwersji web: part={SourcePath} declaredType={ContentType} detected={Kind} " +
                "size={Size}B status={Status} strategie=[{Strategies}] powód={Reason}",
                sourcePath, contentType, result.Diagnostics.InputKind, bytes.Length,
                result.Diagnostics.Status, string.Join(",", result.Diagnostics.AttemptedStrategies),
                result.Diagnostics.FailureReason);
        }
        return result.Web != null ? (result.Web.ToDataUrl(), result.Web.IsBlankFallback) : null;
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
        _listCounters.Clear();
        _appliedStartOverrides.Clear();
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
            PageSize = ExtractPageSize(document),
            SectionHeadersFooters = ExtractSectionHeadersFooters(document)
        };

        return content;
    }

    /// <summary>
    /// All w:sectPr in DOCUMENT order. In OOXML the sectPr that is a direct child of
    /// w:body describes the LAST section; every earlier section ends with a paragraph
    /// carrying its sectPr inside pPr. Callers that need "what the user sees on page 1"
    /// must take the FIRST element here — Body.Elements&lt;SectionProperties&gt;() alone
    /// silently returns the last section's geometry/references (R-10).
    /// </summary>
    private static List<SectionProperties> GetSectionPropertiesInDocumentOrder(Body? body)
    {
        var result = new List<SectionProperties>();
        if (body == null) return result;

        result.AddRange(body.Descendants<SectionProperties>()
            .Where(sp => sp.Parent is ParagraphProperties));

        var bodyLevel = body.Elements<SectionProperties>().FirstOrDefault();
        if (bodyLevel != null) result.Add(bodyLevel);
        return result;
    }

    private static SectionProperties? GetFirstSectionProperties(WordprocessingDocument document)
        => GetSectionPropertiesInDocumentOrder(document.MainDocumentPart?.Document?.Body).FirstOrDefault();

    /// <summary>
    /// w:sectPr/w:type of the given section — how the section BEGINS relative to the
    /// previous one. Absent w:type means a next-page break (Word default).
    /// </summary>
    private static string GetSectionBreakType(SectionProperties sectPr)
    {
        var type = sectPr.GetFirstChild<SectionType>()?.Val?.Value;
        if (type == null) return "nextPage";
        if (type == SectionMarkValues.Continuous) return "continuous";
        if (type == SectionMarkValues.OddPage) return "oddPage";
        if (type == SectionMarkValues.EvenPage) return "evenPage";
        if (type == SectionMarkValues.NextColumn) return "nextColumn";
        return "nextPage";
    }

    /// <summary>
    /// Marker końca sekcji dla edytora. <c>endedSection</c> to sectPr paragrafu kończącego
    /// sekcję; marker niesie geometrię sekcji NASTĘPNEJ (tej, która zaczyna się za nim) —
    /// wartości w data-* (cm, InvariantCulture). Dla przerw zaczynających nową stronę
    /// (nextPage/oddPage/evenPage) poprzedza go standardowy <c>div.page-break</c>, żeby
    /// edytor łamał stronę; sam marker jest osobnym elementem, bo splitter stron w GUI
    /// kanonizuje divy page-break i zgubiłby data-*.
    /// </summary>
    private static string BuildSectionBreakMarkerHtml(SectionProperties endedSection, List<SectionProperties> orderedSections)
    {
        var idx = orderedSections.IndexOf(endedSection);
        if (idx < 0 || idx + 1 >= orderedSections.Count) return string.Empty;

        var next = orderedSections[idx + 1];
        var breakType = GetSectionBreakType(next);
        var page = SectionPropertiesReader.ReadPageSettings(next);
        var inv = System.Globalization.CultureInfo.InvariantCulture;

        var sb = new StringBuilder();
        if (breakType is "nextPage" or "oddPage" or "evenPage")
            sb.Append("<div class=\"page-break\"></div>");

        sb.Append("<div class=\"docx-section-break\" data-break-type=\"").Append(breakType).Append('"');

        if (page.PageWidthTwips is { } w && page.PageHeightTwips is { } h)
        {
            sb.Append(string.Format(inv, " data-page-width-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(w)));
            sb.Append(string.Format(inv, " data-page-height-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(h)));
            sb.Append(page.Orientation == PageOrientation.Landscape
                ? " data-orientation=\"landscape\""
                : " data-orientation=\"portrait\"");
        }

        if (page.HasPageMargin)
        {
            // Top/Bottom jak w ExtractPageMargins: wartość bezwzględna (mirror/overlap).
            if (page.TopMarginTwips is { } t)
                sb.Append(string.Format(inv, " data-margin-top-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(Math.Abs(t))));
            if (page.BottomMarginTwips is { } b)
                sb.Append(string.Format(inv, " data-margin-bottom-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(Math.Abs(b))));
            if (page.LeftMarginTwips is { } l)
                sb.Append(string.Format(inv, " data-margin-left-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(l)));
            if (page.RightMarginTwips is { } r)
                sb.Append(string.Format(inv, " data-margin-right-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(r)));
            if (page.HeaderDistanceTwips is { } hd)
                sb.Append(string.Format(inv, " data-header-distance-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(hd)));
            if (page.FooterDistanceTwips is { } fd)
                sb.Append(string.Format(inv, " data-footer-distance-cm=\"{0:0.##}\"", OoxmlUnits.TwipsToCm(fd)));
        }

        sb.Append("></div>");
        return sb.ToString();
    }

    /// <summary>
    /// Page size + orientation (cm) from the first section. Null when the section
    /// declares no w:pgSz (caller falls back to its own default).
    /// </summary>
    private static Domain.Models.PageSize? ExtractPageSize(WordprocessingDocument document)
    {
        var sectionProps = GetFirstSectionProperties(document);
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
        var sectionProps = GetFirstSectionProperties(document);
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

        var sections = GetSectionPropertiesInDocumentOrder(mainPart.Document?.Body);

        // Render the DEFAULT header (the one Word shows on ordinary pages), resolved via
        // sectPr/headerReference — NOT HeaderParts.FirstOrDefault(), whose order is
        // undefined and may return an empty even/first part. Sections are scanned in
        // document order (first section wins — that's what the user sees on page 1;
        // later sections inherit in Word when they declare no reference). Fall back to
        // the first available part only when no section declares a reference.
        var sectionProps = sections.FirstOrDefault(s =>
            ResolveHeaderPart(mainPart, s, HeaderFooterValues.Default) != null) ?? sections.FirstOrDefault();
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

        // Band geometry follows the FIRST section's page margins — the same section whose
        // margins/page size the rest of DocumentContent reports.
        var page = SectionPropertiesReader.ReadPageSettings(sections.FirstOrDefault());
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

        var sections = GetSectionPropertiesInDocumentOrder(mainPart.Document?.Body);

        // See ExtractHeader: resolve the DEFAULT footer via sectPr/footerReference rather
        // than FooterParts.FirstOrDefault(), which can return an empty even/first part.
        // First section with a reference wins (document order).
        var sectionProps = sections.FirstOrDefault(s =>
            ResolveFooterPart(mainPart, s, HeaderFooterValues.Default) != null) ?? sections.FirstOrDefault();
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

        var page = SectionPropertiesReader.ReadPageSettings(sections.FirstOrDefault());
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
    /// Własne nagłówki/stopki sekcji ≥ 1 (0-based, kolejność dokumentu). Wpis powstaje tylko,
    /// gdy sekcja deklaruje WŁASNE referencje — sekcje dziedziczące (bez referencji) nie mają
    /// wpisu i frontend rozwiązuje dziedziczenie jak Word (poprzednia sekcja). Sekcja 0 jest
    /// raportowana w polach Header/Footer (kompatybilność wstecz).
    /// </summary>
    private List<SectionHeaderFooter>? ExtractSectionHeadersFooters(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return null;

        var sections = GetSectionPropertiesInDocumentOrder(mainPart.Document?.Body);
        if (sections.Count < 2) return null;

        var result = new List<SectionHeaderFooter>();
        for (int i = 1; i < sections.Count; i++)
        {
            var header = ExtractHeaderOwnedBySection(mainPart, document, sections[i]);
            var footer = ExtractFooterOwnedBySection(mainPart, document, sections[i]);
            if (header != null || footer != null)
                result.Add(new SectionHeaderFooter { SectionIndex = i, Header = header, Footer = footer });
        }
        return result.Count > 0 ? result : null;
    }

    /// <summary>
    /// Nagłówek zadeklarowany bezpośrednio przez daną sekcję (bez fallbacku do innych
    /// części) + warianty first/even wg opt-inów tej sekcji. Geometria pasma z tej sekcji.
    /// </summary>
    private HeaderFooterContent? ExtractHeaderOwnedBySection(MainDocumentPart mainPart, WordprocessingDocument document, SectionProperties sectionProps)
    {
        var headerPart = ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.Default);
        if (headerPart?.Header == null) return null;

        var html = ConvertHeaderPartToHtml(headerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        string? firstPageHtml = null;
        var differentFirstPage = false;
        if (HasTitlePage(sectionProps)
            && ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.First) is { Header: not null } firstPart
            && ConvertHeaderPartToHtml(firstPart, document) is { Length: > 0 } fph && !string.IsNullOrWhiteSpace(fph))
        {
            firstPageHtml = fph;
            differentFirstPage = true;
        }

        string? evenHtml = null;
        var differentOddEven = false;
        if (HasEvenAndOddHeaders(mainPart)
            && ResolveHeaderPart(mainPart, sectionProps, HeaderFooterValues.Even) is { Header: not null } evenPart
            && ConvertHeaderPartToHtml(evenPart, document) is { Length: > 0 } eh && !string.IsNullOrWhiteSpace(eh))
        {
            evenHtml = eh;
            differentOddEven = true;
        }

        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        var height = page.HasPageMargin ? ComputeBandHeightCm(page.TopMarginTwips, page.HeaderDistanceTwips) : 1.5;

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, height)),
            DifferentFirstPage = differentFirstPage,
            FirstPageHtml = firstPageHtml,
            DifferentOddEven = differentOddEven,
            EvenHtml = evenHtml
        };
    }

    private HeaderFooterContent? ExtractFooterOwnedBySection(MainDocumentPart mainPart, WordprocessingDocument document, SectionProperties sectionProps)
    {
        var footerPart = ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.Default);
        if (footerPart?.Footer == null) return null;

        var html = ConvertFooterPartToHtml(footerPart, document);
        if (string.IsNullOrWhiteSpace(html)) return null;

        string? firstPageHtml = null;
        var differentFirstPage = false;
        if (HasTitlePage(sectionProps)
            && ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.First) is { Footer: not null } firstPart
            && ConvertFooterPartToHtml(firstPart, document) is { Length: > 0 } fph && !string.IsNullOrWhiteSpace(fph))
        {
            firstPageHtml = fph;
            differentFirstPage = true;
        }

        string? evenHtml = null;
        var differentOddEven = false;
        if (HasEvenAndOddHeaders(mainPart)
            && ResolveFooterPart(mainPart, sectionProps, HeaderFooterValues.Even) is { Footer: not null } evenPart
            && ConvertFooterPartToHtml(evenPart, document) is { Length: > 0 } eh && !string.IsNullOrWhiteSpace(eh))
        {
            evenHtml = eh;
            differentOddEven = true;
        }

        var page = SectionPropertiesReader.ReadPageSettings(sectionProps);
        var height = page.HasPageMargin ? ComputeBandHeightCm(page.BottomMarginTwips, page.FooterDistanceTwips) : 1.5;

        return new HeaderFooterContent
        {
            Html = html,
            Height = Math.Max(0.8, Math.Min(8, height)),
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
        var orderedSections = GetSectionPropertiesInDocumentOrder(body);
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

                // Paragraf z pPr/sectPr KOŃCZY sekcję. Emitujemy niewidoczny marker sekcji
                // z geometrią NASTĘPNEJ sekcji (rozmiar/orientacja/marginesy) + zwykły
                // page-break, gdy przerwa zaczyna nową stronę. Marker niesie dane w data-*
                // i wraca w autosave — HtmlToDocxConverter odtwarza z niego w:sectPr, więc
                // dokument wielosekcyjny nie jest już spłaszczany do jednej sekcji (R-10).
                if (element is Paragraph sectionEnd &&
                    sectionEnd.ParagraphProperties?.GetFirstChild<SectionProperties>() is { } endedSection)
                {
                    html.Append(BuildSectionBreakMarkerHtml(endedSection, orderedSections));
                }
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

        // `start` = FAKTYCZNY numer pierwszego elementu wg liczników Worda (kontynuacja po przerwaniu
        // akapitem / współdzielony abstrakt), nie sama definicja w:start. Konsumpcja w pętli niżej.
        var startNumber = listType == "ol" ? PeekNextListNumber(firstNumId, firstLevel) : 1;
        var startAttr = (listType == "ol" && startNumber > 1) ? $" start=\"{startNumber}\"" : "";

        // Tożsamość i definicja listy w data-* — HtmlToDocxConverter odtwarza z nich w:numPr
        // (wspólny numId dla kontynuacji) i w:abstractNum (format/lvlText/start per poziom).
        var identityAttrs = new StringBuilder();
        identityAttrs.Append($" data-num-id=\"{firstNumId}\"");
        var abstractId = ResolveAbstractNumId(firstNumId);
        if (abstractId >= 0) identityAttrs.Append($" data-abstract-num-id=\"{abstractId}\"");
        identityAttrs.Append($" data-ilvl=\"{firstLevel}\"");
        identityAttrs.Append($" data-num-fmt=\"{firstInfo.FmtToken}\"");
        if (firstInfo.Start > 1) identityAttrs.Append($" data-start=\"{firstInfo.Start}\"");
        if (firstInfo.LvlText != null)
            identityAttrs.Append($" data-lvl-text=\"{System.Net.WebUtility.HtmlEncode(firstInfo.LvlText)}\"");
        if (!string.IsNullOrEmpty(firstInfo.BulletFont))
            identityAttrs.Append($" data-bullet-font=\"{System.Net.WebUtility.HtmlEncode(firstInfo.BulletFont)}\"");

        html.Append($"<{listType}{startAttr}{identityAttrs} style=\"{listStyleCss}\">");

        while (index < elements.Count)
        {
            if (elements[index] is not Paragraph p || !IsListParagraph(p))
                break;

            var (currentNumId, _) = GetEffectiveNumberingInfo(p);
            var currentLevel = GetListLevel(p);

            // Inny numId na tym samym/płytszym poziomie = INNA lista (logiczna tożsamość, nie wygląd).
            // Niezależne listy o identycznym formacie nie są już sklejane; kontynuację tej samej
            // logicznej listy (współdzielony abstrakt, brak startOverride) zapewniają liczniki
            // (`start` na kolejnym elemencie), a wspólny data-num-id scala je z powrotem przy zapisie.
            if (currentNumId != firstNumId && currentLevel <= firstLevel)
                break;

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
                // Skonsumuj licznik numeracji (semantyka Worda): element na tym poziomie nadaje
                // kolejny numer i restartuje poziomy głębsze (chyba że w:lvlRestart=0). Dotyczy
                // także punktorów — element płytszy restartuje głębsze poziomy numerowane.
                NextListNumber(currentNumId, currentLevel);

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

    /// <summary>
    /// rId-y są unikalne wyłącznie W OBRĘBIE jednej części pakietu — main, każdy nagłówek i każda
    /// stopka mają WŁASNE przestrzenie relacji zaczynające się od rId1. Cache obrazów musi więc być
    /// kluczowany częścią + rId; sam rId powodował, że obraz nagłówka o kolidującym rId renderował
    /// obraz z body (lub odwrotnie).
    /// </summary>
    private static string ImageCacheKey(OpenXmlPart part, string relationshipId)
        => $"{part.Uri}|{relationshipId}";

    private void LoadImageFromPart(OpenXmlPart part, ImagePart imagePart)
    {
        var relationshipId = part.GetIdOfPart(imagePart);
        var cacheKey = ImageCacheKey(part, relationshipId);
        if (_images.ContainsKey(cacheKey)) return;

        using var stream = imagePart.GetStream();
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);

        var rawBytes = memoryStream.ToArray();
        var contentType = NormalizeImageContentType(imagePart.ContentType);

        // SVG is an ACTIVE format: even embedded via <img src="data:…"> browsers render it inertly,
        // but the whole document HTML is trusted downstream, so we sanitise defensively (strip
        // script/foreignObject/on*-handlers/external refs) before embedding. A part that is not
        // valid SVG, cannot be sanitised, or exceeds the size guard is dropped (rejected) rather
        // than embedded raw — no broken/unsafe image reaches the editor.
        if (IsSvgContentType(contentType))
        {
            if (rawBytes.Length > MaxSvgBytes)
            {
                _log.LogWarning("SVG part pominięty (za duży): part={PartUri} size={Size}B limit={Limit}B",
                    imagePart.Uri, rawBytes.Length, MaxSvgBytes);
                return;
            }
            var sanitized = _graphics.SanitizeSvg(System.Text.Encoding.UTF8.GetString(rawBytes));
            if (sanitized == null)
            {
                _log.LogWarning("SVG part pominięty (niepoprawny/niebezpieczny): part={PartUri}", imagePart.Uri);
                return;
            }
            _images[cacheKey] = new DocumentImage
            {
                Id = relationshipId,
                ContentType = "image/svg+xml",
                Base64Data = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sanitized))
            };
            return;
        }

        // Nie-natywne formaty (EMF/WMF/TIFF): konwersja pure-managed (bez LibreOffice/System.Drawing
        // → identyczne zachowanie na Windows i Linux/GCP). Gdy metafile zawiera osadzony/odczytywalny
        // raster — wyciągamy go (PNG); w innym wypadku zostaje oryginał, a renderer (WebGraphicForLegacy)
        // pokaże przezroczysty blank. Oryginalny part i tak jedzie do DOCX przez pass-through.
        if (IsNonBrowserNativeContentType(contentType))
        {
            var converted = _graphics.ConvertForEditor(new GraphicSource
            {
                Data = rawBytes,
                ContentType = contentType,
                SourcePath = imagePart.Uri?.ToString(),
                Origin = GraphicOrigin.LegacyDocxPart
            });
            if (converted.Web is { IsBlankFallback: false } w && w.MimeType != "image/svg+xml")
            {
                rawBytes = w.Data;       // osadzony/zdekodowany raster (PNG/JPEG)
                contentType = w.MimeType;
            }
            // SVG (tłumaczenie wektorowe) NIE podmienia bajtów w _images — oryginalny metafile
            // musi zostać, żeby renderer dał go do data-original-src (round-trip do DOCX);
            // podgląd SVG powstaje w WebGraphicForLegacy z cache po hashu treści.
            else if (converted.Diagnostics.Status is GraphicConversionStatus.Fallback
                     or GraphicConversionStatus.Unsupported or GraphicConversionStatus.Rejected)
            {
                _log.LogWarning(
                    "Media part bez rastra web: part={PartUri} relId={RelId} declaredType={ContentType} " +
                    "size={Size}B status={Status} powód={Reason}",
                    imagePart.Uri, relationshipId, imagePart.ContentType, rawBytes.Length,
                    converted.Diagnostics.Status, converted.Diagnostics.FailureReason);
            }
        }

        _images[cacheKey] = new DocumentImage
        {
            Id = relationshipId,
            ContentType = contentType,
            Base64Data = System.Convert.ToBase64String(rawBytes)
        };
    }

    /// <summary>Górny limit rozmiaru SVG przyjmowanego do podglądu (anti-DoS).</summary>
    private const int MaxSvgBytes = 2 * 1024 * 1024;

    /// <summary>
    /// Normalizuje jednoznacznie rozpoznawalne, błędne typy MIME obrazów do formy standardowej.
    /// Kluczowy przypadek: <c>img/svg+xml</c> (spotykane w danych źródłowych) → <c>image/svg+xml</c>,
    /// bez którego przeglądarka nie rozpoznaje data-URI i obraz się nie renderuje. Świadomie wąskie:
    /// nie „naprawiamy" dowolnych typów, tylko ten jeden literał.
    /// </summary>
    private static string NormalizeImageContentType(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType)) return string.Empty;
        var ct = contentType.Trim();
        if (ct.Equals("img/svg+xml", StringComparison.OrdinalIgnoreCase)
            || ct.Equals("image/svg", StringComparison.OrdinalIgnoreCase))
            return "image/svg+xml";
        return ct;
    }

    private static bool IsSvgContentType(string? contentType)
        => !string.IsNullOrEmpty(contentType)
           && contentType.Contains("svg", StringComparison.OrdinalIgnoreCase);

    private static bool IsNonBrowserNativeContentType(string? contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return false;
        var ct = contentType.ToLowerInvariant();
        return ct.Contains("emf") || ct.Contains("wmf") || ct.Contains("metafile")
            || ct.Contains("tiff") || ct.Contains("tif")
            || ct.Contains("emz") || ct.Contains("wmz");   // skompresowane metafile (gzip)
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
                    if (IsNonBrowserNativeContentType(contentType))
                    {
                        // Pure-managed (bez LibreOffice/System.Drawing). Osadzony raster → użyj go;
                        // inaczej placeholder SVG jako punktator (bez crasha na Linux/GCP).
                        var conv = _graphics.ConvertForEditor(new GraphicSource
                        {
                            Data = bytes, ContentType = contentType, Origin = GraphicOrigin.LegacyDocxPart
                        });
                        if (conv.Web != null)
                        {
                            _picBulletDataUris[id] = conv.Web.ToDataUrl();
                            continue;
                        }
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
        
        // Tab-stopy: efektywne pozycje (styl + direct pPr). Zawsze serializowane do
        // data-tab-stops (round-trip per akapit — writer odtwarza w:tabs). W nagłówku/stopce
        // akapit z tabulatorami renderuje się POZYCYJNIE: segmenty lądują dokładnie na
        // pozycjach stopów (center = wyśrodkowany NA pozycji, right = kończy się NA pozycji),
        // jak w Wordzie. Flex (przybliżenie 50%/100%) zostaje dla body i braku pozycji.
        var effectiveTabStops = GetEffectiveTabStops(paraProps);
        var hasComplexField = paragraph.Descendants<FieldChar>().Any();
        // Positional tab rendering honours the REAL tab-stop positions (left aligns the following
        // segment's start, right aligns its end, center centres it — the semantic difference Word
        // draws). Applied to body paragraphs as well as header/footer: a flex row only spreads
        // segments evenly and ignores where the stops actually sit, so left/right tabs in the body
        // collapsed to equal gaps. Complex fields keep the legacy path (their runs are stateful).
        var usePositionedTabs = effectiveTabStops.Count > 0
            && paragraph.Descendants<TabChar>().Any()
            && !hasComplexField;

        // Fallback flex row only when there are tab characters but no resolvable stop positions
        // (e.g. a center/right alignment tab with no w:tabs geometry). Tab characters are preserved
        // either way (round-trip stays intact).
        var useFlexTabs = !usePositionedTabs && ParagraphHasAlignmentTab(paraProps);
        if (useFlexTabs)
            cssBuilder.Append("display:flex;align-items:baseline;width:100%;");
        if (usePositionedTabs)
            cssBuilder.Append("position:relative;");

        var tabStopsAttr = effectiveTabStops.Count > 0
            ? $" data-tab-stops=\"{SerializeTabStops(effectiveTabStops)}\""
            : string.Empty;

        var cssStyle = cssBuilder.ToString();
        var classAttr = docClass != null ? $" class=\"{docClass}\"" : string.Empty;
        // data-style-id pozwala eksporterowi HTML→DOCX odtworzyć oryginalny styleId (np. Title, Subtitle),
        // nawet jeśli wizualny tag to <p>.
        var dataStyleAttr = !string.IsNullOrEmpty(styleId) && docClass != null
            ? $" data-style-id=\"{System.Net.WebUtility.HtmlEncode(styleId)}\""
            : string.Empty;

        // Word „podział strony przed" (w:pageBreakBefore w pPr) — bardzo częsty sposób wymuszania
        // nowej strony (checkbox w dialogu akapitu, style nagłówków). Reader wcześniej go IGNOROWAŁ,
        // więc wielostronicowe dokumenty zwijały się do jednej strony (Issue: „3 strony → 1").
        // Emitujemy marker page-break PRZED akapitem (ten sam mechanizm co manualny break; writer
        // mapuje marker → w:br type=page). Pomijamy listy (marker między <li> = niepoprawny HTML).
        if (!isListItem && HasPageBreakBefore(paraProps))
            html.Append("<div class=\"page-break\"></div>");

        if (isListItem)
        {
            html.Append($"<li{classAttr}{dataStyleAttr}{tabStopsAttr} style=\"{cssStyle}\">");
        }
        else
        {
            html.Append($"<{tag}{classAttr}{dataStyleAttr}{tabStopsAttr} style=\"{cssStyle}\">");
        }

        var prevFlexTabs = _flexTabs;
        _flexTabs = useFlexTabs;

        // Obsługa złożonych pól (FieldChar Begin/Separate/End)
        if (hasComplexField)
        {
            html.Append(ConvertComplexFieldParagraphContent(paragraph, document, sourcePart));
        }
        else if (usePositionedTabs)
        {
            html.Append(BuildPositionedTabContent(paragraph, effectiveTabStops, document, sourcePart));
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
    /// True, gdy akapit ma ustawione <c>w:pageBreakBefore</c> (brak val = true; val=false/0 = false).
    /// Word używa tego do wymuszenia startu akapitu od nowej strony.
    /// </summary>
    private static bool HasPageBreakBefore(ParagraphProperties? paraProps)
    {
        var pbb = paraProps?.GetFirstChild<PageBreakBefore>();
        if (pbb == null) return false;
        return pbb.Val == null || pbb.Val.Value;
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

    /// <summary>Jeden efektywny tab-stop akapitu (pozycja w twips, wyrównanie, leader).</summary>
    private sealed record TabStopInfo(int PositionTwips, string Alignment, string? Leader);

    /// <summary>
    /// Efektywne tab-stopy akapitu: łańcuch stylów (od bazy do liścia), potem direct pPr.
    /// Późniejsza definicja na tej samej pozycji nadpisuje wcześniejszą; w:val=clear usuwa
    /// stop odziedziczony ze stylu (tak działa Word).
    /// </summary>
    private List<TabStopInfo> GetEffectiveTabStops(ParagraphProperties? paraProps)
    {
        var result = new List<TabStopInfo>();

        void Apply(Tabs? tabs)
        {
            if (tabs == null) return;
            foreach (var t in tabs.Elements<TabStop>())
            {
                if (t.Position?.Value is not { } pos) continue;
                result.RemoveAll(x => x.PositionTwips == pos);
                var val = t.Val?.Value;
                if (val == TabStopValues.Clear) continue;
                // Bar-tab rysuje pionową linię, nie pozycjonuje tekstu — pomijamy.
                if (val == TabStopValues.Bar) continue;
                result.Add(new TabStopInfo(pos, MapTabAlignment(val), MapTabLeader(t.Leader?.Value)));
            }
        }

        foreach (var style in GetParagraphStyleChainRootFirst(paraProps?.ParagraphStyleId?.Val?.Value))
            Apply(style.StyleParagraphProperties?.GetFirstChild<Tabs>());
        Apply(paraProps?.GetFirstChild<Tabs>());

        result.Sort((a, b) => a.PositionTwips.CompareTo(b.PositionTwips));
        return result;
    }

    /// <summary>Łańcuch stylów akapitowych basedOn, od korzenia do wskazanego stylu.</summary>
    private IEnumerable<Style> GetParagraphStyleChainRootFirst(string? styleId)
    {
        var chain = new List<Style>();
        var visited = new HashSet<string>();
        while (styleId != null && visited.Add(styleId) && _rawStyles.TryGetValue(styleId, out var style))
        {
            chain.Add(style);
            styleId = style.BasedOn?.Val?.Value;
        }
        chain.Reverse();
        return chain;
    }

    private static string MapTabAlignment(TabStopValues? val)
    {
        if (val == TabStopValues.Center) return "center";
        if (val == TabStopValues.Right || val == TabStopValues.End) return "right";
        if (val == TabStopValues.Decimal) return "decimal";
        return "left";
    }

    private static string? MapTabLeader(TabStopLeaderCharValues? leader)
    {
        if (leader == null || leader == TabStopLeaderCharValues.None) return null;
        if (leader == TabStopLeaderCharValues.Dot) return "dot";
        if (leader == TabStopLeaderCharValues.Hyphen) return "hyphen";
        if (leader == TabStopLeaderCharValues.Underscore) return "underscore";
        if (leader == TabStopLeaderCharValues.MiddleDot) return "middleDot";
        if (leader == TabStopLeaderCharValues.Heavy) return "heavy";
        return null;
    }

    /// <summary>Format atrybutu: "pos:align" lub "pos:align:leader", rozdzielane średnikami.</summary>
    private static string SerializeTabStops(List<TabStopInfo> stops) =>
        string.Join(";", stops.Select(s => s.Leader == null
            ? $"{s.PositionTwips}:{s.Alignment}"
            : $"{s.PositionTwips}:{s.Alignment}:{s.Leader}"));

    /// <summary>
    /// Rendering pozycyjny akapitu z tab-stopami (nagłówek/stopka): treść dzielona na segmenty
    /// na KAŻDYM tabulatorze; segment 0 zostaje w przepływie (definiuje wysokość linii),
    /// k-ty segment jest pozycjonowany absolutnie na k-tym stopie — center przez
    /// translateX(-50%) (tekst wyśrodkowany NA pozycji), right przez translateX(-100%)
    /// (tekst kończy się NA pozycji). Nadmiarowe segmenty (więcej tabów niż stopów) płyną
    /// inline. Wrapper formatowania runu jest domykany i otwierany wokół każdego segmentu.
    /// </summary>
    private string BuildPositionedTabContent(Paragraph paragraph, List<TabStopInfo> stops,
        WordprocessingDocument document, OpenXmlPart? sourcePart)
    {
        var segments = new List<StringBuilder> { new() };

        foreach (var child in paragraph.Elements())
        {
            switch (child)
            {
                case Run run:
                    var flags = GetRunSemanticFlags(run.RunProperties);
                    var (prefix, suffix) = BuildRunWrapper(run.RunProperties,
                        flags.Bold, flags.Italic, flags.Underline, flags.Strike, flags.Sup, flags.Sub);
                    var chunk = new StringBuilder();
                    void FlushChunk()
                    {
                        if (chunk.Length > 0)
                        {
                            segments[^1].Append(prefix).Append(chunk).Append(suffix);
                            chunk.Clear();
                        }
                    }
                    foreach (var rc in run.Elements())
                    {
                        if (rc is TabChar)
                        {
                            FlushChunk();
                            segments.Add(new StringBuilder());
                        }
                        else
                        {
                            chunk.Append(ConvertRunChildToHtml(rc, document, sourcePart));
                        }
                    }
                    FlushChunk();
                    break;
                case Hyperlink hyperlink:
                    segments[^1].Append(ConvertHyperlinkToHtml(hyperlink, document));
                    break;
                case SimpleField simpleField:
                    segments[^1].Append(ConvertSimpleFieldToHtml(simpleField));
                    break;
                case SdtRun sdtRun:
                    segments[^1].Append(ConvertSdtRunToHtml(sdtRun, document, sourcePart));
                    break;
            }
        }

        var html = new StringBuilder();
        // Strut: pusty segment 0 (akapit zaczyna się tabem) nie dawałby linii wysokości.
        html.Append(segments[0].Length > 0 ? segments[0].ToString() : "&#8203;");

        for (int k = 1; k < segments.Count; k++)
        {
            var stop = k - 1 < stops.Count ? stops[k - 1] : null;
            if (stop == null)
            {
                html.Append(segments[k]);
                continue;
            }
            var leftPx = TwipsToPx(stop.PositionTwips);
            var transform = stop.Alignment switch
            {
                "center" => "transform:translateX(-50%);",
                "right" => "transform:translateX(-100%);",
                _ => string.Empty
            };
            html.Append($"<span class=\"docx-tab-seg\" data-tab-align=\"{stop.Alignment}\" " +
                        $"style=\"position:absolute;left:{leftPx}px;{transform}white-space:pre;\">")
                .Append(segments[k])
                .Append("</span>");
        }

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

        // w:sz to 1/8 pt; px = pt × 96/72, czyli sz/6 (wcześniejsze sz/8 zaniżało grubość o 25%).
        // Minimum 0.5px, by hairline Worda (0.25–0.5 pt) pozostał widoczny w przeglądarce.
        var size = border.Size?.Value ?? 4;
        var sizePx = Math.Max(0.5, size / 6.0);
        var color = border.Color?.Value;
        if ((color == null || color == "auto") && border.ThemeColor?.HasValue == true)
        {
            var themeHex = ResolveThemeColor(border.ThemeColor.Value)?.TrimStart('#');
            if (themeHex != null)
                color = ApplyTintShade(themeHex, border.ThemeTint?.Value, border.ThemeShade?.Value);
        }
        if (color == null || color == "auto") color = "000000";

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
    /// Rozwiązuje abstractNumId dla numId, podążając za w:numStyleLink (abstrakt delegujący do
    /// stylu numeracji, którego pPr/numPr wskazuje inny numId → abstrakt). Zwraca -1, gdy brak.
    /// </summary>
    private int ResolveAbstractNumId(int numId)
    {
        if (_numberingPart?.Numbering == null) return -1;
        var visited = new HashSet<int>();
        while (visited.Add(numId))
        {
            var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID?.Value == numId);
            var absId = numInstance?.AbstractNumId?.Val?.Value;
            if (absId == null) return -1;

            var abs = _numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
            var linkedStyleId = abs?.GetFirstChild<NumberingStyleLink>()?.Val?.Value;
            if (string.IsNullOrEmpty(linkedStyleId))
                return absId.Value;

            // numStyleLink → styl numeracji → jego numPr/numId → kolejna iteracja.
            var linkedNumId = _rawStyles.TryGetValue(linkedStyleId, out var style)
                ? style.StyleParagraphProperties?.GetFirstChild<NumberingProperties>()?.NumberingId?.Val?.Value
                : null;
            if (linkedNumId == null) return absId.Value;
            numId = linkedNumId.Value;
        }
        return -1;
    }

    /// <summary>
    /// Definicja poziomu dla (numId, level): w:lvlOverride/w:lvl z instancji ma pierwszeństwo,
    /// potem w:lvl z abstraktu (po rozwiązaniu numStyleLink). Zwraca też startOverride (−1 = brak).
    /// </summary>
    private (Level? levelDef, int startOverride) FindLevelDefinition(int numId, int level)
    {
        if (_numberingPart?.Numbering == null) return (null, -1);

        var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return (null, -1);

        var levelOverrideElem = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(lo => lo.LevelIndex?.Value == level);
        int startOverride = levelOverrideElem?.StartOverrideNumberingValue?.Val?.Value ?? -1;

        Level? levelDef = levelOverrideElem?.GetFirstChild<Level>();
        if (levelDef == null)
        {
            var absId = ResolveAbstractNumId(numId);
            var abstractNum = _numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
            levelDef = abstractNum?.Elements<Level>()
                .FirstOrDefault(l => l.LevelIndex?.Value == level);
        }
        return (levelDef, startOverride);
    }

    /// <summary>Wartość początkowa poziomu: startOverride instancji, inaczej w:start, inaczej 1.</summary>
    private int GetLevelStart(int numId, int level)
    {
        var (levelDef, startOverride) = FindLevelDefinition(numId, level);
        if (startOverride > 0) return startOverride;
        return levelDef?.StartNumberingValue?.Val?.Value ?? 1;
    }

    /// <summary>Klucz licznika: abstrakt (współdzielony między instancjami); fallback per-numId.</summary>
    private int CounterKeyFor(int numId)
    {
        var absId = ResolveAbstractNumId(numId);
        return absId >= 0 ? absId : -numId;
    }

    /// <summary>
    /// „Rozpocznij od nowa" w Wordzie = nowa instancja ze startOverride na ten sam abstrakt.
    /// Reset licznika wykonujemy raz — przy pierwszym użyciu instancji w dokumencie.
    /// </summary>
    private void ApplyStartOverridesOnFirstUse(int numId, int counterKey)
    {
        if (!_appliedStartOverrides.Add(numId)) return;
        if (_numberingPart?.Numbering == null) return;

        var numInstance = _numberingPart.Numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return;

        foreach (var lo in numInstance.Elements<LevelOverride>())
        {
            var lvl = lo.LevelIndex?.Value;
            var so = lo.StartOverrideNumberingValue?.Val?.Value;
            if (lvl != null && so != null)
                _listCounters[(counterKey, lvl.Value)] = so.Value - 1;
        }
    }

    /// <summary>Numer, jaki dostanie następny element (numId, level) — bez konsumowania licznika.</summary>
    private int PeekNextListNumber(int numId, int level)
    {
        var key = CounterKeyFor(numId);
        ApplyStartOverridesOnFirstUse(numId, key);
        return _listCounters.TryGetValue((key, level), out var last)
            ? last + 1
            : GetLevelStart(numId, level);
    }

    /// <summary>
    /// Konsumuje kolejny numer dla (numId, level) i restartuje poziomy głębsze
    /// (domyślne zachowanie Worda; w:lvlRestart=0 wyłącza restart danego poziomu).
    /// </summary>
    private int NextListNumber(int numId, int level)
    {
        var key = CounterKeyFor(numId);
        ApplyStartOverridesOnFirstUse(numId, key);

        var next = _listCounters.TryGetValue((key, level), out var last)
            ? last + 1
            : GetLevelStart(numId, level);
        _listCounters[(key, level)] = next;

        for (int deeper = level + 1; deeper <= 8; deeper++)
        {
            if (!_listCounters.ContainsKey((key, deeper))) continue;
            var (deeperDef, _) = FindLevelDefinition(numId, deeper);
            var lvlRestart = deeperDef?.LevelRestart?.Val?.Value;
            if (lvlRestart == 0) continue; // nigdy nie restartuj
            _listCounters.Remove((key, deeper));
        }
        return next;
    }

    /// <summary>Token formatu numeracji do round-tripu w data-num-fmt (nazwy z w:numFmt).</summary>
    private static string NumFmtToken(Level? levelDef)
    {
        var fmt = levelDef?.NumberingFormat?.Val?.Value;
        if (fmt == NumberFormatValues.Decimal) return "decimal";
        if (fmt == NumberFormatValues.DecimalZero) return "decimalZero";
        if (fmt == NumberFormatValues.LowerLetter) return "lowerLetter";
        if (fmt == NumberFormatValues.UpperLetter) return "upperLetter";
        if (fmt == NumberFormatValues.LowerRoman) return "lowerRoman";
        if (fmt == NumberFormatValues.UpperRoman) return "upperRoman";
        if (fmt == NumberFormatValues.Bullet) return "bullet";
        if (fmt == NumberFormatValues.None) return "none";
        return "decimal";
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
        /// <summary>Token w:numFmt do round-tripu (data-num-fmt).</summary>
        public string FmtToken { get; init; }
        /// <summary>Surowy w:lvlText (np. "%1)" albo znak punktatora) do round-tripu.</summary>
        public string? LvlText { get; init; }
    }

    private ListLevelInfo GetListLevelInfo(NumberingProperties? numPr, int levelOverride = -1)
    {
        var fallback = new ListLevelInfo { Tag = "ul", ListStyleType = "disc", Start = 1, FmtToken = "bullet" };
        if (numPr == null || _numberingPart?.Numbering == null) return fallback;

        var numId = numPr.NumberingId?.Val?.Value;
        if (numId == null) return fallback;
        var level = levelOverride >= 0 ? levelOverride : (numPr.NumberingLevelReference?.Val?.Value ?? 0);

        // Wspólny resolver: lvlOverride/w:lvl instancji → w:lvl abstraktu (z numStyleLink).
        var (levelDef, startOverride) = FindLevelDefinition(numId.Value, level);
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
            Start = start,
            FmtToken = NumFmtToken(levelDef),
            LvlText = string.IsNullOrEmpty(levelText) ? null : levelText
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
                    // Rozróżnienie reguły dla round-tripu: bez tego markera writer mapował
                    // KAŻDE line-height w pt z powrotem na w:lineRule=exact, a atLeast→exact
                    // przycina w Wordzie tekst wyższy niż linia (np. większe glify, obrazki).
                    if (lineRule == LineSpacingRuleValues.AtLeast)
                        css.Append("--w-line-rule:atLeast;");
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

        var (prefix, suffix) = BuildRunWrapper(runProps, needsBold, needsItalic, needsUnderline, needsStrike, needsSup, needsSub);
        html.Append(prefix);
        foreach (var child in run.Elements())
        {
            html.Append(ConvertRunChildToHtml(child, document, sourcePart));
        }
        html.Append(suffix);

        return html.ToString();
    }

    /// <summary>
    /// Otwarcie/zamknięcie formatowania runu (span z CSS + tagi semantyczne). Wydzielone,
    /// bo segmentacja po tabulatorach (tab-stopy) musi domykać i ponownie otwierać ten sam
    /// wrapper wokół każdego segmentu runu.
    /// </summary>
    private (string Prefix, string Suffix) BuildRunWrapper(RunProperties? runProps,
        bool needsBold, bool needsItalic, bool needsUnderline, bool needsStrike, bool needsSup, bool needsSub)
    {
        var cleanCss = GetRunStyleClean(runProps);

        // Resolve a named character style (w:rStyle) and lay its inherited CSS *underneath*
        // the run's direct formatting (direct wins on conflicts, which come later in the
        // declaration). Without this, runs formatted only via a character style — Hyperlink,
        // Strong, Emphasis, custom — rendered with no formatting at all.
        var rStyleId = runProps?.RunStyle?.Val?.Value;
        var rStyleCss = rStyleId != null && _styles.TryGetValue(rStyleId, out var rsCss)
            ? rsCss
            : string.Empty;

        var prefix = new StringBuilder();
        prefix.Append($"<span style=\"{rStyleCss}{cleanCss}\">");
        if (needsBold) prefix.Append("<strong>");
        if (needsItalic) prefix.Append("<em>");
        if (needsUnderline) prefix.Append("<u>");
        if (needsStrike) prefix.Append("<s>");
        if (needsSup) prefix.Append("<sup>");
        if (needsSub) prefix.Append("<sub>");

        var suffix = new StringBuilder();
        if (needsSub) suffix.Append("</sub>");
        if (needsSup) suffix.Append("</sup>");
        if (needsStrike) suffix.Append("</s>");
        if (needsUnderline) suffix.Append("</u>");
        if (needsItalic) suffix.Append("</em>");
        if (needsBold) suffix.Append("</strong>");
        suffix.Append("</span>");

        return (prefix.ToString(), suffix.ToString());
    }

    private static (bool Bold, bool Italic, bool Underline, bool Strike, bool Sup, bool Sub) GetRunSemanticFlags(RunProperties? runProps)
    {
        if (runProps == null) return default;
        var bold = runProps.Bold != null && (runProps.Bold.Val == null || runProps.Bold.Val.Value);
        var italic = runProps.Italic != null && (runProps.Italic.Val == null || runProps.Italic.Val.Value);
        var underline = runProps.Underline != null && runProps.Underline.Val?.Value != UnderlineValues.None;
        var strike = (runProps.Strike != null && (runProps.Strike.Val == null || runProps.Strike.Val.Value)) ||
                     (runProps.DoubleStrike != null && (runProps.DoubleStrike.Val == null || runProps.DoubleStrike.Val.Value));
        var vertAlign = runProps.VerticalTextAlignment;
        var sup = vertAlign?.Val != null && vertAlign.Val.Value == VerticalPositionValues.Superscript;
        var sub = vertAlign?.Val != null && vertAlign.Val.Value == VerticalPositionValues.Subscript;
        return (bold, italic, underline, strike, sup, sub);
    }

    private string ConvertRunChildToHtml(OpenXmlElement child, WordprocessingDocument document, OpenXmlPart? sourcePart)
    {
        switch (child)
        {
            case Text text:
                return EscapeHtml(text.Text);
            case Break br:
                return br.Type?.Value == BreakValues.Page ? "<div class=\"page-break\"></div>" : "<br/>";
            case TabChar _:
                return "<span style=\"display:inline-block;min-width:2em;\">\t</span>";
            case Drawing drawing:
                return ConvertDrawingToHtml(drawing, document, sourcePart);
            case Picture picture:
                return ConvertPictureToHtml(picture, document, sourcePart);
            case AlternateContent alternate:
                return ConvertAlternateContentToHtml(alternate, document, sourcePart);
            case NoBreakHyphen _:
                return "&#8209;";
            case SoftHyphen _:
                return "&shy;";
            case SymbolChar sym:
                if (sym.Char?.Value != null)
                {
                    try { return $"&#x{sym.Char.Value};"; } catch { return string.Empty; }
                }
                return string.Empty;
            default:
                return string.Empty;
        }
    }

    /// <summary>
    /// Word owija nowsze rysunki (obrazy zakotwiczone z efektami, grupy, kanwy, kształty) w
    /// mc:AlternateContent: mc:Choice niesie nowoczesny markup (w:drawing), mc:Fallback wersję
    /// VML (w:pict) dla starych czytników. Bez tej gałęzi KAŻDY taki element znikał bez śladu
    /// (default w switchu → pusty string). Bierzemy pierwszą gałąź, z której da się
    /// wyprodukować HTML — Choice w kolejności dokumentu, potem Fallback.
    /// </summary>
    private string ConvertAlternateContentToHtml(AlternateContent alternate, WordprocessingDocument document, OpenXmlPart? sourcePart)
    {
        foreach (var branch in alternate.ChildElements)
        {
            if (branch is not (AlternateContentChoice or AlternateContentFallback)) continue;

            var html = new StringBuilder();
            foreach (var drawing in branch.Descendants<Drawing>())
                html.Append(ConvertDrawingToHtml(drawing, document, sourcePart));
            if (html.Length == 0)
            {
                foreach (var pict in branch.Descendants<Picture>())
                    html.Append(ConvertPictureToHtml(pict, document, sourcePart));
            }
            if (html.Length > 0) return html.ToString();
        }
        // Żadna gałąź nie dała obrazu — ostatnia szansa: pole tekstowe w kształcie (drop treści
        // = utrata danych). Pierwsza gałąź z txbxContent wygrywa (Choice i Fallback niosą TĘ SAMĄ
        // treść, więc bierzemy jedną — bez duplikacji).
        foreach (var branch in alternate.ChildElements)
        {
            if (branch is not (AlternateContentChoice or AlternateContentFallback)) continue;
            var textBox = RenderTextBoxContent(branch, document, sourcePart);
            if (!string.IsNullOrEmpty(textBox)) return textBox;
        }

        _log.LogDebug("mc:AlternateContent bez konwertowalnego obrazu ani pola tekstowego — element pominięty.");
        return string.Empty;
    }

    /// <summary>
    /// Renderuje treść pola tekstowego Worda (<c>w:txbxContent</c>, wspólne dla nowoczesnych
    /// kształtów <c>wps:txbx</c> i legacy VML <c>v:textbox</c>) jako blok. Przybliżenie KR-06:
    /// pozycja/obramowanie kształtu nie są w pełni odwzorowane, ale TREŚĆ tekstowa jest widoczna
    /// i edytowalna zamiast być cicho tracona. Zwraca pusty string, gdy brak txbxContent.
    /// </summary>
    private string RenderTextBoxContent(OpenXmlElement container, WordprocessingDocument document, OpenXmlPart? sourcePart)
    {
        var txbx = container.Descendants<TextBoxContent>().FirstOrDefault();
        if (txbx == null) return string.Empty;

        var inner = new StringBuilder();
        foreach (var child in txbx.Elements())
        {
            switch (child)
            {
                case Paragraph para:
                    inner.Append(ConvertParagraphToHtml(para, document, sourcePart));
                    break;
                case Table table:
                    inner.Append(ConvertTableToHtml(table, document, sourcePart));
                    break;
            }
        }
        if (inner.Length == 0) return string.Empty;

        // Geometria (rozmiar + pozycja). Kotwiczony (wp:anchor) text box dostaje pozycję
        // ABSOLUTNĄ z offsetów Worda (jak w MS Word), inline zostaje w przepływie. Pozycjonowanie
        // jest inline-CSS (reader-side), więc działa też w podglądzie stopki/nagłówka (JS edytora
        // nie musi go odtwarzać).
        var layout = BuildTextBoxLayoutCss(container);
        return $"<div class=\"docx-textbox\" data-textbox=\"1\" style=\"{layout}"
             + "border:1px solid #ccc;padding:4px 6px;box-sizing:border-box;\">"
             + inner + "</div>";
    }

    /// <summary>
    /// Renderuje wektorowy kształt DrawingML bez obrazu/tekstu (linia lub prostokąt) jako
    /// przybliżenie HTML: preset line/straightConnector → pozioma linia (border-top z grubości/
    /// koloru <c>a:ln</c>); prostokąt z <c>a:solidFill</c> → kolorowy blok. Kotwica → pozycja
    /// absolutna (jak w Wordzie). Zwraca pusty string dla nieobsługiwanej geometrii (drop bez zmian).
    /// </summary>
    private static string RenderVectorShapeAsHtml(Drawing drawing)
    {
        var preset = drawing.Descendants<DocumentFormat.OpenXml.Drawing.PresetGeometry>()
            .FirstOrDefault()?.Preset?.Value;

        var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        var widthPx = extent?.Cx != null ? (int)OoxmlUnits.EmuToPixels(extent.Cx.Value) : 0;
        var heightPx = extent?.Cy != null ? (int)OoxmlUnits.EmuToPixels(extent.Cy.Value) : 0;

        var outline = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
        var lineColor = HexColorOrNull(outline?.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()?.Val?.Value)
                        ?? "000000";
        var lineWidthPx = outline?.Width != null && outline.Width.Value > 0
            ? Math.Max(1, (int)Math.Round(OoxmlUnits.EmuToPixels(outline.Width.Value)))
            : 1;

        var isLine = preset == DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line
                     || preset == DocumentFormat.OpenXml.Drawing.ShapeTypeValues.StraightConnector1;

        var pos = BuildTextBoxLayoutCss(drawing); // reużycie: rozmiar + ewentualna pozycja absolutna
        // BuildTextBoxLayoutCss dokłada width/min-height; dla linii chcemy własną wysokość/tło.

        if (isLine)
        {
            // Pozioma linia: wysokość = grubość, tło = kolor; szerokość z extentu (fallback 100%).
            var w = widthPx > 0 ? $"{widthPx}px" : "100%";
            return $"<div class=\"docx-shape docx-line\" data-shape=\"line\" "
                 + $"style=\"{StripSize(pos)}width:{w};height:{lineWidthPx}px;"
                 + $"background:#{lineColor};margin:2px 0;\"></div>";
        }

        // Prostokąt z wypełnieniem — potrzebny widoczny rozmiar i tło.
        if (preset == DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle && widthPx > 0 && heightPx > 0)
        {
            var fill = HexColorOrNull(drawing.Descendants<DocumentFormat.OpenXml.Drawing.SolidFill>()
                .FirstOrDefault()?.RgbColorModelHex?.Val?.Value);
            var bg = fill != null ? $"background:#{fill};" : string.Empty;
            var border = $"border:{lineWidthPx}px solid #{lineColor};";
            return $"<div class=\"docx-shape docx-rect\" data-shape=\"rect\" "
                 + $"style=\"{pos}{bg}{border}box-sizing:border-box;\"></div>";
        }

        return string.Empty;
    }

    private static string? HexColorOrNull(string? value)
        => !string.IsNullOrEmpty(value) && System.Text.RegularExpressions.Regex.IsMatch(value, "^[0-9A-Fa-f]{6}$")
            ? value : null;

    /// <summary>Usuwa deklaracje width/min-height z gotowego CSS geometrii (linia ma własne).</summary>
    private static string StripSize(string css)
        => System.Text.RegularExpressions.Regex.Replace(css, @"(?:min-height|width):[^;]+;", string.Empty);

    /// <summary>
    /// CSS geometrii pola tekstowego: rozmiar z <c>wp:extent</c>/VML, a dla kotwiczonego
    /// <c>wp:anchor</c> — pozycja absolutna z offsetów (EMU→px) względem najbliższego
    /// pozycjonowanego przodka (strona / pasmo nagłówka-stopki). Inline → blok w przepływie.
    /// Przybliżenie: <c>relativeFrom</c> (page/margin/column/paragraph) nie jest w pełni
    /// rozróżniane — offset stosowany bezpośrednio (najczęstszy przypadek page/margin).
    /// </summary>
    private static string BuildTextBoxLayoutCss(OpenXmlElement container)
    {
        var extent = container.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        var widthPx = extent?.Cx != null ? (int)OoxmlUnits.EmuToPixels(extent.Cx.Value) : 0;
        var heightPx = extent?.Cy != null ? (int)OoxmlUnits.EmuToPixels(extent.Cy.Value) : 0;

        var sizeCss = new StringBuilder();
        if (widthPx > 0) sizeCss.Append($"width:{widthPx}px;");
        if (heightPx > 0) sizeCss.Append($"min-height:{heightPx}px;");

        var anchor = container.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
        if (anchor == null)
            return "display:inline-block;max-width:100%;vertical-align:top;margin:4px 0;" + sizeCss;

        long ReadOffset(OpenXmlElement? pos) =>
            pos?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset>()?.Text is string s
            && long.TryParse(s, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;

        var leftPx = (int)OoxmlUnits.EmuToPixels(
            ReadOffset(anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition>()));
        var topPx = (int)OoxmlUnits.EmuToPixels(
            ReadOffset(anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition>()));
        var zIndex = anchor.BehindDoc?.Value == true ? "z-index:0;" : "z-index:1;";

        return $"position:absolute;left:{leftPx}px;top:{topPx}px;{zIndex}" + sizeCss;
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
        if (blip?.Embed?.Value == null)
        {
            // Kształt bez obrazu (wps:wsp) może nieść POLE TEKSTOWE (wps:txbx → w:txbxContent).
            // Bez tego jego treść znikała bez śladu, a autosave tracił ją na stałe (KR-06).
            var textBox = RenderTextBoxContent(drawing, document, sourcePart);
            if (!string.IsNullOrEmpty(textBox)) return textBox;

            // Kształt wektorowy bez obrazu i tekstu (linia/prostokąt) — częsty w stopkach jako
            // separator/ramka. Wcześniej dropowany (brak blipa) → niewidoczny. Renderujemy
            // przybliżenie: linia = border, prostokąt z wypełnieniem = kolorowy blok.
            var shape = RenderVectorShapeAsHtml(drawing);
            if (!string.IsNullOrEmpty(shape)) return shape;

            // r:link = obraz linkowany (plik poza pakietem DOCX) — nie mamy jego bajtów i nie
            // pobieramy zewnętrznych URL-i po stronie serwera (SSRF). Kontrolowane pominięcie z logiem.
            if (blip?.Link?.Value != null)
                _log.LogWarning("Pominięto obraz z relacją zewnętrzną r:link={RelId} (obrazy linkowane nie są osadzone w pakiecie).",
                    blip.Link.Value);
            return string.Empty;
        }

        var relationshipId = blip.Embed.Value;

        // Relacje rozwiązujemy względem części, w której siedzi w:drawing (body/nagłówek/stopka) —
        // rId z nagłówka NIE wolno szukać w relacjach body (kolizje numeracji rId między częściami).
        var effectivePart = sourcePart ?? (OpenXmlPart?)document.MainDocumentPart;
        if (effectivePart == null) return string.Empty;

        var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        var width = extent?.Cx != null ? EmuToPx(extent.Cx.Value) : 200;
        var height = extent?.Cy != null ? EmuToPx(extent.Cy.Value) : 200;
        var widthEmu = extent?.Cx?.Value ?? OoxmlUnits.PixelsToEmu(width);
        var heightEmu = extent?.Cy?.Value ?? OoxmlUnits.PixelsToEmu(height);

        string? base64Data = null;
        string? contentType = null;

        if (_images.TryGetValue(ImageCacheKey(effectivePart, relationshipId), out var image))
        {
            base64Data = image.Base64Data;
            contentType = image.ContentType;
        }
        else
        {
            try
            {
                var imagePart = effectivePart.GetPartById(relationshipId) as ImagePart;
                if (imagePart != null)
                {
                    LoadImageFromPart(effectivePart, imagePart);
                    if (_images.TryGetValue(ImageCacheKey(effectivePart, relationshipId), out var lazy))
                    {
                        base64Data = lazy.Base64Data;
                        contentType = lazy.ContentType;
                    }
                }
                else
                {
                    _log.LogWarning("Relacja obrazu {RelId} w części {PartUri} nie wskazuje na ImagePart — obraz pominięty.",
                        relationshipId, effectivePart.Uri);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning("Nie udało się rozwiązać relacji obrazu {RelId} w części {PartUri}: {Error}",
                    relationshipId, effectivePart.Uri, ex.Message);
            }
        }

        if (base64Data == null || contentType == null) return string.Empty;

        // Zerowy/nieprawidłowy wp:extent (cx/cy = 0 u niektórych generatorów) dawał width:0px —
        // obraz istniał w DOM, ale był niewidoczny. Bierzemy wtedy wymiary intrinsic z nagłówka pliku.
        if (width <= 0 || height <= 0)
        {
            var probe = _graphics.ConvertForEditor(new GraphicSource
            {
                Data = System.Convert.FromBase64String(base64Data),
                ContentType = contentType,
                Origin = GraphicOrigin.LegacyDocxPart
            });
            width = probe.Web is { WidthPx: > 0 } pw ? pw.WidthPx : 200;
            height = probe.Web is { HeightPx: > 0 } ph ? ph.HeightPx : 200;
            widthEmu = OoxmlUnits.PixelsToEmu(width);
            heightEmu = OoxmlUnits.PixelsToEmu(height);
        }

        // Legacy metafile (EMF/WMF) → renderowalny data:URL (osadzony raster / placeholder SVG);
        // dla web-native zostaje oryginalny data:URL. Oryginalny part nietknięty (pass-through).
        var legacySrc = WebGraphicForLegacy(System.Convert.FromBase64String(base64Data), contentType, widthEmu, heightEmu);
        var drawingSrc = legacySrc?.dataUrl ?? $"data:{contentType};base64,{base64Data}";

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

        var legacyAttr = legacySrc?.isBlank == true ? " data-legacy-graphic=\"blank\"" : string.Empty;
        // Dla KAŻDEGO legacy metafile (EMF/WMF) — niezależnie czy w `src` jest przezroczysty blank SVG
        // czy zrasteryzowany PNG — niesiemy ORYGINALNY metafile w `data-original-src`. Dzięki temu writer
        // zapisuje do DOCX prawdziwy wektorowy EMF/WMF (Word renderuje natywnie), a PNG/blank
        // służy tylko do podglądu w przeglądarce. `legacySrc != null` ⇔ part był EMF/WMF/TIFF.
        var originalAttr = legacySrc != null
            ? $" data-original-src=\"data:{contentType};base64,{base64Data}\""
            : string.Empty;
        return $"<img src=\"{drawingSrc}\" " +
               $"style=\"max-width:100%;width:{width}px;height:{height}px;\" " +
               $"data-image-id=\"{relationshipId}\" " +
               $"data-width-emu=\"{widthEmu}\" data-height-emu=\"{heightEmu}\"" +
               $"{altAttr}{posAttrs}{borderAttrs}{cropAttrs}{legacyAttr}{originalAttr} />";
    }

    /// <summary>
    /// Konwertuje Picture (stary format VML) na HTML z odczytem wymiarów
    /// </summary>
    private string ConvertPictureToHtml(Picture picture, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var imageData = picture.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
        if (imageData?.RelationshipId?.Value == null)
        {
            // Legacy VML pole tekstowe (v:textbox → w:txbxContent) bez obrazu — zachowaj treść.
            return RenderTextBoxContent(picture, document, sourcePart);
        }

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

        // Jak w ConvertDrawingToHtml: relacje per część (kolizje rId między body a nagłówkiem/stopką).
        var effectivePart = sourcePart ?? (OpenXmlPart?)document.MainDocumentPart;
        if (effectivePart == null) return string.Empty;

        if (_images.TryGetValue(ImageCacheKey(effectivePart, relationshipId), out var image))
        {
            base64Data = image.Base64Data;
            contentType = image.ContentType;
        }
        else
        {
            try
            {
                if (effectivePart.GetPartById(relationshipId) is ImagePart imagePart)
                {
                    LoadImageFromPart(effectivePart, imagePart);
                    if (_images.TryGetValue(ImageCacheKey(effectivePart, relationshipId), out var lazy))
                    {
                        base64Data = lazy.Base64Data;
                        contentType = lazy.ContentType;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning("Nie udało się rozwiązać relacji VML v:imagedata {RelId} w części {PartUri}: {Error}",
                    relationshipId, effectivePart.Uri, ex.Message);
            }
        }

        if (base64Data == null || contentType == null) return string.Empty;

        // VML v:imagedata często wskazuje na EMF/WMF → renderowalny data:URL zamiast x-emf.
        var legacyVml = WebGraphicForLegacy(
            System.Convert.FromBase64String(base64Data), contentType,
            (long)(vmlWidth * 9525.0), (long)(vmlHeight * 9525.0));
        var vmlSrc = legacyVml?.dataUrl ?? $"data:{contentType};base64,{base64Data}";
        var vmlLegacyAttr = legacyVml?.isBlank == true ? " data-legacy-graphic=\"blank\"" : string.Empty;
        // Jak wyżej: oryginalny EMF/WMF do round-tripu zapisu zawsze, gdy part był metafile.
        var vmlOriginalAttr = legacyVml != null
            ? $" data-original-src=\"data:{contentType};base64,{base64Data}\""
            : string.Empty;

        return $"<img src=\"{vmlSrc}\" " +
               $"style=\"max-width:100%;width:{vmlWidth}px;height:{vmlHeight}px;\" " +
               $"data-image-id=\"{relationshipId}\"{vmlLegacyAttr}{vmlOriginalAttr} />";
    }

    /// <summary>
    /// Konwertuje tabelę na HTML z dokładnym odwzorowaniem obramowań, paddingu i stylów
    /// </summary>
    private string ConvertTableToHtml(Table table, WordprocessingDocument document, OpenXmlPart? sourcePart = null)
    {
        var html = new StringBuilder();
        var tableProps = table.GetFirstChild<TableProperties>();

        // Rozwiąż styl tabeli (w:tblStyle → łańcuch basedOn → tblLook → tblStylePr).
        // Większość tabel Worda (np. „Tabela – Siatka") ma obramowania/cieniowanie w STYLU,
        // nie w bezpośrednim tblPr — bez tego kroku renderowały się jako tabele bez linii.
        var styleCtx = ResolveTableStyleContext(tableProps);

        // Szerokość tabeli
        var tableWidth = "auto";
        var hasExplicitWidth = false;
        if (tableProps?.TableWidth?.Width?.Value != null)
        {
            var w = tableProps.TableWidth;
            if (w.Type?.Value == TableWidthUnitValues.Pct
                && double.TryParse(w.Width.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var pct50) && pct50 > 0)
            {
                // w:tblW pct = 1/50 procenta — zachowaj ułamek (3333 → 66.66%, nie 66%).
                tableWidth = string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.##}%", pct50 / 50.0);
                hasExplicitWidth = true;
            }
            else if (w.Type?.Value == TableWidthUnitValues.Dxa && int.TryParse(w.Width.Value, out var wtw) && wtw > 0)
            {
                tableWidth = $"{TwipsToPx(wtw)}px";
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

        // Odstęp między komórkami (w:tblCellSpacing) — Word renderuje wtedy rozdzielone
        // ramki komórek; w CSS odpowiada temu border-collapse:separate + border-spacing.
        var collapseCss = "border-collapse:collapse;";
        var cellSpacingAttr = string.Empty;
        var cellSpacingTw = GetTwipsValue(tableProps?.GetFirstChild<TableCellSpacing>());
        if (cellSpacingTw is > 0)
        {
            collapseCss = $"border-collapse:separate;border-spacing:{TwipsToPx(cellSpacingTw.Value)}px;";
            cellSpacingAttr = $" data-cell-spacing-tw=\"{cellSpacingTw.Value}\"";
        }

        // Sygnał dla HtmlToDocx: czy tabela ma jakiekolwiek zdefiniowane obramowania —
        // teraz liczone z EFEKTYWNYCH borderów (bezpośrednie tblBorders LUB styl tabeli).
        var tblBordersMarker = styleCtx.Borders.IsEmpty ? " data-no-borders=\"1\"" : "";

        // Referencja stylu tabeli — zachowywana w data-*, by eksport mógł ponownie
        // wyemitować w:tblStyle/w:tblLook (rozwiązane wartości i tak są w inline CSS).
        var styleAttrs = string.Empty;
        if (!string.IsNullOrEmpty(styleCtx.StyleId))
            styleAttrs = $" data-tbl-style=\"{System.Net.WebUtility.HtmlEncode(styleCtx.StyleId)}\" data-tbl-look=\"{styleCtx.LookHex}\"";

        // Domyślny padding komórek: bezpośredni tblCellMar → tblCellMar ze stylu tabeli →
        // domyślne marginesy Worda (TableNormal): top/bottom = 0, left/right = 108 twips.
        var defaultPadding = styleCtx.DefaultCellPaddingCss;

        var rows = table.Elements<TableRow>().ToList();
        var renderCtx = new TableRenderContext(
            styleCtx,
            defaultPadding,
            rows.Count,
            CountGridColumns(table, rows));

        html.Append($"<table{tblBordersMarker}{styleAttrs}{cellSpacingAttr} style=\"{collapseCss}width:{tableWidth};margin:4px 0;{layoutCss}{tableAlign}{tableIndent}\">");
        html.Append(colgroupHtml);

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            var trPr = row.TableRowProperties;

            // Wysokość wiersza. Na <tr> działa wyłącznie `height` (w tabelach zachowuje się
            // jak min-height); `min-height` na tr jest przez przeglądarki IGNOROWANE, więc
            // wiersze atLeast traciły wysokość. Oryginalne twips + reguła idą w data-*,
            // żeby eksport nie tracił hRule ani precyzji px→twips.
            var rowStyle = "";
            var rowAttrs = new StringBuilder();
            var trHeight = trPr?.Elements<TableRowHeight>().FirstOrDefault();
            if (trHeight?.Val?.Value != null)
            {
                var hRule = trHeight.HeightType?.Value ?? HeightRuleValues.AtLeast;
                if (hRule != HeightRuleValues.Auto)
                {
                    var hPx = TwipsToPx((int)trHeight.Val.Value);
                    rowStyle = $" style=\"height:{hPx}px;\"";
                    rowAttrs.Append($" data-row-height-tw=\"{trHeight.Val.Value}\"");
                    if (hRule == HeightRuleValues.Exact)
                        rowAttrs.Append(" data-row-hrule=\"exact\"");
                }
            }

            // Wiersz nagłówkowy powtarzany na stronach + zakaz dzielenia wiersza —
            // brak odpowiednika w edytorze (nie paginuje jak Word), ale round-trip
            // przez data-* chroni właściwości przy eksporcie.
            if (trPr?.Elements<TableHeader>().Any() == true)
                rowAttrs.Append(" data-tbl-header=\"1\"");
            if (trPr?.Elements<CantSplit>().Any() == true)
                rowAttrs.Append(" data-cant-split=\"1\"");

            html.Append($"<tr{rowAttrs}{rowStyle}>");

            // Iteruj komórki uwzględniając komórki opakowane w SDT (Content Control / formant).
            // SdtCell zawiera SdtContentCell, a w nim faktyczne TableCell — inaczej znikają dane.
            // gridCursor śledzi pozycję komórki w siatce (gridSpan przesuwa kursor; komórki
            // kontynuacji vMerge też zajmują kolumny, mimo że nie emitują <td>).
            var rowCells = FlattenRowCells(row).ToList();

            // Short row: the sum of the row's gridSpans is smaller than the table grid. Word merges
            // the remainder into the last cell (a row that declares one cell for a fully-merged row
            // has gridSpan=1, not N). Without this, `table-layout:fixed`+colgroup pins that cell to
            // a single narrow column and the rest of the row renders as empty phantom columns —
            // exactly the "merged content squeezed into the first column" defect.
            var rowGridTotal = rowCells.Sum(GetGridSpan);
            var deficit = renderCtx.GridColumnCount - rowGridTotal;

            var gridCursor = 0;
            for (var ci = 0; ci < rowCells.Count; ci++)
            {
                var extraColspan = (deficit > 0 && ci == rowCells.Count - 1) ? deficit : 0;
                AppendTableCellHtml(html, table, rows, rowIndex, rowCells[ci], renderCtx,
                    ref gridCursor, extraColspan, document, sourcePart);
            }

            html.Append("</tr>");
        }

        html.Append("</table>");
        return html.ToString();
    }

    /// <summary>
    /// Liczba kolumn siatki tabeli: z w:tblGrid, a gdy brak — maksimum sumy gridSpan po wierszach.
    /// </summary>
    private static int CountGridColumns(Table table, List<TableRow> rows)
    {
        var grid = table.GetFirstChild<TableGrid>();
        var fromGrid = grid?.Elements<GridColumn>().Count() ?? 0;
        if (fromGrid > 0) return fromGrid;

        var max = 0;
        foreach (var row in rows)
        {
            var count = row.Elements<TableCell>().Sum(GetGridSpan);
            max = Math.Max(max, count);
        }
        return max;
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

    /// <summary>Komórki wiersza w kolejności dokumentu, z rozpakowaniem komórek w SDT.</summary>
    private static IEnumerable<TableCell> FlattenRowCells(TableRow row)
    {
        foreach (var cellLike in row.Elements())
        {
            if (cellLike is TableCell cell)
                yield return cell;
            else if (cellLike is SdtCell sdtCell
                     && sdtCell.GetFirstChild<SdtContentCell>() is { } sdtContent)
                foreach (var innerCell in sdtContent.Elements<TableCell>())
                    yield return innerCell;
        }
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
    /// Pobiera szczegółowy styl CSS komórki: obramowania (bezpośrednie → styl warunkowy →
    /// styl tabeli, pozycyjnie: krawędź zewnętrzna vs insideH/insideV), padding, tło
    /// (z warunkowym formatowaniem stylu: firstRow/lastRow/kolumny/pasy), wyrównania.
    /// </summary>
    private string GetTableCellStyleDetailed(
        TableCell cell,
        TableRenderContext ctx,
        int rowIndex,
        int gridColStart,
        int gridSpan,
        int rowSpan)
    {
        var css = new StringBuilder();
        var props = cell.TableCellProperties;

        // Regiony warunkowego formatowania stylu (najbardziej specyficzny pierwszy).
        var regions = ComputeConditionalRegions(ctx, rowIndex, gridColStart, gridSpan, rowSpan);

        // Pozycja komórki w siatce decyduje, która strona tabeli jest jej „domyślną" ramką:
        // krawędzie zewnętrzne biorą top/bottom/left/right, wewnętrzne — insideH/insideV.
        var isFirstRow = rowIndex == 0;
        var isLastRow = rowIndex + rowSpan >= ctx.RowCount;
        var isFirstCol = gridColStart == 0;
        var isLastCol = ctx.GridColumnCount <= 0 || gridColStart + gridSpan >= ctx.GridColumnCount;

        var cb = props?.TableCellBorders;
        css.Append($"border-top:{ResolveCellBorderSide(cb?.TopBorder, regions, TableCellEdge.Top, isFirstRow ? ctx.Style.Borders.Top : ctx.Style.Borders.InsideH, ctx)};");
        css.Append($"border-bottom:{ResolveCellBorderSide(cb?.BottomBorder, regions, TableCellEdge.Bottom, isLastRow ? ctx.Style.Borders.Bottom : ctx.Style.Borders.InsideH, ctx)};");
        css.Append($"border-left:{ResolveCellBorderSide(cb?.LeftBorder, regions, TableCellEdge.Left, isFirstCol ? ctx.Style.Borders.Left : ctx.Style.Borders.InsideV, ctx)};");
        css.Append($"border-right:{ResolveCellBorderSide(cb?.RightBorder, regions, TableCellEdge.Right, isLastCol ? ctx.Style.Borders.Right : ctx.Style.Borders.InsideV, ctx)};");

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
            css.Append($"padding:{ctx.DefaultPadding};");
        }

        // Wyrównanie pionowe komórki (top/middle/bottom). Emitowane raz z rozwiązaną wartością
        // (domyślnie top jak w Wordzie) — wcześniej „top" szło bezwarunkowo, a rzeczywista wartość
        // dopisywana była drugi raz niżej, zostawiając zduplikowaną deklarację w inline style.
        var vAlign = props?.TableCellVerticalAlignment?.Val != null
            ? GetTableVerticalAlignment(props.TableCellVerticalAlignment.Val.Value)
            : "top";
        css.Append($"vertical-align:{vAlign};");

        // Tło: bezpośrednie tcPr → regiony stylu warunkowego → tcPr stylu (cała tabela) →
        // tblPr shd (bezpośrednie lub ze stylu). Rozwiązuje themeFill/tint/shade i wzory pct.
        var shadingHex = ResolveShadingHex(props?.Shading);
        if (shadingHex == null)
        {
            foreach (var region in regions)
            {
                shadingHex = ResolveShadingHex(region.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.GetFirstChild<Shading>());
                if (shadingHex != null) break;
            }
        }
        shadingHex ??= ResolveShadingHex(ctx.Style.WholeTableCellShading);
        shadingHex ??= ResolveShadingHex(ctx.Style.TableShading);
        if (shadingHex != null)
            css.Append($"background-color:#{shadingHex};");

        if (props != null)
        {
            // Szerokość — tylko realne jednostki. w:tcW type=auto/nil ma zwykle w:w="0",
            // co dawało width:0px i łamało układ.
            var w = props.TableCellWidth;
            if (w?.Width?.Value != null)
            {
                if (w.Type?.Value == TableWidthUnitValues.Pct
                    && double.TryParse(w.Width.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var pct50) && pct50 > 0)
                    css.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, "width:{0:0.##}%;", pct50 / 50.0));
                else if ((w.Type == null || w.Type.Value == TableWidthUnitValues.Dxa)
                    && int.TryParse(w.Width.Value, out var wtw) && wtw > 0)
                    css.Append($"width:{TwipsToPx(wtw)}px;");
            }

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

    private enum TableCellEdge { Top, Bottom, Left, Right }

    /// <summary>
    /// Rozstrzyga jedną krawędź komórki: bezpośredni tcBorders (w tym jawne none/nil) →
    /// tcBorders/tblBorders regionów stylu warunkowego → tcPr stylu (cała tabela) →
    /// pozycyjna krawędź z efektywnych tblBorders (przekazana jako fallback).
    /// </summary>
    private string ResolveCellBorderSide(
        BorderType? directBorder,
        List<TableStyleProperties> regions,
        TableCellEdge edge,
        BorderType? tableFallback,
        TableRenderContext ctx)
    {
        // Bezpośrednia definicja na komórce wygrywa zawsze — także jawne "brak linii".
        if (directBorder != null)
        {
            var v = directBorder.Val?.Value;
            if (v == null || v == BorderValues.None || v == BorderValues.Nil) return "none";
            return GetBorderCss(directBorder);
        }

        foreach (var region in regions)
        {
            var tcB = region.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>()?.GetFirstChild<TableCellBorders>();
            var b = PickEdge(tcB, edge);
            if (b == null)
            {
                var tblB = region.GetFirstChild<TableStyleConditionalFormattingTableProperties>()?.GetFirstChild<TableBorders>();
                b = PickEdge(tblB, edge);
            }
            if (b != null)
            {
                var v = b.Val?.Value;
                if (v == null || v == BorderValues.None || v == BorderValues.Nil) return "none";
                return GetBorderCss(b);
            }
        }

        var whole = PickEdge(ctx.Style.WholeTableCellBorders, edge);
        if (whole != null)
        {
            var v = whole.Val?.Value;
            if (v == null || v == BorderValues.None || v == BorderValues.Nil) return "none";
            return GetBorderCss(whole);
        }

        if (tableFallback == null) return "none";
        var fv = tableFallback.Val?.Value;
        if (fv == null || fv == BorderValues.None || fv == BorderValues.Nil) return "none";
        return GetBorderCss(tableFallback);
    }

    private static BorderType? PickEdge(TableCellBorders? b, TableCellEdge edge) => edge switch
    {
        TableCellEdge.Top => b?.TopBorder,
        TableCellEdge.Bottom => b?.BottomBorder,
        TableCellEdge.Left => b?.LeftBorder,
        TableCellEdge.Right => b?.RightBorder,
        _ => null
    };

    private static BorderType? PickEdge(TableBorders? b, TableCellEdge edge) => edge switch
    {
        TableCellEdge.Top => b?.TopBorder,
        TableCellEdge.Bottom => b?.BottomBorder,
        TableCellEdge.Left => b?.LeftBorder,
        TableCellEdge.Right => b?.RightBorder,
        _ => null
    };

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

    #region Rozwiązywanie stylu tabeli (w:tblStyle / w:tblLook / w:tblStylePr)

    /// <summary>Efektywne obramowania tabeli: bezpośrednie tblBorders scalone per strona ze stylem tabeli.</summary>
    private sealed class EffectiveTableBorders
    {
        public BorderType? Top, Bottom, Left, Right, InsideH, InsideV;

        public bool IsEmpty
        {
            get
            {
                static bool Blank(BorderType? b)
                {
                    if (b == null) return true;
                    var v = b.Val?.Value;
                    return v == null || v == BorderValues.None || v == BorderValues.Nil;
                }
                return Blank(Top) && Blank(Bottom) && Blank(Left) && Blank(Right) && Blank(InsideH) && Blank(InsideV);
            }
        }
    }

    /// <summary>
    /// Rozwiązany kontekst stylu tabeli: łańcuch basedOn, flagi tblLook, efektywne obramowania,
    /// cieniowanie i marginesy komórek oraz formaty warunkowe (tblStylePr) per region.
    /// </summary>
    private sealed class TableStyleContext
    {
        public string? StyleId;
        public string LookHex = "04A0";
        public bool FirstRow, LastRow, FirstColumn, LastColumn, RowBands = true, ColumnBands;
        public int RowBandSize = 1, ColBandSize = 1;
        public EffectiveTableBorders Borders = new();
        public Shading? TableShading;            // w:tblPr/w:shd (bezpośrednie lub ze stylu)
        public Shading? WholeTableCellShading;   // w:tcPr/w:shd stylu — tło każdej komórki
        public TableCellBorders? WholeTableCellBorders; // w:tcPr/w:tcBorders stylu
        public string DefaultCellPaddingCss = "";
        public Dictionary<TableStyleOverrideValues, TableStyleProperties> Conditional = new();
    }

    /// <summary>Kontekst renderowania jednej tabeli (styl + geometria siatki).</summary>
    private sealed class TableRenderContext
    {
        public TableStyleContext Style { get; }
        public string DefaultPadding { get; }
        public int RowCount { get; }
        public int GridColumnCount { get; }

        public TableRenderContext(TableStyleContext style, string defaultPadding, int rowCount, int gridColumnCount)
        {
            Style = style;
            DefaultPadding = defaultPadding;
            RowCount = rowCount;
            GridColumnCount = gridColumnCount;
        }
    }

    private TableStyleContext ResolveTableStyleContext(TableProperties? tblPr)
    {
        var ctx = new TableStyleContext();

        // Łańcuch stylów tabeli (najbardziej pochodny pierwszy), z limitem na cykle.
        var chain = new List<Style>();
        var styleId = tblPr?.TableStyle?.Val?.Value;
        ctx.StyleId = styleId;
        var guard = 0;
        while (!string.IsNullOrEmpty(styleId) && guard++ < 12 && _rawStyles.TryGetValue(styleId!, out var st))
        {
            chain.Add(st);
            styleId = st.BasedOn?.Val?.Value;
        }

        // tblLook: atrybuty boolowskie mają pierwszeństwo; starszy zapis to maska hex w @w:val
        // (0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn, 0x0100 lastColumn,
        //  0x0200 noHBand, 0x0400 noVBand). Domyślne Worda: firstRow + firstColumn + noVBand.
        var look = tblPr?.GetFirstChild<TableLook>();
        int mask = 0x0020 | 0x0080 | 0x0400;
        if (look?.Val?.Value is string lookVal
            && int.TryParse(lookVal, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var parsedMask))
            mask = parsedMask;
        ctx.FirstRow = look?.FirstRow?.Value ?? (mask & 0x0020) != 0;
        ctx.LastRow = look?.LastRow?.Value ?? (mask & 0x0040) != 0;
        ctx.FirstColumn = look?.FirstColumn?.Value ?? (mask & 0x0080) != 0;
        ctx.LastColumn = look?.LastColumn?.Value ?? (mask & 0x0100) != 0;
        ctx.RowBands = !(look?.NoHorizontalBand?.Value ?? (mask & 0x0200) != 0);
        ctx.ColumnBands = !(look?.NoVerticalBand?.Value ?? (mask & 0x0400) != 0);
        ctx.LookHex = ((ctx.FirstRow ? 0x0020 : 0) | (ctx.LastRow ? 0x0040 : 0)
            | (ctx.FirstColumn ? 0x0080 : 0) | (ctx.LastColumn ? 0x0100 : 0)
            | (ctx.RowBands ? 0 : 0x0200) | (ctx.ColumnBands ? 0 : 0x0400))
            .ToString("X4", System.Globalization.CultureInfo.InvariantCulture);

        // Efektywne obramowania: per strona — bezpośrednie tblBorders wygrywa, potem
        // pierwsza definicja tej strony w łańcuchu stylów (dziedziczenie element-wise).
        BorderType? Side(Func<TableBorders, BorderType?> pick)
        {
            if (tblPr?.TableBorders is TableBorders direct && pick(direct) is BorderType d) return d;
            foreach (var st in chain)
            {
                var sb = st.StyleTableProperties?.GetFirstChild<TableBorders>();
                if (sb != null && pick(sb) is BorderType b) return b;
            }
            return null;
        }
        ctx.Borders.Top = Side(b => b.TopBorder);
        ctx.Borders.Bottom = Side(b => b.BottomBorder);
        ctx.Borders.Left = Side(b => b.LeftBorder);
        ctx.Borders.Right = Side(b => b.RightBorder);
        ctx.Borders.InsideH = Side(b => b.InsideHorizontalBorder);
        ctx.Borders.InsideV = Side(b => b.InsideVerticalBorder);

        // Cieniowanie całej tabeli / wszystkich komórek ze stylu.
        ctx.TableShading = tblPr?.GetFirstChild<Shading>()
            ?? chain.Select(s => s.StyleTableProperties?.GetFirstChild<Shading>()).FirstOrDefault(s => s != null);
        ctx.WholeTableCellShading = chain
            .Select(s => s.StyleTableCellProperties?.GetFirstChild<Shading>())
            .FirstOrDefault(s => s != null);
        ctx.WholeTableCellBorders = chain
            .Select(s => s.StyleTableCellProperties?.GetFirstChild<TableCellBorders>())
            .FirstOrDefault(b => b != null);

        // Rozmiar pasów (banding).
        ctx.RowBandSize = (int?)tblPr?.GetFirstChild<TableStyleRowBandSize>()?.Val?.Value
            ?? chain.Select(s => (int?)s.StyleTableProperties?.GetFirstChild<TableStyleRowBandSize>()?.Val?.Value)
                .FirstOrDefault(v => v != null) ?? 1;
        ctx.ColBandSize = (int?)tblPr?.GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value
            ?? chain.Select(s => (int?)s.StyleTableProperties?.GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value)
                .FirstOrDefault(v => v != null) ?? 1;

        // Formaty warunkowe: od bazy do najbardziej pochodnego, by pochodny nadpisał region.
        for (var i = chain.Count - 1; i >= 0; i--)
        {
            foreach (var tsp in chain[i].Elements<TableStyleProperties>())
            {
                if (tsp.Type?.Value is TableStyleOverrideValues t)
                    ctx.Conditional[t] = tsp;
            }
        }

        // Domyślne marginesy komórek (padding): bezpośredni tblCellMar → styl → default Worda.
        const int wordDefaultCellMarginTwips = 108; // 0.19 cm
        var cellMars = new List<TableCellMarginDefault>();
        if (tblPr?.TableCellMarginDefault != null) cellMars.Add(tblPr.TableCellMarginDefault);
        foreach (var st in chain)
        {
            var m = st.StyleTableProperties?.GetFirstChild<TableCellMarginDefault>();
            if (m != null) cellMars.Add(m);
        }
        int PadSide(Func<TableCellMarginDefault, int?> pick, int fallback)
        {
            foreach (var m in cellMars)
                if (pick(m) is int v) return v;
            return fallback;
        }
        var topPad = PadSide(m => GetTwipsValue(m.TopMargin), 0);
        var bottomPad = PadSide(m => GetTwipsValue(m.BottomMargin), 0);
        var leftPad = PadSide(m => GetDxaValue(m.TableCellLeftMargin), wordDefaultCellMarginTwips);
        var rightPad = PadSide(m => GetDxaValue(m.TableCellRightMargin), wordDefaultCellMarginTwips);
        ctx.DefaultCellPaddingCss = $"{TwipsToPx(topPad)}px {TwipsToPx(rightPad)}px {TwipsToPx(bottomPad)}px {TwipsToPx(leftPad)}px";

        return ctx;
    }

    /// <summary>
    /// Regiony formatowania warunkowego stylu obejmujące komórkę, w kolejności priorytetu
    /// (najbardziej specyficzny pierwszy): firstRow/lastRow → firstCol/lastCol → pasy.
    /// Pasy liczone z pominięciem wiersza nagłówkowego / pierwszej kolumny (jak w Wordzie).
    /// </summary>
    private static List<TableStyleProperties> ComputeConditionalRegions(
        TableRenderContext ctx, int rowIndex, int gridColStart, int gridSpan, int rowSpan)
    {
        var s = ctx.Style;
        var result = new List<TableStyleProperties>();
        if (s.Conditional.Count == 0) return result;

        var isFirstRow = s.FirstRow && rowIndex == 0;
        var isLastRow = s.LastRow && rowIndex + rowSpan >= ctx.RowCount;
        var isFirstCol = s.FirstColumn && gridColStart == 0;
        var isLastCol = s.LastColumn && ctx.GridColumnCount > 0 && gridColStart + gridSpan >= ctx.GridColumnCount;

        void Add(TableStyleOverrideValues t)
        {
            if (s.Conditional.TryGetValue(t, out var p)) result.Add(p);
        }

        if (isFirstRow) Add(TableStyleOverrideValues.FirstRow);
        if (isLastRow) Add(TableStyleOverrideValues.LastRow);
        if (isFirstCol) Add(TableStyleOverrideValues.FirstColumn);
        if (isLastCol) Add(TableStyleOverrideValues.LastColumn);

        if (s.ColumnBands && !isFirstCol && !isLastCol)
        {
            var colForBand = gridColStart - (s.FirstColumn ? 1 : 0);
            if (colForBand >= 0)
                Add((colForBand / Math.Max(1, s.ColBandSize)) % 2 == 0
                    ? TableStyleOverrideValues.Band1Vertical
                    : TableStyleOverrideValues.Band2Vertical);
        }
        if (s.RowBands && !isFirstRow && !isLastRow)
        {
            var rowForBand = rowIndex - (s.FirstRow ? 1 : 0);
            if (rowForBand >= 0)
                Add((rowForBand / Math.Max(1, s.RowBandSize)) % 2 == 0
                    ? TableStyleOverrideValues.Band1Horizontal
                    : TableStyleOverrideValues.Band2Horizontal);
        }
        return result;
    }

    /// <summary>
    /// Rozwiązuje w:shd na kolor hex (bez '#'): fill/themeFill(+tint/shade), a wzory pctNN
    /// przybliża mieszając kolor wzoru z tłem w zadanej proporcji (val=solid → kolor wzoru).
    /// Zwraca null dla braku/auto/clear-bez-fill.
    /// </summary>
    private string? ResolveShadingHex(Shading? shd)
    {
        if (shd == null) return null;

        string? fill = null;
        if (shd.Fill?.Value is string f && !string.Equals(f, "auto", StringComparison.OrdinalIgnoreCase))
            fill = f;
        else if (shd.ThemeFill?.HasValue == true)
        {
            var themeHex = ResolveThemeColor(shd.ThemeFill.Value)?.TrimStart('#');
            if (themeHex != null)
                fill = ApplyTintShade(themeHex, shd.ThemeFillTint?.Value, shd.ThemeFillShade?.Value);
        }

        var patternPct = GetShadingPatternPercent(shd.Val?.Value);
        if (patternPct is double pct && pct > 0)
        {
            string patternColor = "000000";
            if (shd.Color?.Value is string c && !string.Equals(c, "auto", StringComparison.OrdinalIgnoreCase))
                patternColor = c;
            else if (shd.ThemeColor?.HasValue == true)
            {
                var th = ResolveThemeColor(shd.ThemeColor.Value)?.TrimStart('#');
                if (th != null) patternColor = ApplyTintShade(th, shd.ThemeTint?.Value, shd.ThemeShade?.Value);
            }
            fill = BlendHex(patternColor, fill ?? "FFFFFF", pct / 100.0);
        }

        if (fill == null) return null;
        return Regex.IsMatch(fill, "^[0-9A-Fa-f]{6}$") ? fill.ToUpperInvariant() : null;
    }

    /// <summary>Procent wzoru cieniowania (pct5..pct95, solid=100). Null dla clear/braku.</summary>
    private static double? GetShadingPatternPercent(ShadingPatternValues? val)
    {
        if (val == null) return null;
        if (val == ShadingPatternValues.Solid) return 100;
        // Nazwa literału OOXML ma postać "pctNN" / "pctNNN" (pct12 = 12.5% itd.).
        var literal = ((DocumentFormat.OpenXml.IEnumValue)val.Value).Value;
        var m = Regex.Match(literal, @"^pct(\d+)$");
        if (!m.Success) return null;
        var n = double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        // pct12/37/62/87 to w OOXML 12.5/37.5/62.5/87.5.
        if (n == 12 || n == 37 || n == 62 || n == 87) n += 0.5;
        return n;
    }

    /// <summary>Aplikuje themeTint/themeShade (hex bajt) na kolor hex RRGGBB.</summary>
    private static string ApplyTintShade(string hex, string? tintHex, string? shadeHex)
    {
        if (!Regex.IsMatch(hex, "^[0-9A-Fa-f]{6}$")) return hex;
        double r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
        double g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
        double b = System.Convert.ToInt32(hex.Substring(4, 2), 16);

        if (tintHex != null && int.TryParse(tintHex, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var tint))
        {
            var t = tint / 255.0;
            r = r * t + 255 * (1 - t);
            g = g * t + 255 * (1 - t);
            b = b * t + 255 * (1 - t);
        }
        if (shadeHex != null && int.TryParse(shadeHex, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var shade))
        {
            var s = shade / 255.0;
            r *= s; g *= s; b *= s;
        }
        return $"{(int)Math.Round(r):X2}{(int)Math.Round(g):X2}{(int)Math.Round(b):X2}";
    }

    /// <summary>Miesza kolor fg z tłem bg w proporcji weight (0..1) kanał po kanale.</summary>
    private static string BlendHex(string fgHex, string bgHex, double weight)
    {
        if (!Regex.IsMatch(fgHex, "^[0-9A-Fa-f]{6}$") || !Regex.IsMatch(bgHex, "^[0-9A-Fa-f]{6}$"))
            return fgHex;
        weight = Math.Clamp(weight, 0, 1);
        int Mix(int i)
        {
            var fg = System.Convert.ToInt32(fgHex.Substring(i, 2), 16);
            var bg = System.Convert.ToInt32(bgHex.Substring(i, 2), 16);
            return (int)Math.Round(fg * weight + bg * (1 - weight));
        }
        return $"{Mix(0):X2}{Mix(2):X2}{Mix(4):X2}";
    }

    #endregion

    /// <summary>
    /// Atrybuty data-* formantu (SDT). Oprócz czytelnych `data-sdt-tag`/`data-sdt-alias`
    /// niesie PEŁNE właściwości `w:sdtPr` (base64 OuterXml) w `data-sdt-props`, żeby eksport
    /// odtworzył typ formantu i jego właściwości (checkbox/dropDownList/comboBox/date/text/
    /// richText/picture, opcje listy, format daty, lock, id, placeholder, databinding) zamiast
    /// degradować formant do generycznego. Treść formantu jest edytowalna osobno.
    /// </summary>
    private static string BuildSdtDataAttrs(SdtProperties? props)
    {
        if (props == null) return string.Empty;
        var sb = new StringBuilder();
        var tag = props.Elements<Tag>().FirstOrDefault()?.Val?.Value ?? "";
        var alias = props.Elements<SdtAlias>().FirstOrDefault()?.Val?.Value ?? "";
        if (!string.IsNullOrEmpty(tag)) sb.Append($" data-sdt-tag=\"{System.Net.WebUtility.HtmlEncode(tag)}\"");
        if (!string.IsNullOrEmpty(alias)) sb.Append($" data-sdt-alias=\"{System.Net.WebUtility.HtmlEncode(alias)}\"");

        // Pełne właściwości jako base64 (bez problemów z cudzysłowami/escapowaniem XML w atrybucie).
        var xml = props.OuterXml;
        if (!string.IsNullOrEmpty(xml))
        {
            var b64 = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(xml));
            sb.Append($" data-sdt-props=\"{b64}\"");
        }
        return sb.ToString();
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
            html.Append($"<div class=\"sdt-block\"{BuildSdtDataAttrs(props)}>");

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
        var dataAttrs = BuildSdtDataAttrs(props);

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
        List<TableRow> rows,
        int rowIndex,
        TableCell cell,
        TableRenderContext ctx,
        ref int gridCursor,
        int extraColspan,
        WordprocessingDocument document,
        OpenXmlPart? sourcePart)
    {
        var cellProps = cell.TableCellProperties;
        // extraColspan absorbs the columns a short row leaves unfilled (see ConvertTableToHtml).
        var gridSpan = GetGridSpan(cell) + Math.Max(0, extraColspan);
        var gridColStart = gridCursor;
        gridCursor += gridSpan; // komórka (także kontynuacja vMerge) zajmuje kolumny siatki

        var colspan = "";
        if (gridSpan > 1)
            colspan = $" colspan=\"{gridSpan}\"";

        var rowspan = "";
        var rowSpanCount = 1;
        var row = rows[rowIndex];
        var vMerge = cellProps?.VerticalMerge;
        if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
        {
            rowSpanCount = CountRowSpan(table, row, cell);
            if (rowSpanCount > 1) rowspan = $" rowspan=\"{rowSpanCount}\"";
        }
        else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
        {
            // An omitted vMerge val defaults to "continue" (ECMA-376) — drop the merged cell.
            return;
        }

        var cellStyle = GetTableCellStyleDetailed(cell, ctx, rowIndex, gridColStart, gridSpan, rowSpanCount);

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

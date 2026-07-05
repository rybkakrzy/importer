using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Conversion;
using HtmlAgilityPack;
using OoxmlPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
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
    // HTML data-num-id (tożsamość listy z readera) → numId w generowanym pakiecie. Fragmenty listy
    // rozdzielone akapitem, ale należące do tej samej listy logicznej Worda, współdzielą jedną
    // NumberingInstance — Word kontynuuje wtedy numerację. Różne data-num-id → osobne instancje.
    private readonly Dictionary<string, int> _numIdByHtmlList = new();

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
    /// Geometria jednej sekcji strony (cm). Wartości null = dziedziczone/domyślne.
    /// <c>BreakType</c> opisuje, jak sekcja się ZACZYNA (w:sectPr/w:type następnej sekcji).
    /// </summary>
    private sealed class SectionGeometry
    {
        public Domain.Models.PageSize? PageSize { get; set; }
        public PageMargins? Margins { get; set; }
        public double? HeaderDistanceCm { get; set; }
        public double? FooterDistanceCm { get; set; }
        public string? BreakType { get; set; }
    }

    /// <summary>
    /// Geometria AKTUALNIE otwartej sekcji podczas konwersji body. Start = argumenty
    /// Convert (pierwsza sekcja); każdy marker <c>div.docx-section-break</c> zamyka
    /// bieżącą sekcję (w:p/pPr/sectPr z tą geometrią) i otwiera następną z data-*.
    /// Na końcu trzyma geometrię OSTATNIEJ sekcji → trafia do body-level sectPr.
    /// </summary>
    private SectionGeometry _currentSection = new();

    /// <summary>Pierwszy sectPr w kolejności dokumentu — tu wpinamy referencje
    /// nagłówka/stopki i titlePg (kolejne sekcje dziedziczą je w Wordzie).</summary>
    private SectionProperties? _firstSectionProps;

    /// <summary>Paragraph-level sectPr w kolejności emisji: element [k] zamyka sekcję o
    /// indeksie k (0-based) — cel dla własnych nagłówków/stopek sekcji (SectionHeaderFooter).</summary>
    private readonly List<SectionProperties> _emittedSectionProps = new();

    /// <summary>Czy body zawierało markery sekcji (dokument wielosekcyjny).</summary>
    private bool _hasSectionMarkers;

    /// <summary>Wysokości pasm nagłówka/stopki (cm) do wyliczenia w:header/w:footer
    /// distance dla sectPr sekcji pośrednich (te same, co dla body-level sectPr).</summary>
    private double? _headerBandCm;
    private double? _footerBandCm;

    /// <summary>
    /// Konwertuje HTML na plik DOCX
    /// </summary>
    public byte[] Convert(string html, DocumentMetadata? metadata = null, HeaderFooterContent? header = null, HeaderFooterContent? footer = null, PageMargins? margins = null, Domain.Models.PageSize? pageSize = null, IReadOnlyList<SectionHeaderFooter>? sectionHeadersFooters = null)
    {
        using var memoryStream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
        {
            _mainPart = document.AddMainDocumentPart();
            _mainPart.Document = new Document();

            // Stan sekcji per konwersja: pierwsza sekcja = geometria z argumentów;
            // markery w HTML nadpisują ją dla kolejnych sekcji.
            _currentSection = new SectionGeometry { PageSize = pageSize, Margins = margins };
            _firstSectionProps = null;
            _emittedSectionProps.Clear();
            _numIdByHtmlList.Clear();
            _hasSectionMarkers = false;
            _headerBandCm = header?.Height;
            _footerBandCm = footer?.Height;

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
            AddPageSettings(body, header, footer, margins, pageSize);

            // Własne nagłówki/stopki sekcji ≥ 1 (po AddPageSettings — body-level sectPr istnieje)
            AddSectionHeadersFooters(document, sectionHeadersFooters);

            document.Save();
        }

        return memoryStream.ToArray();
    }

    public byte[] ConvertPreservingPackage(string html, Stream? originalPackage,
        DocumentMetadata? metadata = null, HeaderFooterContent? header = null,
        HeaderFooterContent? footer = null, PageMargins? margins = null, Domain.Models.PageSize? pageSize = null,
        IReadOnlyList<SectionHeaderFooter>? sectionHeadersFooters = null)
    {
        var generated = Convert(html, metadata, header, footer, margins, pageSize, sectionHeadersFooters);

        if (originalPackage == null || !originalPackage.CanRead)
            return generated;

        try
        {
            return PreserveOriginalParts(generated, originalPackage);
        }
        catch
        {
            // Pass-through is best-effort: a malformed / unexpected original must never break the
            // save — fall back to the fully self-contained generated document.
            return generated;
        }
    }

    /// <summary>
    /// Replaces the generated package's styles.xml, theme and fontTable with the originals so the
    /// full style set (incl. table styles) and the document/theme fonts survive the round-trip.
    /// Body, sectPr, headers/footers, images and numbering remain as generated (they carry the
    /// editor's actual changes and standard style IDs that exist in the original styles.xml).
    /// </summary>
    private static byte[] PreserveOriginalParts(byte[] generated, Stream originalPackage)
    {
        var ms = new MemoryStream();
        ms.Write(generated, 0, generated.Length);
        ms.Position = 0;

        if (originalPackage.CanSeek) originalPackage.Position = 0;

        using (var original = WordprocessingDocument.Open(originalPackage, false))
        using (var target = WordprocessingDocument.Open(ms, true))
        {
            var origMain = original.MainDocumentPart;
            var targetMain = target.MainDocumentPart;
            if (origMain == null || targetMain == null)
                return generated;

            // styles.xml — pełny zestaw stylów (w tym ~100 stylów tabel) + docDefaults (font/theme).
            // FeedData do ISTNIEJĄCEGO partu zachowuje kanoniczną nazwę (styles.xml) i relację;
            // nie odwołujemy się do .Styles, więc Save nie nadpisze strumienia zserializowanym DOM.
            if (origMain.StyleDefinitionsPart != null)
            {
                var styles = targetMain.StyleDefinitionsPart ?? targetMain.AddNewPart<StyleDefinitionsPart>();
                using var s = origMain.StyleDefinitionsPart.GetStream(FileMode.Open, FileAccess.Read);
                styles.FeedData(s);
            }

            // theme — definicje theme fonts/colors (np. minor=Cambria), do których odwołują się
            // docDefaults (asciiTheme=minorHAnsi). Musi być spójny ze stylami.
            if (origMain.ThemePart != null)
            {
                var theme = targetMain.ThemePart ?? targetMain.AddNewPart<ThemePart>();
                using var s = origMain.ThemePart.GetStream(FileMode.Open, FileAccess.Read);
                theme.FeedData(s);
            }

            // fontTable — tabela fontów używanych w dokumencie.
            if (origMain.FontTablePart != null)
            {
                var fonts = targetMain.FontTablePart ?? targetMain.AddNewPart<FontTablePart>();
                using var s = origMain.FontTablePart.GetStream(FileMode.Open, FileAccess.Read);
                fonts.FeedData(s);
            }

            target.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Writes the document's headers and footers. The default variant is always emitted;
    /// first-page (DifferentFirstPage + FirstPageHtml) and even (DifferentOddEven + EvenHtml)
    /// variants are emitted as additional parts with type=First/Even references, and the
    /// section/settings opt-ins (titlePg, evenAndOddHeaders) are written so Word/round-trip
    /// import picks them up.
    /// </summary>
    private void AddHeaderAndFooter(WordprocessingDocument document, HeaderFooterContent? header, HeaderFooterContent? footer)
    {
        if (_mainPart == null) return;

        if (header != null && !string.IsNullOrWhiteSpace(header.Html))
        {
            WriteHeaderPart(header.Html, HeaderFooterValues.Default);

            if (header.DifferentFirstPage && !string.IsNullOrWhiteSpace(header.FirstPageHtml))
            {
                WriteHeaderPart(header.FirstPageHtml!, HeaderFooterValues.First);
                EnsureTitlePage();
            }

            if (header.DifferentOddEven && !string.IsNullOrWhiteSpace(header.EvenHtml))
            {
                WriteHeaderPart(header.EvenHtml!, HeaderFooterValues.Even);
                EnsureEvenAndOddHeaders(document);
            }
        }

        if (footer != null && !string.IsNullOrWhiteSpace(footer.Html))
        {
            WriteFooterPart(footer.Html, HeaderFooterValues.Default);

            if (footer.DifferentFirstPage && !string.IsNullOrWhiteSpace(footer.FirstPageHtml))
            {
                WriteFooterPart(footer.FirstPageHtml!, HeaderFooterValues.First);
                EnsureTitlePage();
            }

            if (footer.DifferentOddEven && !string.IsNullOrWhiteSpace(footer.EvenHtml))
            {
                WriteFooterPart(footer.EvenHtml!, HeaderFooterValues.Even);
                EnsureEvenAndOddHeaders(document);
            }
        }
    }

    private void WriteHeaderPart(string html, HeaderFooterValues type, SectionProperties? targetSection = null)
    {
        var headerPart = _mainPart!.AddNewPart<HeaderPart>();
        var headerElement = new Header();

        var prepared = html
            .Replace("{page}", "<span class=\"field-page\"></span>")
            .Replace("{pages}", "<span class=\"field-numpages\"></span>");
        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(prepared);

        // Image relationships must be scoped to THIS part — Word won't resolve cross-part rIds.
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

        AddHeaderReference(_mainPart.GetIdOfPart(headerPart), type, targetSection);
    }

    private void WriteFooterPart(string html, HeaderFooterValues type, SectionProperties? targetSection = null)
    {
        var footerPart = _mainPart!.AddNewPart<FooterPart>();
        var footerElement = new Footer();

        var prepared = html
            .Replace("{page}", "<span class=\"field-page\"></span>")
            .Replace("{pages}", "<span class=\"field-numpages\"></span>");
        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(prepared);

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

        AddFooterReference(_mainPart.GetIdOfPart(footerPart), type, targetSection);
    }

    private void EnsureTitlePage(SectionProperties? targetSection = null)
    {
        var sectionProps = targetSection ?? GetReferenceSectionProps();
        if (sectionProps == null) return;
        if (!sectionProps.Elements<TitlePage>().Any())
        {
            sectionProps.Append(new TitlePage());
        }
    }

    private static void EnsureEvenAndOddHeaders(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return;
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        if (!settingsPart.Settings.Elements<EvenAndOddHeaders>().Any())
        {
            settingsPart.Settings.AppendChild(new EvenAndOddHeaders());
        }
        settingsPart.Settings.Save();
    }

    private SectionProperties? GetOrCreateSectionProps()
    {
        var body = _mainPart?.Document?.Body;
        if (body == null) return null;
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }
        return sectionProps;
    }

    /// <summary>
    /// sectPr, do którego wpinamy referencje nagłówka/stopki i titlePg: PIERWSZY sectPr
    /// w kolejności dokumentu. W dokumencie wielosekcyjnym to sectPr pierwszego markera —
    /// sekcje bez własnych referencji dziedziczą je w Wordzie z poprzedniej sekcji; gdyby
    /// referencje trafiły tylko do body-level sectPr (ostatnia sekcja), wcześniejsze strony
    /// nie miałyby nagłówka/stopki.
    /// </summary>
    private SectionProperties? GetReferenceSectionProps() => _firstSectionProps ?? GetOrCreateSectionProps();

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
                if (!pendingTextParagraph.Elements<Run>().Any()
                    && !pendingTextParagraph.Elements<Hyperlink>().Any()
                    && !pendingTextParagraph.Elements<SimpleField>().Any())
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
                    if (child.HasClass("field-page") || child.HasClass("page-number"))
                    {
                        pendingTextParagraph.Append(BuildFieldRun(" PAGE ", child));
                    }
                    else if (child.HasClass("field-numpages"))
                    {
                        pendingTextParagraph.Append(BuildFieldRun(" NUMPAGES ", child));
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

    private void AddHeaderReference(string headerPartId, HeaderFooterValues type, SectionProperties? targetSection = null)
    {
        var sectionProps = targetSection ?? GetReferenceSectionProps();
        if (sectionProps == null) return;
        sectionProps.InsertAt(new HeaderReference { Type = type, Id = headerPartId }, 0);
    }

    private void AddFooterReference(string footerPartId, HeaderFooterValues type, SectionProperties? targetSection = null)
    {
        var sectionProps = targetSection ?? GetReferenceSectionProps();
        if (sectionProps == null) return;
        sectionProps.InsertAt(new FooterReference { Type = type, Id = footerPartId }, 0);
    }

    /// <summary>
    /// Wpina WŁASNE nagłówki/stopki sekcji ≥ 1 na ich sectPr: sekcja o indeksie s (0-based)
    /// → _emittedSectionProps[s] (paragraph-level sectPr zamykający tę sekcję) albo body-level
    /// sectPr, gdy s to sekcja ostatnia. Sekcja 0 (bazowa) idzie standardową ścieżką
    /// AddHeaderAndFooter → pierwszy sectPr; sekcje bez wpisu dziedziczą w Wordzie.
    /// </summary>
    private void AddSectionHeadersFooters(WordprocessingDocument document, IReadOnlyList<SectionHeaderFooter>? sections)
    {
        if (sections == null || sections.Count == 0) return;
        var body = _mainPart?.Document?.Body;
        if (body == null) return;
        var bodySectPr = body.Elements<SectionProperties>().FirstOrDefault();

        foreach (var entry in sections)
        {
            if (entry.SectionIndex < 1) continue; // sekcja 0 = pola bazowe Header/Footer

            SectionProperties? target = null;
            if (entry.SectionIndex < _emittedSectionProps.Count)
                target = _emittedSectionProps[entry.SectionIndex];
            else if (entry.SectionIndex == _emittedSectionProps.Count)
                target = bodySectPr; // ostatnia sekcja
            if (target == null) continue; // markery usunięte z treści → sekcja nie istnieje

            if (entry.Header is { } h && !string.IsNullOrWhiteSpace(h.Html))
            {
                WriteHeaderPart(h.Html, HeaderFooterValues.Default, target);
                if (h.DifferentFirstPage && !string.IsNullOrWhiteSpace(h.FirstPageHtml))
                {
                    WriteHeaderPart(h.FirstPageHtml!, HeaderFooterValues.First, target);
                    EnsureTitlePage(target);
                }
                if (h.DifferentOddEven && !string.IsNullOrWhiteSpace(h.EvenHtml))
                {
                    WriteHeaderPart(h.EvenHtml!, HeaderFooterValues.Even, target);
                    EnsureEvenAndOddHeaders(document);
                }
            }

            if (entry.Footer is { } f && !string.IsNullOrWhiteSpace(f.Html))
            {
                WriteFooterPart(f.Html, HeaderFooterValues.Default, target);
                if (f.DifferentFirstPage && !string.IsNullOrWhiteSpace(f.FirstPageHtml))
                {
                    WriteFooterPart(f.FirstPageHtml!, HeaderFooterValues.First, target);
                    EnsureTitlePage(target);
                }
                if (f.DifferentOddEven && !string.IsNullOrWhiteSpace(f.EvenHtml))
                {
                    WriteFooterPart(f.EvenHtml!, HeaderFooterValues.Even, target);
                    EnsureEvenAndOddHeaders(document);
                }
            }
        }
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
            
            // Kolejność w <w:pPr> wg schematu (EG_PPrBase): keepNext → keepLines → spacing → outlineLvl.
            var paraProps = new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = headingSpaceBefore[i - 1].ToString(), After = "0" },
                new OutlineLevel { Val = i - 1 }
            );
            headingStyle.Append(paraProps);
            
            // Kolejność dzieci w <w:rPr> jest narzucona schematem OOXML (EG_RPrBase):
            // rFonts → b → i → … → color → … → sz. Zła kolejność (np. sz przed color, b po color)
            // = błąd schematu → Word zgłasza „dokument uszkodzony". Emitujemy w poprawnej sekwencji.
            var runPropsElements = new List<OpenXmlElement>
            {
                new RunFonts { Ascii = headingFont, HighAnsi = headingFont },
            };
            if (headingBold[i - 1]) runPropsElements.Add(new Bold());
            if (headingItalic[i - 1]) runPropsElements.Add(new Italic());
            runPropsElements.Add(new Color { Val = headingColors[i - 1] });
            runPropsElements.Add(new FontSize { Val = headingSizes[i - 1] });

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
                if (IsSectionBreakNode(node))
                {
                    elements.Add(CreateSectionBreakParagraph(node));
                }
                else if (IsPageBreakNode(node))
                {
                    // Marker sekcji tuż za page-breakiem = przerwa sekcji typu nextPage;
                    // sectPr sam łamie stronę, dodatkowy w:br type=page dawałby pustą stronę.
                    if (!NextElementSiblingIsSectionBreak(node))
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

        // Tab-stopy per akapit (reader: data-tab-stops="pos:align[:leader];…", pos w twips) —
        // odtwarzane jako w:tabs, żeby pozycje/wyrównania/leadery nie ginęły przy zapisie.
        var tabStopsAttr = node.GetAttributeValue("data-tab-stops", "");
        if (!string.IsNullOrEmpty(tabStopsAttr) && ParseTabStops(tabStopsAttr) is { } tabs)
        {
            props.Append(tabs);
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
            EnsureNumberingPart();

            // Tożsamość listy z readera: fragmenty tej samej listy logicznej (np. rozdzielone
            // zwykłym akapitem) niosą ten sam data-num-id → współdzielą jedną NumberingInstance,
            // więc Word kontynuuje numerację. Bez atrybutu (lista utworzona w edytorze) — nowa.
            var htmlListId = node.GetAttributeValue("data-num-id", "");
            if (htmlListId.Length > 0 && _numIdByHtmlList.TryGetValue(htmlListId, out var existingNumId))
            {
                numId = existingNumId;
            }
            else
            {
                // Przeskanuj strukturę listy aby określić format dla każdego poziomu
                var levelFormats = new Dictionary<int, bool>();
                ScanListLevels(node, ordered, level, levelFormats);

                // Wykryj poziomy z punktatorem obrazkowym (DocxToHtmlConverter wstawia
                // <span class="list-marker"><img .../></span> jako wizualny marker).
                // Dla takich poziomów wyłączymy automatyczny punktator Worda, żeby nie
                // dublować markera (kropka + grafika).
                var pictureBulletLevels = new HashSet<int>();
                ScanPictureBulletLevels(node, level, pictureBulletLevels);

                // Definicje poziomów round-tripowane z readera (data-num-fmt / data-lvl-text /
                // data-start / data-bullet-font) — zachowują oryginalny format numeracji
                // (np. upperRoman, "%1)", start=5) zamiast hardkodowanej drabinki.
                var levelSpecs = new Dictionary<int, HtmlListLevelSpec>();
                ScanListLevelSpecs(node, level, levelSpecs);

                var abstractNumId = CreateAbstractNumbering(levelFormats, pictureBulletLevels, levelSpecs);
                numId = CreateNumberingInstance(abstractNumId);
                if (htmlListId.Length > 0)
                    _numIdByHtmlList[htmlListId] = numId;
            }
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
                var liChildName = liChild.Name.ToLower();
                if (liChildName == "ul" || liChildName == "ol")
                    continue; // Zagnieżdżona lista będzie obsłużona osobno

                // <span class="list-marker"> jest artefaktem prezentacyjnym dodanym przez
                // DocxToHtmlConverter, żeby przeglądarka pokazała niestandardowy punktator
                // (obrazek lub znak Wingdings/Symbol). Przy eksporcie:
                //   - jeżeli zawiera <img> (picture bullet) — zachowaj sam obrazek jako
                //     wiodący inline run; poziom ma format=None, więc Word nie doda
                //     dodatkowej kropki.
                //   - tekstowy marker (np. ✓, ✗) pomijamy — Word wstawi własny automatyczny
                //     punktator z definicji numeracji.
                if (liChildName == "span" && IsListMarkerSpan(liChild))
                {
                    var img = liChild.SelectSingleNode(".//img");
                    if (img != null)
                    {
                        foreach (var run in CreateRunsFromNode(img, liBaseProps))
                            para.Append(run);
                    }
                    continue;
                }

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
    /// Definicja poziomu listy odczytana z data-* kontenera (round-trip z DocxToHtmlConverter).
    /// </summary>
    private readonly record struct HtmlListLevelSpec(string? Fmt, string? LvlText, int Start, string? BulletFont);

    /// <summary>
    /// Zbiera definicje poziomów z atrybutów data-num-fmt / data-lvl-text / data-start /
    /// data-bullet-font na kontenerach ul/ol (każdy zagnieżdżony kontener opisuje swój poziom).
    /// Pierwsze napotkane wystąpienie poziomu wygrywa.
    /// </summary>
    private static void ScanListLevelSpecs(HtmlNode node, int level, Dictionary<int, HtmlListLevelSpec> specs)
    {
        if (!specs.ContainsKey(level))
        {
            var fmt = node.GetAttributeValue("data-num-fmt", "");
            var lvlText = node.GetAttributeValue("data-lvl-text", "");
            var startRaw = node.GetAttributeValue("data-start", "");
            var bulletFont = node.GetAttributeValue("data-bullet-font", "");
            if (fmt.Length > 0 || lvlText.Length > 0 || startRaw.Length > 0)
            {
                _ = int.TryParse(startRaw, out var start);
                specs[level] = new HtmlListLevelSpec(
                    fmt.Length > 0 ? fmt : null,
                    lvlText.Length > 0 ? HtmlEntity.DeEntitize(lvlText) : null,
                    start > 0 ? start : 1,
                    bulletFont.Length > 0 ? HtmlEntity.DeEntitize(bulletFont) : null);
            }
        }

        foreach (var child in node.ChildNodes)
        {
            if (child.Name.ToLower() != "li") continue;
            var nested = child.SelectNodes("./ul|./ol");
            if (nested == null) continue;
            foreach (var nestedList in nested)
                ScanListLevelSpecs(nestedList, level + 1, specs);
        }
    }

    /// <summary>Mapuje token data-num-fmt (nazwy w:numFmt) na NumberFormatValues.</summary>
    private static bool TryMapNumFmt(string token, out NumberFormatValues fmt)
    {
        switch (token)
        {
            case "decimal": fmt = NumberFormatValues.Decimal; return true;
            case "decimalZero": fmt = NumberFormatValues.DecimalZero; return true;
            case "lowerLetter": fmt = NumberFormatValues.LowerLetter; return true;
            case "upperLetter": fmt = NumberFormatValues.UpperLetter; return true;
            case "lowerRoman": fmt = NumberFormatValues.LowerRoman; return true;
            case "upperRoman": fmt = NumberFormatValues.UpperRoman; return true;
            case "bullet": fmt = NumberFormatValues.Bullet; return true;
            case "none": fmt = NumberFormatValues.None; return true;
            default: fmt = NumberFormatValues.Decimal; return false;
        }
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
    /// Wykrywa poziomy listy, które używają punktatora obrazkowego — czyli mają
    /// <c>&lt;span class="list-marker"&gt;&lt;img/&gt;&lt;/span&gt;</c> wewnątrz &lt;li&gt;.
    /// Dla takich poziomów wyłączymy automatyczny punktator Worda.
    /// </summary>
    private static void ScanPictureBulletLevels(HtmlNode node, int level, HashSet<int> pictureBulletLevels)
    {
        foreach (var child in node.ChildNodes)
        {
            if (child.Name.ToLower() != "li") continue;

            // Sprawdź bezpośrednie dzieci <li>, czy któreś z nich jest markerem obrazkowym.
            foreach (var liChild in child.ChildNodes)
            {
                if (liChild.Name.ToLower() != "span") continue;
                if (!IsListMarkerSpan(liChild)) continue;
                if (liChild.SelectSingleNode(".//img") != null)
                {
                    pictureBulletLevels.Add(level);
                    break;
                }
            }

            var nested = child.SelectNodes("./ul|./ol");
            if (nested == null) continue;
            foreach (var nestedList in nested)
                ScanPictureBulletLevels(nestedList, level + 1, pictureBulletLevels);
        }
    }

    /// <summary>
    /// Czy element to <c>&lt;span class="list-marker"&gt;</c> emitowany przez DocxToHtmlConverter.
    /// </summary>
    private static bool IsListMarkerSpan(HtmlNode node)
    {
        var cls = node.GetAttributeValue("class", "");
        if (string.IsNullOrEmpty(cls)) return false;
        foreach (var token in cls.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (token.Equals("list-marker", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
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
    private int CreateAbstractNumbering(
        Dictionary<int, bool> levelFormats,
        HashSet<int>? pictureBulletLevels = null,
        Dictionary<int, HtmlListLevelSpec>? levelSpecs = null)
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

            HtmlListLevelSpec? spec = levelSpecs != null && levelSpecs.TryGetValue(lvl, out var s) ? s : null;

            var levelDef = new Level { LevelIndex = lvl };
            levelDef.Append(new StartNumberingValue { Val = spec?.Start ?? 1 });

            // Poziom z punktatorem obrazkowym — wyłącz automatyczny marker Worda,
            // grafika została wstawiona jako wiodący inline run w paragrafie.
            if (pictureBulletLevels != null && pictureBulletLevels.Contains(lvl))
            {
                levelDef.Append(new NumberingFormat { Val = NumberFormatValues.None });
                levelDef.Append(new LevelText { Val = string.Empty });
            }
            else if (spec is { Fmt: not null } sp && TryMapNumFmt(sp.Fmt, out var mappedFmt))
            {
                // Round-trip z data-*: dokładny format/lvlText/font punktatora z oryginału.
                if (mappedFmt == NumberFormatValues.Bullet)
                {
                    levelDef.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
                    levelDef.Append(new LevelText { Val = !string.IsNullOrEmpty(sp.LvlText) ? sp.LvlText : "•" });
                    var bulletFont = sp.BulletFont;
                    levelDef.Append(new NumberingSymbolRunProperties(string.IsNullOrEmpty(bulletFont)
                        ? new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }
                        : new RunFonts { Ascii = bulletFont, HighAnsi = bulletFont, Hint = FontTypeHintValues.Default }));
                }
                else
                {
                    levelDef.Append(new NumberingFormat { Val = mappedFmt });
                    // lvlText z placeholderem %N; bez niego (lub uszkodzony) — standardowe "%N.".
                    var lvlText = !string.IsNullOrEmpty(sp.LvlText) && sp.LvlText.Contains('%')
                        ? sp.LvlText
                        : $"%{lvl + 1}.";
                    levelDef.Append(new LevelText { Val = lvlText });
                    levelDef.Append(new NumberingSymbolRunProperties(
                        new RunFonts { Hint = FontTypeHintValues.Default }
                    ));
                }
            }
            else if (isOrdered)
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
            else if (!isOrdered)
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

        // Referencja stylu tabeli Worda zachowana przez reader w data-* — emitujemy ją z powrotem.
        // w:tblStyle musi być PIERWSZYM dzieckiem tblPr (kolejność schematu). Rozwiązane wartości
        // stylu i tak są w inline CSS komórek, więc wygląd odtwarza formatowanie bezpośrednie,
        // a referencja stylu przeżywa round-trip (edycja w Wordzie dalej "widzi" styl).
        var tblStyleId = node.GetAttributeValue("data-tbl-style", "");
        if (!string.IsNullOrEmpty(tblStyleId))
            tableProps.Append(new TableStyle { Val = System.Net.WebUtility.HtmlDecode(tblStyleId) });

        // Szerokość (reader emituje też ułamkowe %: 66.66%)
        var tableWidthMatch = Regex.Match(tableStyle, @"width:\s*([\d.]+)(px|%)?");
        if (tableWidthMatch.Success)
        {
            var widthValue = double.Parse(tableWidthMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var unit = tableWidthMatch.Groups[2].Value;

            if (unit == "%")
            {
                tableProps.Append(new TableWidth { Width = ((int)Math.Round(widthValue * 50)).ToString(), Type = TableWidthUnitValues.Pct });
            }
            else if (unit == "px" || string.IsNullOrEmpty(unit))
            {
                tableProps.Append(new TableWidth { Width = ((int)Math.Round(widthValue * 15)).ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }
        else
        {
            // No explicit width (e.g. CSS width:auto) → AUTO, sizing to content/grid like Word.
            // Previously this forced 100% (pct 5000), which stretched content-sized tables to
            // full text width. Patrz analiza orginał_GOOD (tblW auto) vs zapisany_BAD (pct 5000).
            tableProps.Append(new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
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

        // Odstęp między komórkami: preferuj dokładne twips z data-*, inaczej border-spacing px.
        var cellSpacingTwAttr = node.GetAttributeValue("data-cell-spacing-tw", "");
        if (int.TryParse(cellSpacingTwAttr, out var cellSpacingTw) && cellSpacingTw > 0)
        {
            tableProps.Append(new TableCellSpacing { Width = cellSpacingTw.ToString(), Type = TableWidthUnitValues.Dxa });
        }
        else
        {
            var spacingMatch = Regex.Match(tableStyle, @"border-spacing:\s*([\d.]+)px");
            if (spacingMatch.Success)
            {
                var spacingPx = double.Parse(spacingMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                if (spacingPx > 0)
                    tableProps.Append(new TableCellSpacing { Width = ((int)Math.Round(spacingPx * 15)).ToString(), Type = TableWidthUnitValues.Dxa });
            }
        }

        // Wcięcie tabeli (margin-left w px, nie 'auto') → w:tblInd. Wcześniej gubione na eksporcie.
        var indentMatch = Regex.Match(tableStyle, @"margin-left:\s*(-?[\d.]+)px");
        if (indentMatch.Success)
        {
            var indentPx = double.Parse(indentMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            if (Math.Abs(indentPx) > 0.01)
                tableProps.Append(new TableIndentation { Width = (int)Math.Round(indentPx * 15), Type = TableWidthUnitValues.Dxa });
        }

        // Parsuj obramowania tabeli z CSS
        var borderMatch = Regex.Match(tableStyle, @"(?<![a-z-])border:\s*([\d.]+)px\s+(\w+)\s+#?([a-fA-F0-9]{3,6})");
        if (borderMatch.Success)
        {
            var bSize = CssPxToBorderEighthPoints(borderMatch.Groups[1].Value);
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

        // Reader emituje table-layout:fixed dla tabel z geometrią kolumn z tblGrid —
        // wymuszanie Autofit gubiło układ Worda przy każdym zapisie (autosave!).
        var isFixedLayout = tableStyle.Contains("table-layout:fixed");
        tableProps.Append(new TableLayout { Type = isFixedLayout ? TableLayoutValues.Fixed : TableLayoutValues.Autofit });
        
        // Domyślne marginesy komórek = domyślne Worda (TableNormal): top/bottom=0, left/right=108
        // twips. Wcześniej hardkodowane 40/80 dodawało pionowy margines do każdej komórki (tabele
        // rosły w pionie). Per-komórkowe tcMar z CSS i tak nadpisują tę wartość.
        tableProps.Append(new TableCellMarginDefault(
            new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
            new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
            new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
            new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa }
        ));

        // w:tblLook (flagi formatowania warunkowego stylu) — round-trip z data-tbl-look.
        var tblLookHex = node.GetAttributeValue("data-tbl-look", "");
        if (Regex.IsMatch(tblLookHex, "^[0-9A-Fa-f]{4}$"))
        {
            var lookMask = int.Parse(tblLookHex, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture);
            tableProps.Append(new TableLook
            {
                Val = tblLookHex.ToUpperInvariant(),
                FirstRow = (lookMask & 0x0020) != 0,
                LastRow = (lookMask & 0x0040) != 0,
                FirstColumn = (lookMask & 0x0080) != 0,
                LastColumn = (lookMask & 0x0100) != 0,
                NoHorizontalBand = (lookMask & 0x0200) != 0,
                NoVerticalBand = (lookMask & 0x0400) != 0
            });
        }

        table.Append(tableProps);

        // Oblicz liczbę kolumn
        var maxCols = 0;
        // Tylko wiersze TEJ tabeli (bezpośrednie lub w thead/tbody/tfoot) — `.//tr` schodziło
        // do tabel ZAGNIEŻDŻONYCH i duplikowało ich wiersze jako wiersze tabeli zewnętrznej.
        var rowNodes = node.SelectNodes("./tr | ./thead/tr | ./tbody/tr | ./tfoot/tr");
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

        // Siatka tabeli. Szerokości kolumn pochodzą z <colgroup> (reader emituje je z tblGrid);
        // wcześniej grid był odtwarzany z samej LICZBY komórek bez szerokości, więc geometria
        // kolumn Worda ginęła przy każdym zapisie.
        var colWidthsTwips = ReadColgroupWidthsTwips(node);
        var gridColCount = Math.Max(maxCols, colWidthsTwips.Count);
        if (gridColCount > 0)
        {
            var grid = new TableGrid();
            for (int i = 0; i < gridColCount; i++)
            {
                var col = new GridColumn();
                if (i < colWidthsTwips.Count && colWidthsTwips[i] > 0)
                    col.Width = colWidthsTwips[i].ToString();
                grid.Append(col);
            }
            table.Append(grid);
        }

        // Przetwórz wiersze. HTML pomija komórki przykryte przez rowspan w kolejnych
        // wierszach — OOXML wymaga tam jawnych komórek kontynuacji (vMerge bez val).
        // Bez nich komórki przesuwały się w lewo i tabela była uszkodzona w Wordzie.
        // activeRowSpans: kolumna gridu → (pozostałe wiersze scalenia, rozpiętość kolumn).
        var activeRowSpans = new Dictionary<int, (int RemainingRows, int ColSpan)>();
        if (rowNodes != null)
        {
            foreach (var rowNode in rowNodes)
            {
                var row = new TableRow();
                var gridCursor = 0;
                var spansStartedThisRow = new HashSet<int>();

                void AppendPendingContinuations()
                {
                    while (activeRowSpans.TryGetValue(gridCursor, out var span))
                    {
                        row.Append(CreateVerticalMergeContinuationCell(span.ColSpan));
                        gridCursor += span.ColSpan;
                    }
                }

                // Właściwości wiersza (kolejność schematu trPr: cantSplit → trHeight → tblHeader).
                // Wysokość: preferuj dokładne twips + regułę z data-* (round-trip bez strat
                // px→twips i bez gubienia hRule=exact); fallback: height/min-height px → atLeast.
                var rowStyle = rowNode.GetAttributeValue("style", "");
                var rowProps = new TableRowProperties();

                if (rowNode.GetAttributeValue("data-cant-split", "") == "1")
                    rowProps.Append(new CantSplit());

                var hRule = rowNode.GetAttributeValue("data-row-hrule", "") == "exact"
                    ? HeightRuleValues.Exact
                    : HeightRuleValues.AtLeast;
                var heightTwAttr = rowNode.GetAttributeValue("data-row-height-tw", "");
                if (uint.TryParse(heightTwAttr, out var heightTwFromAttr) && heightTwFromAttr > 0)
                {
                    rowProps.Append(new TableRowHeight { Val = heightTwFromAttr, HeightType = hRule });
                }
                else
                {
                    var rowHeightMatch = Regex.Match(rowStyle, @"(?:min-)?height:\s*([\d.]+)px");
                    if (rowHeightMatch.Success)
                    {
                        var heightPx = (int)Math.Round(double.Parse(rowHeightMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
                        var heightTwips = PxToTwips(heightPx);
                        if (heightTwips > 0)
                            rowProps.Append(new TableRowHeight { Val = (uint)heightTwips, HeightType = hRule });
                    }
                }

                // Wiersz nagłówkowy powtarzany na kolejnych stronach (w:tblHeader).
                if (rowNode.GetAttributeValue("data-tbl-header", "") == "1")
                    rowProps.Append(new TableHeader());

                if (rowProps.HasChildren)
                    row.Append(rowProps);

                var cells = rowNode.SelectNodes("./td|./th");
                if (cells != null)
                {
                    foreach (var cellNode in cells)
                    {
                        AppendPendingContinuations();

                        var cell = new TableCell();
                        var cellProps = new TableCellProperties();

                        // Colspan
                        var colspanAttr = cellNode.GetAttributeValue("colspan", "1");
                        if (!int.TryParse(colspanAttr, out var colspan) || colspan < 1) colspan = 1;
                        if (colspan > 1)
                            cellProps.Append(new GridSpan { Val = colspan });

                        // Rowspan
                        var rowspanAttr = cellNode.GetAttributeValue("rowspan", "1");
                        if (int.TryParse(rowspanAttr, out var rowspan) && rowspan > 1)
                        {
                            cellProps.Append(new VerticalMerge { Val = MergedCellValues.Restart });
                            activeRowSpans[gridCursor] = (rowspan - 1, colspan);
                            spansStartedThisRow.Add(gridCursor);
                        }
                        gridCursor += colspan;

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

                // Kontynuacje scaleń wypadające ZA ostatnią komórką wiersza.
                AppendPendingContinuations();

                // Scalenia rozpoczęte w tym wierszu obejmują dopiero KOLEJNE wiersze.
                foreach (var col in activeRowSpans.Keys.ToList())
                {
                    if (spansStartedThisRow.Contains(col)) continue;
                    var (remaining, span) = activeRowSpans[col];
                    if (remaining <= 1) activeRowSpans.Remove(col);
                    else activeRowSpans[col] = (remaining - 1, span);
                }

                table.Append(row);
            }
        }

        return table;
    }

    /// <summary>
    /// Komórka kontynuacji scalenia pionowego (w:vMerge bez w:val = continue) —
    /// odpowiednik komórki, którą HTML pomija pod komórką z rowspan.
    /// </summary>
    private static TableCell CreateVerticalMergeContinuationCell(int colSpan)
    {
        var props = new TableCellProperties();
        if (colSpan > 1)
            props.Append(new GridSpan { Val = colSpan });
        props.Append(new VerticalMerge());
        return new TableCell(props, new Paragraph());
    }

    /// <summary>
    /// Szerokości kolumn (twips) z &lt;colgroup&gt; tabeli — reader emituje je z w:tblGrid
    /// (px). Kolumna bez szerokości daje 0 (Word rozłoży resztę).
    /// </summary>
    private static List<int> ReadColgroupWidthsTwips(HtmlNode tableNode)
    {
        var result = new List<int>();
        var cols = tableNode.SelectNodes("./colgroup/col");
        if (cols == null) return result;

        foreach (var col in cols)
        {
            var style = col.GetAttributeValue("style", "");
            var m = Regex.Match(style, @"width:\s*([\d.]+)px");
            result.Add(m.Success
                ? (int)Math.Round(OoxmlUnits.PixelsToTwips(
                    double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)))
                : 0);
        }
        return result;
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
        
        // Szerokość (min-width pomijamy — to artefakt edytora, nie geometria Worda).
        var widthMatch = Regex.Match(style, @"(?<![a-z-])width:\s*([\d.]+)(px|%)?");
        if (widthMatch.Success)
        {
            var widthVal = double.Parse(widthMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var widthUnit = widthMatch.Groups[2].Value;

            if (widthUnit == "%")
                cellProps.Append(new TableCellWidth { Width = ((int)Math.Round(widthVal * 50)).ToString(), Type = TableWidthUnitValues.Pct });
            else
                cellProps.Append(new TableCellWidth { Width = ((int)Math.Round(widthVal * 15)).ToString(), Type = TableWidthUnitValues.Dxa });
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
                var size = CssPxToBorderEighthPoints(match.Groups[1].Value);
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
                var size = CssPxToBorderEighthPoints(borderAll.Groups[1].Value);
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
    /// <summary>
    /// CSS px → w:sz (1/8 pt): sz = px × 72/96 × 8 = px × 6. Symetryczne do readera (sz/6 → px);
    /// wcześniejsze ×8 pogrubiało każdą linię o 33% przy każdym round-tripie. Minimum 2 (0.25 pt).
    /// </summary>
    private static uint CssPxToBorderEighthPoints(string px)
    {
        var v = double.Parse(px, System.Globalization.CultureInfo.InvariantCulture);
        return (uint)Math.Max(2, Math.Round(v * 6));
    }

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
    /// <summary>
    /// Zwraca efektywny data:URL obrazu do zapisu. Dla legacy EMF/WMF `src` to tylko placeholder
    /// SVG (podgląd w przeglądarce) — prawdziwy metafile jest w `data-original-src`. Zapis placeholdera
    /// SVG jako gołego `a:blip` daje NIEPOPRAWNY DOCX (Word: „dokument uszkodzony"), więc preferujemy oryginał.
    /// </summary>
    private static string ResolveImageSrc(HtmlNode node)
    {
        var original = node.GetAttributeValue("data-original-src", "");
        if (!string.IsNullOrEmpty(original) && original.StartsWith("data:")) return original;
        return node.GetAttributeValue("src", "");
    }

    private Paragraph? ConvertImageElement(HtmlNode node)
    {
        var src = ResolveImageSrc(node);
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
        var src = ResolveImageSrc(node);
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
            "image/x-emf" or "image/emf" => ImagePartType.Emf,
            "image/x-wmf" or "image/wmf" => ImagePartType.Wmf,
            "image/svg+xml" => ImagePartType.Svg,
            _ => ImagePartType.Jpeg
        };

        // HARD GUARD: goły `a:blip` na SVG (bez rastrowego fallbacku) jest NIEPOPRAWNY w DOCX —
        // Word zgłasza uszkodzony plik. Placeholdery legacy niosą prawdziwy metafile w
        // `data-original-src` (obsłużone w ResolveImageSrc), więc tu SVG = brak fallbacku → pomiń
        // (kontrolowana strata, NIGDY uszkodzony dokument).
        if (imagePartType == ImagePartType.Svg)
            return null;

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

        // Alt text → round-tripped into wp:docPr/@descr. DeEntitize so the stored text is
        // the real string (the reader re-encodes once); otherwise entities double-escape.
        var altText = HtmlEntity.DeEntitize(node.GetAttributeValue("alt", "")) ?? string.Empty;

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

            widthEmu = OoxmlUnits.PixelsToEmu(width);
            heightEmu = OoxmlUnits.PixelsToEmu(height);
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
        if (widthEmu < OoxmlUnits.EmuPerPixel) widthEmu = OoxmlUnits.EmuPerPixel;   // min 1 px
        if (heightEmu < OoxmlUnits.EmuPerPixel) heightEmu = OoxmlUnits.EmuPerPixel;

        _imageCounter++;

        // Word-like positioning: when the editor marks the image as floating
        // (data-pos-mode="front"|"behind") we emit wp:anchor with position offsets;
        // otherwise it stays an inline image (default OOXML behaviour). Offsets are
        // read from data-x-emu / data-y-emu set by the editor on drag-end.
        var posMode = node.GetAttributeValue("data-pos-mode", "");
        var isFloating = posMode == "front" || posMode == "behind";

        // Optional border (a:ln in pic:spPr): width in EMU = px * 9525.
        int.TryParse(node.GetAttributeValue("data-border-width", "0"), out var borderWidthPx);
        var borderColor = node.GetAttributeValue("data-border-color", "").TrimStart('#');
        var borderStyle = node.GetAttributeValue("data-border-style", "solid");

        // Optional crop (a:srcRect on pic:blipFill): l/t/r/b in 1/1000 of a percent.
        int.TryParse(node.GetAttributeValue("data-crop-l", "0"), out var cropL);
        int.TryParse(node.GetAttributeValue("data-crop-r", "0"), out var cropR);
        int.TryParse(node.GetAttributeValue("data-crop-t", "0"), out var cropT);
        int.TryParse(node.GetAttributeValue("data-crop-b", "0"), out var cropB);
        var hasCrop = cropL > 0 || cropR > 0 || cropT > 0 || cropB > 0;

        // BlipFill — with optional srcRect carrying the crop percentages.
        var blip = new DocumentFormat.OpenXml.Drawing.Blip { Embed = relationshipId };
        var blipFill = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(blip);
        if (hasCrop)
        {
            blipFill.Append(new DocumentFormat.OpenXml.Drawing.SourceRectangle
            {
                Left = cropL * 1000,
                Right = cropR * 1000,
                Top = cropT * 1000,
                Bottom = cropB * 1000
            });
        }
        blipFill.Append(new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle()));

        // ShapeProperties — with optional outline (a:ln) for the border.
        var shapeProps = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthEmu, Cy = heightEmu }),
            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                new DocumentFormat.OpenXml.Drawing.AdjustValueList())
            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle });
        if (borderWidthPx > 0 && System.Text.RegularExpressions.Regex.IsMatch(borderColor, "^[0-9A-Fa-f]{6}$"))
        {
            var dashStyle = borderStyle switch
            {
                "dashed" => DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash,
                "dotted" => DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot,
                _ => DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid
            };
            shapeProps.Append(new DocumentFormat.OpenXml.Drawing.Outline(
                new DocumentFormat.OpenXml.Drawing.SolidFill(
                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = borderColor.ToUpperInvariant() }),
                new DocumentFormat.OpenXml.Drawing.PresetDash { Val = dashStyle })
            { Width = borderWidthPx * OoxmlUnits.EmuPerPixel });
        }

        var graphic = new DocumentFormat.OpenXml.Drawing.Graphic(
            new DocumentFormat.OpenXml.Drawing.GraphicData(
                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = (uint)_imageCounter, Name = $"Image{_imageCounter}" },
                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                    blipFill,
                    shapeProps)
            )
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

        if (!isFloating)
        {
            return new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    BuildImageDocProperties((uint)_imageCounter, $"Image{_imageCounter}", altText),
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
                    graphic
                )
            );
        }

        // Floating mode: wp:anchor with position offsets and the "no wrap" mode that
        // matches Word's "Behind text" / "In front of text" options (no text reflow).
        long.TryParse(node.GetAttributeValue("data-x-emu", "0"), out var xEmu);
        long.TryParse(node.GetAttributeValue("data-y-emu", "0"), out var yEmu);
        var behind = posMode == "behind";

        var anchor = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.SimplePosition { X = 0L, Y = 0L },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset(xEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset(yEmu.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmu, Cy = heightEmu },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone(),
            BuildImageDocProperties((uint)_imageCounter, $"Image{_imageCounter}", altText),
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
            graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
            SimplePos = false,
            RelativeHeight = (uint)(251_660_288 + _imageCounter),
            BehindDoc = behind,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true
        };

        return new Drawing(anchor);
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
                    // Literalny znak tabulacji → w:tab (element), nie tekst w w:t — Word nie
                    // renderuje tabów zapisanych w treści w:t, więc ginęły po round-tripie.
                    var parts = text.Split('\t');
                    for (int pi = 0; pi < parts.Length; pi++)
                    {
                        if (pi > 0)
                        {
                            var tabRun = new Run();
                            if (inheritedProps != null)
                                tabRun.Append(inheritedProps.CloneNode(true));
                            tabRun.Append(new TabChar());
                            runs.Add(tabRun);
                        }
                        if (parts[pi].Length == 0) continue;
                        var run = new Run();
                        if (inheritedProps != null)
                            run.Append(inheritedProps.CloneNode(true));
                        run.Append(new Text(parts[pi]) { Space = SpaceProcessingModeValues.Preserve });
                        runs.Add(run);
                    }
                }
                break;

            case HtmlNodeType.Element:
                var tagName = node.Name.ToLower();

                // Manualny page break: reader emituje <div class="page-break"> WEWNĄTRZ akapitu
                // (Break siedzi w runie), więc trafia tu, a nie do bloku. Bez tego znak rozpoczęcia
                // nowej strony ginął po round-tripie (R-15). Jeden węzeł page-break → jeden Break,
                // bez duplikacji; dokument bez page-breaków nie dostaje żadnego.
                if (IsPageBreakNode(node))
                {
                    runs.Add(new Run(new Break { Type = BreakValues.Page }));
                    break;
                }

                // Segment pozycyjny tab-stopu (reader: nagłówek/stopka z w:tabs) — segment
                // zaczyna się od tabulatora; pozycję odtwarza pPr/w:tabs z data-tab-stops.
                if (node.HasClass("docx-tab-seg"))
                {
                    var segTab = new Run();
                    if (inheritedProps != null)
                        segTab.Append(inheritedProps.CloneNode(true));
                    segTab.Append(new TabChar());
                    runs.Add(segTab);
                    foreach (var segChild in node.ChildNodes)
                        runs.AddRange(CreateRunsFromNode(segChild, inheritedProps));
                    break;
                }

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
    /// Parsuje atrybut data-tab-stops ("4536:center;9072:right:dot") na w:tabs.
    /// Zwraca null przy braku poprawnych wpisów.
    /// </summary>
    private static Tabs? ParseTabStops(string attr)
    {
        var tabs = new Tabs();
        foreach (var entry in attr.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = entry.Split(':');
            if (parts.Length < 2 || !int.TryParse(parts[0], out var pos)) continue;

            var val = parts[1] switch
            {
                "center" => TabStopValues.Center,
                "right" => TabStopValues.Right,
                "decimal" => TabStopValues.Decimal,
                _ => TabStopValues.Left
            };
            var stop = new TabStop { Val = val, Position = pos };
            if (parts.Length >= 3)
            {
                stop.Leader = parts[2] switch
                {
                    "dot" => TabStopLeaderCharValues.Dot,
                    "hyphen" => TabStopLeaderCharValues.Hyphen,
                    "underscore" => TabStopLeaderCharValues.Underscore,
                    "middleDot" => TabStopLeaderCharValues.MiddleDot,
                    "heavy" => TabStopLeaderCharValues.Heavy,
                    _ => TabStopLeaderCharValues.None
                };
            }
            tabs.Append(stop);
        }
        return tabs.HasChildren ? tabs : null;
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
            if (unit == "px") val = OoxmlUnits.PixelsToPoints(val);
            spacing.Before = ((int)Math.Round(OoxmlUnits.PointsToTwips(val))).ToString();
            hasSpacing = true;
        }

        var marginBottomMatch = Regex.Match(style, @"margin-bottom:\s*([\d.,]+)(px|pt)");
        if (marginBottomMatch.Success)
        {
            var val = double.Parse(marginBottomMatch.Groups[1].Value.Replace(',', '.'), inv);
            var unit = marginBottomMatch.Groups[2].Value;
            if (unit == "px") val = OoxmlUnits.PixelsToPoints(val);
            spacing.After = ((int)Math.Round(OoxmlUnits.PointsToTwips(val))).ToString();
            hasSpacing = true;
        }

        var lineHeightMatch = Regex.Match(style, @"line-height:\s*([\d.,]+)(pt)?");
        if (lineHeightMatch.Success)
        {
            var val = double.Parse(lineHeightMatch.Groups[1].Value.Replace(',', '.'), inv);
            var unit = lineHeightMatch.Groups[2].Value;
            if (unit == "pt")
            {
                // Dokładna wartość w pt. Reader oznacza regułę atLeast markerem
                // --w-line-rule:atLeast — bez niego atLeast wracało jako exact,
                // a exact przycina w Wordzie tekst wyższy niż linia.
                spacing.Line = ((int)Math.Round(OoxmlUnits.PointsToTwips(val))).ToString();
                spacing.LineRule = Regex.IsMatch(style, @"--w-line-rule\s*:\s*atLeast")
                    ? LineSpacingRuleValues.AtLeast
                    : LineSpacingRuleValues.Exact;
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
                border.Size = CssPxToBorderEighthPoints(match.Groups[1].Value);
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
                "px" => OoxmlUnits.PixelsToPoints(size),
                "em" => size * 11, // Assume base 11pt
                "rem" => size * 11,
                _ => size // pt
            };

            var halfPoints = ((int)OoxmlUnits.PointsToHalfPoints(ptSize)).ToString();
            if (!props.Elements<FontSize>().Any())
                props.Append(new FontSize { Val = halfPoints });
        }
        // font-size: smaller/larger
        else if (style.Contains("font-size:smaller") || style.Contains("font-size: smaller"))
        {
            if (!props.Elements<FontSize>().Any())
                props.Append(new FontSize { Val = "18" }); // ~9pt
        }

        // Font-family. The first family wins (the rest is the generic fallback list, e.g.
        // 'Times New Roman',serif). Multi-word names arrive quoted, and crucially the browser
        // serialises innerHTML with inner double quotes as the ENTITY &quot; — which HtmlAgilityPack
        // does NOT decode. We must HTML-decode the style FIRST: otherwise the `;` inside `&quot;`
        // is mistaken for a CSS declaration separator and the name is truncated (e.g. "&quot"),
        // leaking into w:rFonts so Word silently reverts to its default font.
        var decodedStyle = System.Net.WebUtility.HtmlDecode(style);
        var fontFamilyMatch = Regex.Match(decodedStyle, @"font-family:\s*([^,;]+)");
        if (fontFamilyMatch.Success)
        {
            var fontName = fontFamilyMatch.Groups[1].Value.Trim().Trim('"', '\'').Trim();
            if (fontName.Length > 0 && !props.Elements<RunFonts>().Any())
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
            if (lsUnit == "px") ls = OoxmlUnits.PixelsToPoints(ls);
            props.Append(new Spacing { Val = (int)OoxmlUnits.PointsToTwips(ls) });
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

    /// <summary>
    /// Recognises a manual page break in any of the representations the pipeline may produce:
    /// the project marker <c>class="page-break"</c> (reader output + insertPageBreak), an explicit
    /// <c>data-docx-break="page"</c>, or a CSS <c>page-break-before</c>/<c>break-before: page</c>.
    /// Used so the break survives DOCX → HTML → DOCX (R-15). Natural Word pagination is NOT a
    /// page break and never matches here.
    /// </summary>
    private static bool IsPageBreakNode(HtmlNode node)
    {
        if (node.NodeType != HtmlNodeType.Element) return false;
        if (node.HasClass("page-break")) return true;
        if (string.Equals(node.GetAttributeValue("data-docx-break", ""), "page", StringComparison.OrdinalIgnoreCase))
            return true;
        var style = node.GetAttributeValue("style", "");
        return Regex.IsMatch(style, @"(page-break-before|break-before)\s*:\s*(always|page)", RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// Marker końca sekcji emitowany przez DocxToHtmlConverter (R-10): niewidoczny
    /// <c>div.docx-section-break</c> z geometrią NASTĘPNEJ sekcji w data-*.
    /// </summary>
    private static bool IsSectionBreakNode(HtmlNode node) =>
        node.NodeType == HtmlNodeType.Element && node.HasClass("docx-section-break");

    /// <summary>
    /// Czy następny znaczący sąsiad (pomijając komentarze i białe znaki) to marker sekcji.
    /// Reader emituje parę <c>div.page-break</c> + <c>div.docx-section-break</c> dla przerwy
    /// nextPage — page-break jest wtedy tylko wizualny i nie może stać się w:br type=page.
    /// </summary>
    private static bool NextElementSiblingIsSectionBreak(HtmlNode node)
    {
        var next = node.NextSibling;
        while (next != null && (next.NodeType == HtmlNodeType.Comment ||
               (next.NodeType == HtmlNodeType.Text && string.IsNullOrWhiteSpace(next.InnerText))))
        {
            next = next.NextSibling;
        }
        return next != null && IsSectionBreakNode(next);
    }

    /// <summary>
    /// Zamyka bieżącą sekcję: paragraf z pPr/sectPr niosącym geometrię sekcji ZAMYKANEJ
    /// (tak koduje to OOXML), po czym otwiera następną sekcję geometrią z data-* markera.
    /// Pierwszy taki sectPr zapamiętujemy — do niego pójdą referencje nagłówka/stopki
    /// (kolejne sekcje dziedziczą je w Wordzie, gdy nie deklarują własnych).
    /// </summary>
    private Paragraph CreateSectionBreakParagraph(HtmlNode node)
    {
        var closingSection = new SectionProperties();
        AppendSectionBreakType(closingSection, _currentSection.BreakType);
        AppendSectionGeometry(closingSection, _currentSection);

        _firstSectionProps ??= closingSection;
        _emittedSectionProps.Add(closingSection);
        _hasSectionMarkers = true;
        _currentSection = ReadSectionGeometryFromMarker(node, _currentSection);

        return new Paragraph(new ParagraphProperties(closingSection));
    }

    /// <summary>
    /// Geometria sekcji z data-* markera; brakujące wartości dziedziczą z poprzedniej
    /// sekcji (przybliżenie dziedziczenia sekcji Worda).
    /// </summary>
    private static SectionGeometry ReadSectionGeometryFromMarker(HtmlNode node, SectionGeometry previous)
    {
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        double? Attr(string name)
        {
            var raw = node.GetAttributeValue(name, string.Empty);
            return !string.IsNullOrEmpty(raw) &&
                   double.TryParse(raw, System.Globalization.NumberStyles.Float, inv, out var v)
                ? v : null;
        }

        Domain.Models.PageSize? pageSize = null;
        if (Attr("data-page-width-cm") is { } w && Attr("data-page-height-cm") is { } h)
        {
            pageSize = new Domain.Models.PageSize
            {
                WidthCm = w,
                HeightCm = h,
                Orientation = node.GetAttributeValue("data-orientation", "portrait")
            };
        }

        PageMargins? margins = null;
        var top = Attr("data-margin-top-cm");
        var bottom = Attr("data-margin-bottom-cm");
        var left = Attr("data-margin-left-cm");
        var right = Attr("data-margin-right-cm");
        if (top != null || bottom != null || left != null || right != null)
        {
            margins = new PageMargins
            {
                Top = top ?? previous.Margins?.Top ?? 2.5,
                Bottom = bottom ?? previous.Margins?.Bottom ?? 2.5,
                Left = left ?? previous.Margins?.Left ?? 2.5,
                Right = right ?? previous.Margins?.Right ?? 2.5
            };
        }

        return new SectionGeometry
        {
            PageSize = pageSize ?? previous.PageSize,
            Margins = margins ?? previous.Margins,
            HeaderDistanceCm = Attr("data-header-distance-cm"),
            FooterDistanceCm = Attr("data-footer-distance-cm"),
            BreakType = node.GetAttributeValue("data-break-type", "nextPage")
        };
    }

    /// <summary>
    /// w:type sekcji — zapisywany tylko, gdy różni się od domyślnego nextPage.
    /// Musi poprzedzać pgSz/pgMar (kolejność schematu sectPr).
    /// </summary>
    private static void AppendSectionBreakType(SectionProperties sectionProps, string? breakType)
    {
        SectionMarkValues? val = breakType switch
        {
            "continuous" => SectionMarkValues.Continuous,
            "oddPage" => SectionMarkValues.OddPage,
            "evenPage" => SectionMarkValues.EvenPage,
            "nextColumn" => SectionMarkValues.NextColumn,
            _ => null
        };
        if (val is { } v && !sectionProps.Elements<SectionType>().Any())
            sectionProps.Append(new SectionType { Val = v });
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

        extPropsPart.Properties.Application = new DocumentFormat.OpenXml.ExtendedProperties.Application("Qutas D2Tools");
        extPropsPart.Properties.Save();
    }

    /// <summary>
    /// Dodaje ustawienia strony z dokładnymi marginesami
    /// </summary>
    private void AddPageSettings(Body body, HeaderFooterContent? header = null, HeaderFooterContent? footer = null, PageMargins? margins = null, Domain.Models.PageSize? pageSize = null)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }

        // Body-level sectPr opisuje OSTATNIĄ sekcję. Gdy body zawierało markery sekcji,
        // jej geometria pochodzi z ostatniego markera (_currentSection); bez markerów —
        // z argumentów (dotychczasowe zachowanie, jedna sekcja).
        var geometry = _hasSectionMarkers
            ? _currentSection
            : new SectionGeometry { PageSize = pageSize, Margins = margins };

        if (_hasSectionMarkers)
            AppendSectionBreakType(sectionProps, geometry.BreakType);
        AppendSectionGeometry(sectionProps, geometry);
    }

    /// <summary>
    /// Dopisuje w:pgSz + w:pgMar do sectPr wg geometrii sekcji (jedno źródło reguł dla
    /// body-level i paragraph-level sectPr).
    ///
    /// Margins (cm) → twips via the central converter (the exact factor is 1440/2.54).
    /// Defaults are 1 inch sides, 0.5 inch header/footer bands. Body margins are written
    /// AS AUTHORED — they must not be inflated. The reader derives the header/footer band
    /// height as (margin − distance); here we invert that to recover the original
    /// w:header / w:footer distance = (margin − band), clamped to [0, 720] — unless the
    /// section marker carried the authored distances (data-header/footer-distance-cm),
    /// which round-trip verbatim. The previous Math.Max(top, headerHeight + 720) +
    /// hardcoded Header/Footer=720 pushed a small-margin / small-header document
    /// (top=567, header=6) to top=1281, adding pages. Patrz analiza orginał_GOOD vs zapisany_BAD.
    /// </summary>
    private void AppendSectionGeometry(SectionProperties sectionProps, SectionGeometry geometry)
    {
        // Round-trip the authored page size/orientation; fall back to A4 portrait when unknown.
        if (!sectionProps.Elements<OoxmlPageSize>().Any())
        {
            sectionProps.Append(BuildPageSize(geometry.PageSize));
        }

        int defaultMarginTwips = OoxmlUnits.TwipsPerInch;
        int defaultBandTwips = OoxmlUnits.TwipsPerInch / 2;
        var margins = geometry.Margins;
        int leftTwips  = margins != null ? (int)Math.Round(OoxmlUnits.CmToTwips(margins.Left))  : defaultMarginTwips;
        int rightTwips = margins != null ? (int)Math.Round(OoxmlUnits.CmToTwips(margins.Right)) : defaultMarginTwips;

        var headerHeightTwips = _headerBandCm is { } hb ? (int)OoxmlUnits.CmToTwips(hb) : defaultBandTwips;
        var footerHeightTwips = _footerBandCm is { } fb ? (int)OoxmlUnits.CmToTwips(fb) : defaultBandTwips;

        int topTwips    = margins != null ? (int)Math.Round(OoxmlUnits.CmToTwips(margins.Top))    : defaultMarginTwips;
        int bottomTwips = margins != null ? (int)Math.Round(OoxmlUnits.CmToTwips(margins.Bottom)) : defaultMarginTwips;

        const int maxBandDistanceTwips = 720; // 0.5"
        uint headerDistance = geometry.HeaderDistanceCm is { } hd
            ? (uint)Math.Max(0, (int)Math.Round(OoxmlUnits.CmToTwips(hd)))
            : (uint)Math.Clamp(topTwips - headerHeightTwips, 0, maxBandDistanceTwips);
        uint footerDistance = geometry.FooterDistanceCm is { } fd
            ? (uint)Math.Max(0, (int)Math.Round(OoxmlUnits.CmToTwips(fd)))
            : (uint)Math.Clamp(bottomTwips - footerHeightTwips, 0, maxBandDistanceTwips);

        if (!sectionProps.Elements<PageMargin>().Any())
        {
            sectionProps.Append(new PageMargin
            {
                Top    = topTwips,
                Right  = (uint)rightTwips,
                Bottom = bottomTwips,
                Left   = (uint)leftTwips,
                Header = headerDistance,
                Footer = footerDistance
            });
        }
    }

    private static OoxmlPageSize BuildPageSize(Domain.Models.PageSize? pageSize)
    {
        // A4 portrait default (twips) matches Word's default new-document section.
        const int a4WidthTwips = 11906;
        const int a4HeightTwips = 16838;

        if (pageSize == null || pageSize.WidthCm <= 0 || pageSize.HeightCm <= 0)
            return new OoxmlPageSize { Width = a4WidthTwips, Height = a4HeightTwips };

        var result = new OoxmlPageSize
        {
            Width = (uint)Math.Round(OoxmlUnits.CmToTwips(pageSize.WidthCm)),
            Height = (uint)Math.Round(OoxmlUnits.CmToTwips(pageSize.HeightCm))
        };
        if (string.Equals(pageSize.Orientation, "landscape", StringComparison.OrdinalIgnoreCase))
            result.Orient = PageOrientationValues.Landscape;
        return result;
    }

    private static DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties BuildImageDocProperties(uint id, string name, string? altText)
    {
        var props = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = id, Name = name };
        if (!string.IsNullOrWhiteSpace(altText))
            props.Description = altText;
        return props;
    }

    /// <summary>
    /// Builds a PAGE/NUMPAGES field run, carrying the field span's font (font-size etc.) so the
    /// page number keeps the footer's size on round-trip instead of falling back to a default.
    /// Falls back to the parent span's style for the wrapped {page} case.
    /// </summary>
    private SimpleField BuildFieldRun(string instruction, HtmlNode fieldNode)
    {
        // fldSimple is paragraph-level; the run properties (font-size etc.) live on an inner run
        // so Word — and our reader on the next round-trip — keep the page number's footer font.
        var run = new Run();
        var style = fieldNode.GetAttributeValue("style", "");
        if (string.IsNullOrEmpty(style))
            style = fieldNode.ParentNode?.GetAttributeValue("style", "") ?? string.Empty;
        if (!string.IsNullOrEmpty(style))
        {
            var rPr = new RunProperties();
            ApplyRunStyle(rPr, style);
            if (rPr.HasChildren) run.Append(rPr);
        }
        run.Append(new Text(fieldNode.InnerText?.Trim() is { Length: > 0 } t ? t : "1") { Space = SpaceProcessingModeValues.Preserve });
        return new SimpleField(run) { Instruction = instruction };
    }

    private int PxToTwips(int px) => (int)OoxmlUnits.PixelsToTwips(px);

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

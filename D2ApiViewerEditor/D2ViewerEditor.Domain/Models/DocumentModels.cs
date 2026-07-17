namespace D2ViewerEditor.Domain.Models;

/// <summary>
/// Marginesy strony dokumentu (w cm)
/// </summary>
public class PageMargins
{
    public double Top { get; set; } = 2.5;
    public double Bottom { get; set; } = 2.5;
    public double Left { get; set; } = 2.5;
    public double Right { get; set; } = 2.5;
}

/// <summary>
/// Rozmiar i orientacja strony (w cm). Odwzorowuje w:pgSz z sekcji DOCX.
/// </summary>
public class PageSize
{
    public double WidthCm { get; set; }
    public double HeightCm { get; set; }
    /// <summary>"portrait" | "landscape".</summary>
    public string Orientation { get; set; } = "portrait";
}

/// <summary>
/// Układ kolumn sekcji (odwzorowuje w:cols z w:sectPr). Domyślnie jedna kolumna.
/// Szerokości/odstępy w twipach — jednostka DOCX; warstwy prezentacji przeliczają przez
/// centralny konwerter jednostek. <see cref="Columns"/> wypełnione tylko dla kolumn
/// nierównych (w:equalWidth="0" z jawnymi w:col); dla równych wystarczą Count + SpaceTwips.
/// </summary>
public class ColumnLayout
{
    public int Count { get; set; } = 1;
    public bool EqualWidth { get; set; } = true;
    /// <summary>Domyślny odstęp między kolumnami (w:space na w:cols), w twipach.</summary>
    public int SpaceTwips { get; set; } = 720;
    /// <summary>Linia separatora między kolumnami (w:sep).</summary>
    public bool Separator { get; set; }
    /// <summary>Indywidualne kolumny — tylko dla kolumn nierównych (EqualWidth=false).</summary>
    public List<SectionColumn>? Columns { get; set; }
}

/// <summary>Pojedyncza kolumna nierównego układu: szerokość i odstęp po niej (twipy).</summary>
public class SectionColumn
{
    public int WidthTwips { get; set; }
    public int SpaceTwips { get; set; }
}

/// <summary>
/// Nagłówek lub stopka dokumentu
/// </summary>
public class HeaderFooterContent
{
    /// <summary>Default (odd / primary) header or footer shown on ordinary pages.</summary>
    public string Html { get; set; } = string.Empty;
    public double Height { get; set; } = 1.25;
    public bool DifferentFirstPage { get; set; }
    public string? FirstPageHtml { get; set; }
    /// <summary>Section enables a distinct even-page header/footer (w:evenAndOddHeaders).</summary>
    public bool DifferentOddEven { get; set; }
    public string? EvenHtml { get; set; }
}

/// <summary>
/// Nagłówek/stopka JEDNEJ sekcji dokumentu wielosekcyjnego. Wpisy istnieją tylko dla
/// sekcji (indeks 0-based w kolejności dokumentu), które deklarują WŁASNE referencje
/// nagłówka/stopki; sekcja bez wpisu dziedziczy je z poprzedniej sekcji (jak Word).
/// Sekcja 0 pozostaje w polach <see cref="DocumentContent.Header"/>/<see cref="DocumentContent.Footer"/>.
/// </summary>
public class SectionHeaderFooter
{
    public int SectionIndex { get; set; }
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
}

/// <summary>
/// Przypis dolny dokumentu. Tożsamość (<see cref="Id"/>) jest STABILNA i niezależna od
/// numeru widocznego w edytorze — numer wynika z kolejności pierwszych odwołań w treści
/// i jest liczony przy renderowaniu, a nie przechowywany. Numeryczny identyfikator OOXML
/// (w:id) jest przydzielany deterministycznie dopiero podczas eksportu do DOCX. Treść
/// przypisu (<see cref="Html"/>) jest jedynym źródłem prawdy — odwołania w treści dokumentu
/// niosą tylko <c>data-footnote-id</c>, nie kopię treści.
/// </summary>
public class Footnote
{
    /// <summary>Stabilna wewnętrzna tożsamość przypisu (np. "fn-1"). Nie mylić z numerem wyświetlanym.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Treść przypisu jako HTML (akapity/formatowanie w tym samym modelu co treść dokumentu).</summary>
    public string Html { get; set; } = string.Empty;
}

/// <summary>
/// Przypis końcowy dokumentu. Semantyka tożsamości/numeru jest identyczna jak w
/// <see cref="Footnote"/> (stabilne <see cref="Id"/> "en-N" oddzielone od numeru widocznego
/// i od w:id OOXML), ale typ jest ODDZIELNY, bo endnotes mają inną semantykę renderowania
/// (koniec dokumentu, nie dół strony) i inną część OOXML (word/endnotes.xml). Odwołania w
/// treści niosą tylko <c>data-endnote-id</c>.
/// </summary>
public class Endnote
{
    /// <summary>Stabilna wewnętrzna tożsamość przypisu końcowego (np. "en-1"). Nie mylić z numerem wyświetlanym.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Treść przypisu końcowego jako HTML (ten sam model treści co dokument).</summary>
    public string Html { get; set; } = string.Empty;
}

/// <summary>
/// Reprezentuje dokument z konwersji DOCX do HTML
/// </summary>
public class DocumentContent
{
    public string Html { get; set; } = string.Empty;
    public DocumentMetadata Metadata { get; set; } = new();
    public List<DocumentImage> Images { get; set; } = new();
    public List<DocumentStyle> Styles { get; set; } = new();
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
    public PageMargins? Margins { get; set; }
    public PageSize? PageSize { get; set; }
    /// <summary>Układ kolumn sekcji bazowej (0). Null/1 kolumna = układ jednokolumnowy (ADR-0039).</summary>
    public ColumnLayout? Columns { get; set; }
    /// <summary>Własne nagłówki/stopki sekcji ≥ 1 (dokumenty wielosekcyjne, R-10/ADR-0023).</summary>
    public List<SectionHeaderFooter>? SectionHeadersFooters { get; set; }
    /// <summary>
    /// Przypisy dolne w kolejności pierwszych odwołań w treści. Puste/null dla dokumentów
    /// bez przypisów (nie tworzą wtedy nadmiarowej części footnotes.xml na eksporcie).
    /// </summary>
    public List<Footnote>? Footnotes { get; set; }

    /// <summary>
    /// Przypisy końcowe w kolejności pierwszych odwołań w treści. Puste/null dla dokumentów
    /// bez przypisów końcowych (nie tworzą wtedy nadmiarowej części endnotes.xml na eksporcie).
    /// </summary>
    public List<Endnote>? Endnotes { get; set; }

    /// <summary>
    /// Format numeracji przypisów DOLNYCH z <c>w:footnotePr/w:numFmt</c> (settings.xml, document-wide):
    /// token <c>decimal</c>/<c>lowerRoman</c>/<c>upperRoman</c>/<c>lowerLetter</c>/<c>upperLetter</c>.
    /// <c>null</c> = dokument nie ustala formatu → GUI używa domyślnej Worda (dolne = cyfry). Nie
    /// round-tripuje przez zapis (żyje w zachowanym settings.xml pakietu — tylko odczyt/wyświetlanie).
    /// </summary>
    public string? FootnoteNumberFormat { get; set; }

    /// <summary>
    /// Format numeracji przypisów KOŃCOWYCH z <c>w:endnotePr/w:numFmt</c> (settings.xml, document-wide).
    /// Token jak w <see cref="FootnoteNumberFormat"/>; <c>null</c> = domyślna Worda (końcowe = małe rzymskie).
    /// </summary>
    public string? EndnoteNumberFormat { get; set; }

    /// <summary>
    /// True, gdy dokument źródłowy deklaruje ochronę przed edycją w settings.xml:
    /// wymuszone w:documentProtection (Ogranicz edycję, tryb inny niż "none") lub
    /// w:writeProtection (hasło zapisu / zalecenie tylko-do-odczytu). Edytor musi wtedy
    /// otworzyć dokument w trybie tylko do odczytu — nie umiemy egzekwować trybów
    /// częściowych (komentarze/formularze), więc każda wymuszona ochrona blokuje edycję.
    /// </summary>
    public bool IsReadOnlyProtected { get; set; }
}

/// <summary>
/// Metadane dokumentu DOCX
/// </summary>
public class DocumentMetadata
{
    // Core Properties
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public string? Keywords { get; set; }
    public string? Description { get; set; }
    public string? Category { get; set; }
    public string? ContentStatus { get; set; }
    public string? LastModifiedBy { get; set; }
    public string? Revision { get; set; }
    public string? Version { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public int PageCount { get; set; }
    public int WordCount { get; set; }

    // Extended Properties
    public string? Company { get; set; }
    public string? Manager { get; set; }

    // Podpisy cyfrowe
    public List<DigitalSignatureInfo>? Signatures { get; set; }
}

/// <summary>
/// Informacja o podpisie cyfrowym
/// </summary>
public class DigitalSignatureInfo
{
    public string SignerName { get; set; } = string.Empty;
    public string? SignerEmail { get; set; }
    public string? SignerTitle { get; set; }
    public string CertificateSubject { get; set; } = string.Empty;
    public string CertificateIssuer { get; set; } = string.Empty;
    public string CertificateSerialNumber { get; set; } = string.Empty;
    public DateTime SignedAt { get; set; }
    public DateTime CertificateValidFrom { get; set; }
    public DateTime CertificateValidTo { get; set; }
    public bool IsValid { get; set; }
    public string? ValidationMessage { get; set; }
    public string? Reason { get; set; }
}

/// <summary>
/// Request podpisania dokumentu
/// </summary>
public class SignDocumentRequest
{
    public string Html { get; set; } = string.Empty;
    public string? OriginalFileName { get; set; }
    public DocumentMetadata? Metadata { get; set; }
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
    public string CertificateBase64 { get; set; } = string.Empty;
    public string CertificatePassword { get; set; } = string.Empty;
    public string SignerName { get; set; } = string.Empty;
    public string? SignerTitle { get; set; }
    public string? SignerEmail { get; set; }
    public string? SignatureReason { get; set; }
    public PageMargins? Margins { get; set; }
    public PageSize? PageSize { get; set; }
}

/// <summary>
/// Obraz osadzony w dokumencie
/// </summary>
public class DocumentImage
{
    public string Id { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public string Base64Data { get; set; } = string.Empty;
}

/// <summary>
/// Request do zapisu dokumentu
/// </summary>
public class SaveDocumentRequest
{
    public string Html { get; set; } = string.Empty;
    public string? OriginalFileName { get; set; }
    public DocumentMetadata? Metadata { get; set; }
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
    public PageMargins? Margins { get; set; }
    public PageSize? PageSize { get; set; }
    public List<SectionHeaderFooter>? SectionHeadersFooters { get; set; }
    /// <summary>Przypisy dolne (jedno źródło prawdy dla treści; odwołania w Html niosą tylko id).</summary>
    public List<Footnote>? Footnotes { get; set; }
    /// <summary>Przypisy końcowe (jedno źródło prawdy dla treści; odwołania w Html niosą tylko id).</summary>
    public List<Endnote>? Endnotes { get; set; }
    /// <summary>Opcjonalne: gdy podane, zapis idzie przez pass-through oryginalnego pakietu
    /// (zachowuje styles.xml/theme/fontTable/numbering). Brak → pełna regeneracja jak dotąd.</summary>
    public Guid? MasterId { get; set; }
}

/// <summary>
/// Styl paragrafu
/// </summary>
public class ParagraphStyle
{
    public string? FontFamily { get; set; }
    public double? FontSize { get; set; }
    public string? Color { get; set; }
    public string? BackgroundColor { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
    public bool IsStrikethrough { get; set; }
    public string? Alignment { get; set; }
    public double? LineSpacing { get; set; }
    public double? SpaceBefore { get; set; }
    public double? SpaceAfter { get; set; }
    public double? LeftIndent { get; set; }
    public double? RightIndent { get; set; }
    public double? FirstLineIndent { get; set; }
}

/// <summary>
/// Styl dokumentu (Nagłówek 1, Normalny, itp.)
/// </summary>
public class DocumentStyle
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = "paragraph";
    public string? BasedOn { get; set; }
    public string? NextStyle { get; set; }

    public string? FontFamily { get; set; }
    public double? FontSize { get; set; }
    public string? Color { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }

    public string? Alignment { get; set; }
    public double? SpaceBefore { get; set; }
    public double? SpaceAfter { get; set; }
    public double? LineSpacing { get; set; }
    public double? LeftIndent { get; set; }
    public double? RightIndent { get; set; }
    public double? FirstLineIndent { get; set; }

    public int? OutlineLevel { get; set; }
}

/// <summary>
/// Domyślne style Word
/// </summary>
public static class DefaultWordStyles
{
    public static List<DocumentStyle> GetDefaultStyles()
    {
        return new List<DocumentStyle>
        {
            new()
            {
                Id = "Normal", Name = "Normalny", Type = "paragraph",
                FontFamily = "Calibri", FontSize = 11, Color = "#000000",
                Alignment = "left", SpaceAfter = 8, LineSpacing = 1.08
            },
            new()
            {
                Id = "Heading1", Name = "Nagłówek 1", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 16, Color = "#2F5496",
                IsBold = true, SpaceBefore = 12, SpaceAfter = 0, OutlineLevel = 1
            },
            new()
            {
                Id = "Heading2", Name = "Nagłówek 2", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 13, Color = "#2F5496",
                IsBold = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 2
            },
            new()
            {
                Id = "Heading3", Name = "Nagłówek 3", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 12, Color = "#1F3763",
                IsBold = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 3
            },
            new()
            {
                Id = "Heading4", Name = "Nagłówek 4", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#2F5496",
                IsBold = true, IsItalic = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 4
            },
            new()
            {
                Id = "Heading5", Name = "Nagłówek 5", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#2F5496",
                SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 5
            },
            new()
            {
                Id = "Heading6", Name = "Nagłówek 6", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#1F3763",
                IsItalic = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 6
            },
            new()
            {
                Id = "Title", Name = "Tytuł", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 28, Color = "#000000",
                SpaceAfter = 0, LineSpacing = 1.0
            },
            new()
            {
                Id = "Subtitle", Name = "Podtytuł", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri", FontSize = 11, Color = "#5A5A5A",
                IsItalic = true, SpaceAfter = 8
            },
            new()
            {
                Id = "Quote", Name = "Cytat", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri", FontSize = 11, Color = "#404040",
                IsItalic = true, LeftIndent = 1.27, RightIndent = 1.27,
                SpaceBefore = 10, SpaceAfter = 10
            },
            new()
            {
                Id = "ListParagraph", Name = "Akapit listy", Type = "paragraph",
                BasedOn = "Normal", LeftIndent = 1.27
            }
        };
    }
}

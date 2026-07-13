using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Kontrakt konwertera HTML → DOCX
/// </summary>
public interface IHtmlToDocxConverter
{
    /// <summary>
    /// Konwertuje HTML na dokument DOCX (tablica bajtów)
    /// </summary>
    byte[] Convert(string html, DocumentMetadata? metadata = null,
                   HeaderFooterContent? header = null, HeaderFooterContent? footer = null,
                   PageMargins? margins = null, PageSize? pageSize = null,
                   IReadOnlyList<SectionHeaderFooter>? sectionHeadersFooters = null,
                   IReadOnlyList<Footnote>? footnotes = null);

    /// <summary>
    /// Pass-through wariant: konwertuje HTML jak <see cref="Convert"/>, a następnie zachowuje
    /// z oryginalnego pakietu DOCX te party, których edytor nie modyfikuje, a których
    /// regeneracja od zera obniża wierność: <c>styles.xml</c> (pełny zestaw stylów, w tym tabel),
    /// <c>theme</c> i <c>fontTable</c>. Dzięki temu nie giną style, a font treści (np. Cambria
    /// z theme minor) nie jest podmieniany na domyślny. Body, sekcja, nagłówki/stopki, obrazy
    /// i numbering pochodzą z konwersji HTML. Gdy <paramref name="originalPackage"/> jest null/pusty,
    /// zachowuje się jak zwykłe <see cref="Convert"/> (np. dokumenty bez oryginału).
    /// </summary>
    byte[] ConvertPreservingPackage(string html, Stream? originalPackage,
                   DocumentMetadata? metadata = null,
                   HeaderFooterContent? header = null, HeaderFooterContent? footer = null,
                   PageMargins? margins = null, PageSize? pageSize = null,
                   IReadOnlyList<SectionHeaderFooter>? sectionHeadersFooters = null,
                   IReadOnlyList<Footnote>? footnotes = null);
}

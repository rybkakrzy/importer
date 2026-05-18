namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Konfiguracja domyślnych ustawień dokumentu używanych podczas konwersji
/// DOCX↔HTML, gdy konkretny element / run nie ma własnej deklaracji.
///
/// Sekcja w appsettings.json: <c>DocumentDefaults</c>.
/// Pole <see cref="FontFamily"/> ustawia firmowy krój tekstu w treści.
/// Pole <see cref="HeadingFontFamily"/> ustawia firmowy krój nagłówków (H1–H6).
/// Pole <see cref="FontSizePt"/> ustawia bazowy rozmiar tekstu (w punktach).
/// </summary>
public sealed class DocumentDefaultsOptions
{
    /// <summary>Nazwa sekcji konfiguracji.</summary>
    public const string SectionName = "DocumentDefaults";

    /// <summary>
    /// Domyślny krój tekstu w treści (akapity, listy, tabele).
    /// Wartość trafia do <c>w:rFonts</c> w <c>docDefaults</c> oraz w stylu „Normal”
    /// generowanego DOCX-a, a po stronie HTML do <c>font-family</c> kontenera.
    /// </summary>
    public string FontFamily { get; set; } = "Calibri";

    /// <summary>
    /// Domyślny krój nagłówków (H1–H6). Jeśli puste — używany jest
    /// <see cref="FontFamily"/>.
    /// </summary>
    public string HeadingFontFamily { get; set; } = "Calibri Light";

    /// <summary>
    /// Bazowy rozmiar tekstu w punktach (np. 11 dla 11pt).
    /// Trafia do <c>w:sz</c> jako <c>FontSizePt * 2</c>.
    /// </summary>
    public double FontSizePt { get; set; } = 11.0;
}

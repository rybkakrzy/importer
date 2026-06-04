using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Konwerter grafik dokumentowych do formatu wyświetlanego w przeglądarce. Obejmuje legacy
/// grafiki Worda (EMF/WMF metafile, VML) oraz formaty web-native (PNG/JPEG/GIF/BMP/SVG).
///
/// Zasady:
///  - Pure-managed, bez LibreOffice/soffice, bez System.Drawing/GDI — działa w kontenerze na
///    Linux/GCP, deterministycznie, bez zależności natywnych.
///  - EMF/WMF nie są rasteryzowane (brak bezpiecznej, wieloplatformowej ścieżki bez GDI): gdy
///    metafile zawiera osadzony raster, wyciągamy go; w przeciwnym razie zwracamy kontrolowany
///    placeholder SVG, a oryginalny part jest zachowywany do DOCX (pass-through), więc Word
///    nadal renderuje prawdziwą grafikę i dokument nie zostaje uszkodzony.
///  - Wejście jest niezaufane: limity rozmiaru/czasu, twardy XML/SVG sanitizer.
/// </summary>
public interface IGraphicConversionService
{
    /// <summary>Wykrywa typ grafiki z magic bytes (z opcjonalną podpowiedzią content-type).</summary>
    GraphicKind Detect(ReadOnlySpan<byte> data, string? contentType = null);

    /// <summary>
    /// Konwertuje grafikę (media part: PNG/JPEG/GIF/BMP/SVG/EMF/WMF) do reprezentacji webowej.
    /// Nigdy nie rzuca dla danych niezaufanych — błędy mapuje na Fallback/Rejected w diagnostyce.
    /// </summary>
    GraphicConversionResult ConvertForEditor(
        GraphicSource source,
        GraphicConversionOptions? options = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Konwertuje VML shape (w:pict/v:shape/v:rect/v:oval/v:line) do SVG dla bezpiecznego
    /// podzbioru kształtów. Zwraca null, gdy kształt jest nieobsługiwany (caller użyje placeholdera
    /// lub rozwiąże v:imagedata przez <see cref="ConvertForEditor"/>).
    /// </summary>
    GraphicConversionResult? ConvertVmlShapeForEditor(
        string vmlXml,
        GraphicConversionOptions? options = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Sanitizuje SVG do bezpiecznego podzbioru: usuwa script/foreignObject, atrybuty zdarzeń
    /// (on*), zewnętrzne referencje (href/xlink:href poza data:), encje/DTD. Zwraca null, gdy
    /// SVG jest nieparsowalny albo pusty po sanitizacji.
    /// </summary>
    string? SanitizeSvg(string svgXml);
}

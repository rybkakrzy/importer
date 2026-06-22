namespace D2ViewerEditor.Domain.Models;

/// <summary>
/// Wykryty typ grafiki źródłowej. Web-native (Png/Jpeg/Gif/Bmp/Webp/Ico/Svg) renderuje się w
/// przeglądarce bez konwersji; Tiff/Emf/Wmf wymagają konwersji (nie renderowane natywnie); Vml to
/// wektor XML.
/// </summary>
public enum GraphicKind
{
    Unknown = 0,
    Png,
    Jpeg,
    Gif,
    Bmp,
    Webp,
    Ico,
    Svg,
    Tiff,
    Emf,
    Wmf,
    Vml
}

/// <summary>Skąd pochodzi grafika — wpływa na strategię zapisu zwrotnego (pass-through vs media part).</summary>
public enum GraphicOrigin
{
    /// <summary>Oryginalny part z importowanego DOCX (kandydat do pass-through gdy nieedytowany).</summary>
    LegacyDocxPart = 0,
    /// <summary>Grafika wstawiona/edytowana w edytorze (zapisywana jako nowoczesny media part).</summary>
    EditorInserted = 1
}

/// <summary>Wynik konwersji z punktu widzenia integralności dokumentu.</summary>
public enum GraphicConversionStatus
{
    /// <summary>Skonwertowano do formatu webowego.</summary>
    Converted = 0,
    /// <summary>Format już webowy — przekazany bez zmian.</summary>
    PassThrough,
    /// <summary>
    /// Nie udało się wyprodukować rastra do podglądu — zwrócono PRZEZROCZYSTĄ, niewidoczną grafikę
    /// (brak fałszywego placeholdera). Oryginalny metafile jedzie do DOCX (Word renderuje wektor).
    /// </summary>
    Fallback,
    /// <summary>Format rozpoznany, ale nieobsługiwany — przezroczysta grafika + diagnostyka (bez fałszywego placeholdera).</summary>
    Unsupported,
    /// <summary>Odrzucono (za duże / niebezpieczne / uszkodzone).</summary>
    Rejected
}

/// <summary>Jak wierne jest odwzorowanie względem oryginału.</summary>
public enum GraphicFidelity
{
    Lossless = 0,
    Lossy,
    Fallback,
    Unsupported
}

/// <summary>
/// Niezaufane wejście do konwersji. <see cref="Data"/> to surowe bajty media partu (a:blip /
/// v:imagedata) albo VML XML zakodowany w UTF-8.
/// </summary>
public sealed class GraphicSource
{
    public required byte[] Data { get; init; }
    public string? ContentType { get; init; }
    public string? FileName { get; init; }
    /// <summary>Ścieżka/nazwa partu źródłowego (np. /word/media/image1.emf) — do raportowania błędów.</summary>
    public string? SourcePath { get; init; }
    public GraphicOrigin Origin { get; init; } = GraphicOrigin.LegacyDocxPart;
    /// <summary>Rozmiar docelowy z DOCX (EMU) — używany gdy nagłówek grafiki nie ma wymiarów.</summary>
    public long? TargetWidthEmu { get; init; }
    public long? TargetHeightEmu { get; init; }
}

/// <summary>Opcje/limity bezpieczeństwa konwersji.</summary>
public sealed class GraphicConversionOptions
{
    public int MaxInputBytes { get; init; } = 32 * 1024 * 1024;   // 32 MB
    public int MaxPlaceholderWidthPx { get; init; } = 2000;
    public int MaxPlaceholderHeightPx { get; init; } = 2000;
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(5);

    public static GraphicConversionOptions Default { get; } = new();
}

/// <summary>Reprezentacja gotowa do wyświetlenia w przeglądarce (data URL bezpieczny do osadzenia).</summary>
public sealed class WebGraphicRepresentation
{
    public required string MimeType { get; init; }
    public required byte[] Data { get; init; }
    public int WidthPx { get; init; }
    public int HeightPx { get; init; }

    /// <summary>
    /// True, gdy <see cref="Data"/> to PRZEZROCZYSTA grafika zastępcza (brak realnego rastra do
    /// rasteryzacji metafile) — nie jest to widoczny placeholder, tylko niewidzialny element
    /// zachowujący układ. Realna treść wektorowa jedzie do DOCX przez pass-through (Word renderuje).
    /// </summary>
    public bool IsBlankFallback { get; init; }

    /// <summary>data:URL gotowy do `src` w &lt;img&gt;. SVG kodowany base64 (bezpieczne dla cudzysłowów).</summary>
    public string ToDataUrl() =>
        $"data:{MimeType};base64,{Convert.ToBase64String(Data)}";
}

/// <summary>
/// Strukturalna diagnostyka jednej konwersji — do raportowania wewnętrznego (nie pokazywana
/// użytkownikowi). Zawiera ścieżkę partu, status, fidelity, czas, klucz cache, listę prób strategii,
/// powód niepowodzenia, ostrzeżenia i utracone właściwości.
/// </summary>
public sealed class GraphicConversionDiagnostics
{
    public GraphicKind InputKind { get; init; }
    public string? SourcePath { get; init; }
    public string OutputMimeType { get; init; } = string.Empty;
    public GraphicConversionStatus Status { get; init; }
    public GraphicFidelity Fidelity { get; init; }
    public long ElapsedMs { get; init; }
    public string CacheKey { get; init; } = string.Empty;
    /// <summary>Strategie konwersji wypróbowane w kolejności (np. embedded-raster, dib-rasterize).</summary>
    public IReadOnlyList<string> AttemptedStrategies { get; init; } = Array.Empty<string>();
    /// <summary>Powód niepowodzenia/fallbacku (null, gdy konwersja w pełni się powiodła).</summary>
    public string? FailureReason { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> LostProperties { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Wynik konwersji. <see cref="Web"/> = co pokazać w edytorze (null tylko gdy Rejected).
/// <see cref="PreserveOriginalPart"/> = czy zapis zwrotny ma zachować oryginalny media part
/// (pass-through) zamiast re-enkodować — true dla legacy EMF/WMF/TIFF/VML.
/// </summary>
public sealed class GraphicConversionResult
{
    public WebGraphicRepresentation? Web { get; init; }
    public bool PreserveOriginalPart { get; init; }
    public required GraphicConversionDiagnostics Diagnostics { get; init; }

    public bool HasWebRepresentation => Web != null;
}

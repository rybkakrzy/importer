namespace D2ViewerEditor.Domain.Models;

/// <summary>
/// Wykryty typ grafiki źródłowej. Web-native (Png..Svg) renderuje się w przeglądarce bez
/// konwersji; Emf/Wmf to legacy metafile Worda (nie renderowane natywnie), Vml to wektor XML.
/// </summary>
public enum GraphicKind
{
    Unknown = 0,
    Png,
    Jpeg,
    Gif,
    Bmp,
    Svg,
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
    /// <summary>Nie dało się odwzorować — zwrócono kontrolowany placeholder (oryginał zachowany do DOCX).</summary>
    Fallback,
    /// <summary>Format rozpoznany, ale nieobsługiwany — placeholder + diagnostyka.</summary>
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
    public GraphicOrigin Origin { get; init; } = GraphicOrigin.LegacyDocxPart;
    /// <summary>Rozmiar docelowy z DOCX (EMU) — używany do placeholdera, gdy header nie ma wymiarów.</summary>
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
    public bool IsPlaceholder { get; init; }

    /// <summary>data:URL gotowy do `src` w &lt;img&gt;. SVG kodowany base64 (bezpieczne dla cudzysłowów).</summary>
    public string ToDataUrl() =>
        $"data:{MimeType};base64,{Convert.ToBase64String(Data)}";
}

/// <summary>Diagnostyka jednej konwersji — status, fidelity, czas, ostrzeżenia, utracone właściwości.</summary>
public sealed class GraphicConversionDiagnostics
{
    public GraphicKind InputKind { get; init; }
    public string OutputMimeType { get; init; } = string.Empty;
    public GraphicConversionStatus Status { get; init; }
    public GraphicFidelity Fidelity { get; init; }
    public long ElapsedMs { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> LostProperties { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Wynik konwersji. <see cref="Web"/> = co pokazać w edytorze (null tylko gdy Rejected bez
/// placeholdera). <see cref="PreserveOriginalPart"/> = czy zapis zwrotny ma zachować oryginalny
/// media part (pass-through) zamiast re-enkodować — true dla legacy EMF/WMF/VML.
/// </summary>
public sealed class GraphicConversionResult
{
    public WebGraphicRepresentation? Web { get; init; }
    public bool PreserveOriginalPart { get; init; }
    public required GraphicConversionDiagnostics Diagnostics { get; init; }

    public bool HasWebRepresentation => Web != null;
}

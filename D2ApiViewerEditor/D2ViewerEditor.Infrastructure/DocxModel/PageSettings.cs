using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Infrastructure.DocxModel;

public enum PageOrientation
{
    Portrait,
    Landscape
}

/// <summary>
/// First piece of the explicit intermediate document model (Etap 2). Holds a section's
/// page geometry as raw OOXML twips exactly as authored — callers map to cm/px through
/// <c>OoxmlUnits</c>. Parsing is separated from rendering so section/page handling is
/// testable on the model instead of only on the final HTML string.
///
/// Page size and orientation are parsed now (even though the current HTML output does
/// not yet surface them) to seed Etap 4 without a second parse pass.
/// </summary>
public sealed class PageSettings
{
    public int? PageWidthTwips { get; init; }
    public int? PageHeightTwips { get; init; }
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>True when the section declares a w:pgMar element at all.</summary>
    public bool HasPageMargin { get; init; }

    public int? TopMarginTwips { get; init; }
    public int? BottomMarginTwips { get; init; }
    public int? LeftMarginTwips { get; init; }
    public int? RightMarginTwips { get; init; }
    public int? HeaderDistanceTwips { get; init; }
    public int? FooterDistanceTwips { get; init; }

    /// <summary>
    /// Column layout of the section (w:cols), raw as authored. Null when the section declares
    /// no w:cols at all — callers keep single-column semantics. A ColumnLayout with Count == 1
    /// is also single-column; only Count &gt; 1 is a real multi-column section (ADR-0039).
    /// </summary>
    public ColumnLayout? Columns { get; init; }

    /// <summary>
    /// Siatka dokumentu (w:docGrid) sekcji, surowo jak w pliku — tylko round-trip (ADR-0107).
    /// <see cref="DocGridType"/>: "default" | "lines" | "linesAndChars" | "snapToChars";
    /// null we wszystkich trzech = sekcja bez w:docGrid.
    /// </summary>
    public string? DocGridType { get; init; }
    public int? DocGridLinePitchTwips { get; init; }
    public int? DocGridCharSpace { get; init; }
}

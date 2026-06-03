using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Infrastructure.DocxModel;

/// <summary>
/// Parses a w:sectPr into the <see cref="PageSettings"/> model. Pure read-side mapping:
/// no defaulting, no unit conversion — values are the raw twips as authored, with nulls
/// preserved so callers keep their existing fallback semantics.
/// </summary>
public static class SectionPropertiesReader
{
    public static PageSettings ReadPageSettings(SectionProperties? sectPr)
    {
        var pageSize = sectPr?.GetFirstChild<PageSize>();
        var pageMargin = sectPr?.GetFirstChild<PageMargin>();

        var orientation = pageSize?.Orient?.Value == PageOrientationValues.Landscape
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

        return new PageSettings
        {
            PageWidthTwips = pageSize?.Width?.Value is { } w ? (int)w : null,
            PageHeightTwips = pageSize?.Height?.Value is { } h ? (int)h : null,
            Orientation = orientation,
            HasPageMargin = pageMargin != null,
            TopMarginTwips = pageMargin?.Top?.Value,
            BottomMarginTwips = pageMargin?.Bottom?.Value,
            LeftMarginTwips = pageMargin?.Left?.Value is { } l ? (int)l : null,
            RightMarginTwips = pageMargin?.Right?.Value is { } r ? (int)r : null,
            HeaderDistanceTwips = pageMargin?.Header?.Value is { } hd ? (int)hd : null,
            FooterDistanceTwips = pageMargin?.Footer?.Value is { } fd ? (int)fd : null,
        };
    }
}

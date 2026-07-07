namespace D2ViewerEditor.Infrastructure.Conversion;

/// <summary>
/// Single source of truth for Office Open XML measurement conversions.
///
/// OOXML mixes several units: twips (1/1440 inch) for page/paragraph geometry,
/// EMU (English Metric Units, 1/914400 inch) for DrawingML, half-points for font
/// size and spacing, and percentages for table widths. HTML/CSS rendering targets
/// CSS pixels, which are defined at a reference density of 96 px per inch.
///
/// Rendering DPI is fixed at <see cref="DefaultDpi"/> (96) because CSS px is a
/// physical-reference unit at 96 DPI; Word's own screen layout also assumes 96 DPI
/// for its default zoom. Keeping every conversion here removes the scattered magic
/// constants (1440 / 914400 / 9525 / 567 / 20 / 96 / 2.54) that previously lived in
/// both converters and drifted apart (e.g. cm used a 567 approximation while px used
/// the exact 1440 base for the same physical length).
/// </summary>
public static class OoxmlUnits
{
    // Reference density used to map physical units to CSS pixels.
    public const double DefaultDpi = 96.0;

    // OOXML base units per inch.
    public const int TwipsPerInch = 1440;
    public const int EmuPerInch = 914400;
    public const double PointsPerInch = 72.0;
    public const double CmPerInch = 2.54;

    // Derived per-unit factors.
    public const int TwipsPerPoint = 20;           // 1 pt = 20 twips
    public const int EmuPerPoint = 12700;          // 914400 / 72
    public const int EmuPerPixel = 9525;           // 914400 / 96 (exact at 96 DPI)
    public const int EmuPerTwip = 635;             // 914400 / 1440 (exact)
    public const int HalfPointsPerPoint = 2;       // font size / spacing are in half-points

    // --- Twips ---------------------------------------------------------------

    public static double TwipsToPoints(double twips) => twips / TwipsPerPoint;

    public static double TwipsToPixels(double twips) => twips / TwipsPerInch * DefaultDpi;

    public static double TwipsToInches(double twips) => twips / TwipsPerInch;

    public static double TwipsToCm(double twips) => twips / TwipsPerInch * CmPerInch;

    public static long TwipsToEmu(long twips) => twips * EmuPerTwip;

    public static double PointsToTwips(double points) => points * TwipsPerPoint;

    public static double CmToTwips(double cm) => cm / CmPerInch * TwipsPerInch;

    public static double InchesToTwips(double inches) => inches * TwipsPerInch;

    // --- EMU -----------------------------------------------------------------

    public static double EmuToPixels(double emu) => emu / EmuPerInch * DefaultDpi;

    public static double EmuToPoints(double emu) => emu / EmuPerPoint;

    public static double EmuToCm(double emu) => emu / EmuPerInch * CmPerInch;

    public static long PixelsToEmu(double pixels) => (long)(pixels * EmuPerPixel);

    // --- Pixels --------------------------------------------------------------

    public static double PixelsToTwips(double pixels) => pixels / DefaultDpi * TwipsPerInch;

    public static double PixelsToPoints(double pixels) => pixels / DefaultDpi * PointsPerInch;

    // --- Half-points (font size, character spacing) --------------------------

    public static double HalfPointsToPoints(double halfPoints) => halfPoints / HalfPointsPerPoint;

    public static double PointsToHalfPoints(double points) => points * HalfPointsPerPoint;
}

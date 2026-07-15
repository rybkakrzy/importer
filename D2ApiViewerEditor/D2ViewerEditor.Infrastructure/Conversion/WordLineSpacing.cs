using System.Globalization;

namespace D2ViewerEditor.Infrastructure.Conversion;

/// <summary>
/// Kalibracja interlinii „auto" (mnożnikowej) Word → CSS (PG-09).
///
/// Semantyka jest różna: w Wordzie w:line/240 przy lineRule=auto to mnożnik
/// POJEDYNCZEGO odstępu, a pojedynczy odstęp wynika z metryk fontu
/// (ascent + descent + lineGap ≈ 1,15–1,33 em zależnie od kroju). W CSS
/// bezjednostkowe line-height to dokładnie N × font-size. Emisja surowego
/// N = line/240 dawała tekst ~15–25% ciaśniejszy niż w Wordzie — największy
/// pojedynczy wkład w rozjazd wysokości dokumentu (AUDIT PG-09).
///
/// Reader emituje więc wartość SKALIBROWANĄ (mnożnik × współczynnik single
/// fontu) plus marker <c>--w-line-tw</c> z oryginalną wartością w 240-tych,
/// z którego writer bezstratnie odtwarza w:line/lineRule=auto — analogicznie
/// do istniejącego markera <c>--w-line-rule:atLeast</c>. Marker dotyczy TYLKO
/// wartości bezjednostkowych; przy line-height w pt (exact/atLeast) writer go
/// ignoruje.
/// </summary>
public static class WordLineSpacing
{
    /// <summary>Word: mnożnik auto jest wyrażony w 240-tych częściach linii.</summary>
    public const double LineUnitsPerSingle = 240.0;

    /// <summary>
    /// Współczynnik „pojedynczego odstępu" (single spacing) w em dla kroju, gdy
    /// metryki fontu są znane, inaczej <see cref="DefaultSingleFactor"/>.
    /// Wartości = (ascent + descent + lineGap) / unitsPerEm z tabel hhea/OS2.
    /// </summary>
    public static double SingleFactor(string? fontFamily)
    {
        var key = NormalizeFontName(fontFamily);
        return key != null && Factors.TryGetValue(key, out var f) ? f : DefaultSingleFactor;
    }

    /// <summary>
    /// Fallback dla krojów spoza tabeli (np. firmowe). 1,20 to środek zakresu
    /// typowych metryk (1,15–1,33) — mniejszy błąd niż surowe 1,0.
    /// </summary>
    public const double DefaultSingleFactor = 1.20;

    /// <summary>
    /// CSS interlinii dla lineRule=auto: skalibrowane bezjednostkowe line-height
    /// + marker round-trip, np. „line-height:1.294;--w-line-tw:259;".
    /// </summary>
    public static string AutoCss(int lineTwips, string? fontFamily)
    {
        var calibrated = lineTwips / LineUnitsPerSingle * SingleFactor(fontFamily);
        return string.Format(CultureInfo.InvariantCulture,
            "line-height:{0:0.###};--w-line-tw:{1};", calibrated, lineTwips);
    }

    private static string? NormalizeFontName(string? fontFamily)
    {
        if (string.IsNullOrWhiteSpace(fontFamily)) return null;
        // Pierwszy krój z ewentualnej listy CSS, bez cudzysłowów.
        var first = fontFamily.Split(',')[0].Trim().Trim('\'', '"').Trim();
        return first.Length == 0 ? null : first.ToLowerInvariant();
    }

    // (ascent+descent+lineGap)/unitsPerEm popularnych krojów Office/Windows.
    private static readonly Dictionary<string, double> Factors = new()
    {
        ["calibri"] = 1.221,
        ["calibri light"] = 1.221,
        ["cambria"] = 1.171,
        ["arial"] = 1.150,
        ["helvetica"] = 1.150,
        ["times new roman"] = 1.149,
        ["courier new"] = 1.133,
        ["verdana"] = 1.215,
        ["tahoma"] = 1.207,
        ["segoe ui"] = 1.330,
        ["georgia"] = 1.136,
    };
}

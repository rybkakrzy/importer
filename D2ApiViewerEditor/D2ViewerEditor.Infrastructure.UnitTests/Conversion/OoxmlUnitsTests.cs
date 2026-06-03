using D2ViewerEditor.Infrastructure.Conversion;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Conversion;

/// <summary>
/// Reference-value and round-trip tests for the central OOXML unit conversions.
/// One inch is the anchor: 1440 twips = 914400 EMU = 96 px = 72 pt = 2.54 cm.
/// </summary>
[TestFixture]
public class OoxmlUnitsTests
{
    [Test]
    public void OneInch_AllUnitsAgree()
    {
        OoxmlUnits.TwipsToInches(OoxmlUnits.TwipsPerInch).Should().Be(1.0);
        OoxmlUnits.TwipsToPixels(1440).Should().BeApproximately(96.0, 1e-9);
        OoxmlUnits.TwipsToPoints(1440).Should().BeApproximately(72.0, 1e-9);
        OoxmlUnits.TwipsToCm(1440).Should().BeApproximately(2.54, 1e-9);
        OoxmlUnits.EmuToPixels(914400).Should().BeApproximately(96.0, 1e-9);
        OoxmlUnits.EmuToPoints(914400).Should().BeApproximately(72.0, 1e-9);
        OoxmlUnits.EmuToCm(914400).Should().BeApproximately(2.54, 1e-9);
    }

    [Test]
    public void Twips_PointConversion_IsExact()
    {
        OoxmlUnits.TwipsToPoints(240).Should().Be(12.0);   // 12 pt body line
        OoxmlUnits.PointsToTwips(12).Should().Be(240.0);
    }

    [Test]
    public void HalfPoints_ToPoints()
    {
        OoxmlUnits.HalfPointsToPoints(24).Should().Be(12.0);  // sz=24 → 12 pt
        OoxmlUnits.HalfPointsToPoints(22).Should().Be(11.0);  // sz=22 → 11 pt (Calibri default)
        OoxmlUnits.PointsToHalfPoints(11).Should().Be(22.0);
    }

    [Test]
    public void Emu_Pixel_RoundTrip_OneCssPixel()
    {
        OoxmlUnits.EmuPerPixel.Should().Be(9525);
        OoxmlUnits.PixelsToEmu(1).Should().Be(9525);
        OoxmlUnits.EmuToPixels(9525).Should().BeApproximately(1.0, 1e-9);
    }

    [Test]
    public void Twips_Pixel_RoundTrip()
    {
        OoxmlUnits.PixelsToTwips(96).Should().BeApproximately(1440.0, 1e-9);
        OoxmlUnits.TwipsToPixels(1440).Should().BeApproximately(96.0, 1e-9);
    }

    [Test]
    public void Cm_Twips_RoundTrip_SupersedesThe567Approximation()
    {
        // 567 was a rounded stand-in for the exact 1440/2.54 = 566.929... twips per cm.
        OoxmlUnits.CmToTwips(2.5).Should().BeApproximately(1417.32, 0.01);
        OoxmlUnits.TwipsToCm(1417).Should().BeApproximately(2.5, 0.001);
    }

    [Test]
    public void NamedConstants_MatchDerivedValues()
    {
        OoxmlUnits.EmuPerPoint.Should().Be(OoxmlUnits.EmuPerInch / 72);
        OoxmlUnits.EmuPerPixel.Should().Be((int)(OoxmlUnits.EmuPerInch / OoxmlUnits.DefaultDpi));
        OoxmlUnits.TwipsPerPoint.Should().Be(OoxmlUnits.TwipsPerInch / 72);
    }
}

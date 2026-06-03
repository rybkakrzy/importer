using D2ViewerEditor.Infrastructure.DocxModel;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.DocxModel;

/// <summary>
/// Model-level tests for the section/page reader (Etap 2). These assert the parsed
/// intermediate model directly — the kind of test the previous XML→string pipeline
/// could not express.
/// </summary>
[TestFixture]
public class SectionPropertiesReaderTests
{
    [Test]
    public void ReadPageSettings_ParsesSizeMarginsAndDistances()
    {
        var sectPr = new SectionProperties(
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin { Top = 1417, Bottom = 1418, Left = 1440, Right = 1441, Header = 708, Footer = 709 });

        var page = SectionPropertiesReader.ReadPageSettings(sectPr);

        page.PageWidthTwips.Should().Be(11906);
        page.PageHeightTwips.Should().Be(16838);
        page.Orientation.Should().Be(PageOrientation.Portrait);
        page.HasPageMargin.Should().BeTrue();
        page.TopMarginTwips.Should().Be(1417);
        page.BottomMarginTwips.Should().Be(1418);
        page.LeftMarginTwips.Should().Be(1440);
        page.RightMarginTwips.Should().Be(1441);
        page.HeaderDistanceTwips.Should().Be(708);
        page.FooterDistanceTwips.Should().Be(709);
    }

    [Test]
    public void ReadPageSettings_DetectsLandscape()
    {
        var sectPr = new SectionProperties(
            new PageSize { Width = 16838, Height = 11906, Orient = PageOrientationValues.Landscape });

        var page = SectionPropertiesReader.ReadPageSettings(sectPr);

        page.Orientation.Should().Be(PageOrientation.Landscape);
        page.PageWidthTwips.Should().Be(16838);
    }

    [Test]
    public void ReadPageSettings_NoPageMargin_HasPageMarginFalse()
    {
        var sectPr = new SectionProperties(new PageSize { Width = 11906, Height = 16838 });

        var page = SectionPropertiesReader.ReadPageSettings(sectPr);

        page.HasPageMargin.Should().BeFalse();
        page.TopMarginTwips.Should().BeNull();
    }

    [Test]
    public void ReadPageSettings_NullSectPr_ReturnsPortraitEmpty()
    {
        var page = SectionPropertiesReader.ReadPageSettings(null);

        page.HasPageMargin.Should().BeFalse();
        page.PageWidthTwips.Should().BeNull();
        page.Orientation.Should().Be(PageOrientation.Portrait);
    }

    [Test]
    public void ReadPageSettings_PreservesNegativeTopMargin()
    {
        // Mirror/overlapping headers can author a negative top margin; the model keeps the
        // sign (callers take the magnitude when computing the printable band).
        var sectPr = new SectionProperties(
            new PageMargin { Top = -200, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720 });

        var page = SectionPropertiesReader.ReadPageSettings(sectPr);

        page.TopMarginTwips.Should().Be(-200);
    }
}

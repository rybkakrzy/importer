using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Etap 4 (backend): page size + orientation survive the HTML→DOCX→HTML round-trip.
/// Previously the writer hardcoded A4 portrait, so landscape and non-A4 pages were lost.
/// </summary>
[TestFixture]
public class PageSizeRoundTripTests
{
    private HtmlToDocxConverter _writer = null!;
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup()
    {
        _writer = new HtmlToDocxConverter();
        _reader = new DocxToHtmlConverter();
    }

    private DocumentContent WriteThenRead(PageSize? pageSize)
    {
        var bytes = _writer.Convert("<p>x</p>", pageSize: pageSize);
        using var ms = new MemoryStream(bytes);
        return _reader.Convert(ms);
    }

    [Test]
    public void Landscape_A4_RoundTrips()
    {
        var result = WriteThenRead(new PageSize { WidthCm = 29.7, HeightCm = 21.0, Orientation = "landscape" });

        result.PageSize.Should().NotBeNull();
        result.PageSize!.Orientation.Should().Be("landscape");
        result.PageSize.WidthCm.Should().BeApproximately(29.7, 0.05);
        result.PageSize.HeightCm.Should().BeApproximately(21.0, 0.05);
        result.PageSize.WidthCm.Should().BeGreaterThan(result.PageSize.HeightCm);
    }

    [Test]
    public void Portrait_CustomSize_RoundTrips()
    {
        // A5 portrait.
        var result = WriteThenRead(new PageSize { WidthCm = 14.8, HeightCm = 21.0, Orientation = "portrait" });

        result.PageSize.Should().NotBeNull();
        result.PageSize!.Orientation.Should().Be("portrait");
        result.PageSize.WidthCm.Should().BeApproximately(14.8, 0.05);
        result.PageSize.HeightCm.Should().BeApproximately(21.0, 0.05);
    }

    [Test]
    public void NullPageSize_FallsBackToA4Portrait()
    {
        var result = WriteThenRead(null);

        result.PageSize.Should().NotBeNull();
        result.PageSize!.Orientation.Should().Be("portrait");
        result.PageSize.WidthCm.Should().BeApproximately(21.0, 0.05);   // A4 width
        result.PageSize.HeightCm.Should().BeApproximately(29.7, 0.05);  // A4 height
    }
}

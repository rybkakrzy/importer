using System.Text;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Konwerter traktuje grafiki jako niezaufane wejście: sanitizacja SVG (script/handlery/
/// zewnętrzne referencje), blokada XXE/DTD, brak crashy na uszkodzonych danych.
/// </summary>
[TestFixture]
public class GraphicConversionSecurityTests
{
    private GraphicConversionService _svc = null!;

    [SetUp]
    public void Setup() => _svc = new GraphicConversionService();

    [Test]
    public void SanitizeSvg_RemovesScript_EventHandlers_ExternalReferences()
    {
        var dirty =
            "<svg xmlns='http://www.w3.org/2000/svg'>" +
            "<script>alert(1)</script>" +
            "<rect width='10' height='10' onload='steal()' onclick='x()'/>" +
            "<image href='http://evil.example/track.png'/>" +
            "<a xlink:href='javascript:alert(2)' xmlns:xlink='http://www.w3.org/1999/xlink'>x</a>" +
            "</svg>";

        var clean = _svc.SanitizeSvg(dirty);

        clean.Should().NotBeNull();
        clean!.Should().NotContain("script");
        clean.Should().NotContain("onload");
        clean.Should().NotContain("onclick");
        clean.Should().NotContain("evil.example");
        clean.Should().NotContain("javascript:");
        clean.Should().Contain("<rect"); // bezpieczna treść zostaje
    }

    [Test]
    public void SanitizeSvg_KeepsSafeDataHrefAndFragment()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg'><image href='data:image/png;base64,AAAA'/><use href='#g'/></svg>";
        // 'use' jest usuwany (wektor ataku), ale data: href na image musi przetrwać.
        var clean = _svc.SanitizeSvg(svg);
        clean.Should().NotBeNull();
        clean!.Should().Contain("data:image/png;base64,AAAA");
        clean.Should().NotContain("<use"); // 'use' na liście zakazanych
    }

    [Test]
    public void SanitizeSvg_BlocksXxeViaProhibitedDtd()
    {
        var xxe =
            "<?xml version='1.0'?>" +
            "<!DOCTYPE svg [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]>" +
            "<svg xmlns='http://www.w3.org/2000/svg'><text>&xxe;</text></svg>";

        // DtdProcessing.Prohibit → parsowanie rzuca → null (odrzucone), bez odczytu pliku.
        _svc.SanitizeSvg(xxe).Should().BeNull();
    }

    [Test]
    public void SanitizeSvg_NonSvgRoot_ReturnsNull()
        => _svc.SanitizeSvg("<html><body>x</body></html>").Should().BeNull();

    [Test]
    public void ConvertForEditor_MalformedSvg_FallsBackWithoutThrowing()
    {
        var bad = Encoding.UTF8.GetBytes("<svg><rect width='10'"); // niedomknięty
        var result = _svc.ConvertForEditor(new GraphicSource { Data = bad, ContentType = "image/svg+xml" });
        // Nie rzuca; mapuje na fallback/unsupported z placeholderem.
        result.Web.Should().NotBeNull();
        result.Diagnostics.Status.Should().BeOneOf(GraphicConversionStatus.Fallback, GraphicConversionStatus.Unsupported);
    }

    [Test]
    public void ConvertForEditor_CorruptedMetafile_DoesNotThrow()
    {
        var corrupt = new byte[64];
        corrupt[0] = 1; // wygląda jak początek EMF, reszta śmieci
        var result = _svc.ConvertForEditor(new GraphicSource { Data = corrupt, ContentType = "image/x-emf" });
        result.Should().NotBeNull();
        result.Web.Should().NotBeNull(); // placeholder, nie wyjątek
    }

    [Test]
    public void PlaceholderDimensions_AreClampedToMax()
    {
        var opts = new GraphicConversionOptions { MaxPlaceholderWidthPx = 100, MaxPlaceholderHeightPx = 100 };
        // Ogromny EMF frame → po clampie ≤ 100.
        var huge = BuildEmfFrame(10_000_000, 10_000_000);
        var result = _svc.ConvertForEditor(new GraphicSource { Data = huge, ContentType = "image/x-emf" }, opts);
        result.Web!.WidthPx.Should().BeLessThanOrEqualTo(100);
        result.Web.HeightPx.Should().BeLessThanOrEqualTo(100);
    }

    private static byte[] BuildEmfFrame(int right, int bottom)
    {
        var d = new byte[88];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(d, 1);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(4), 88);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(32), right);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(36), bottom);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(40), 0x464D4520);
        return d;
    }
}

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
        // Wewnętrzny <use href='#id'> jest bezpieczny (wskazuje już-sanityzowaną treść tego samego
        // dokumentu) i MUSI przetrwać — typowe logo to <defs>+<use>; data: href na image także.
        var clean = _svc.SanitizeSvg(svg);
        clean.Should().NotBeNull();
        clean!.Should().Contain("data:image/png;base64,AAAA");
        clean.Should().Contain("<use");
    }

    [Test]
    public void SanitizeSvg_LogoBuiltFromDefsAndUse_KeepsVisibleContent()
    {
        // Regresja: logo (np. korporacyjne) zbudowane wyłącznie z <defs>+<use> renderowało się
        // jako pusty biały obraz, bo sanitizer wycinał WSZYSTKIE <use> zostawiając niewidoczne defs.
        var svg =
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' viewBox='0 0 100 40'>" +
            "<defs><path id='lion' d='M10 10 C 20 0, 40 0, 50 10 Z' fill='#ff6200'/></defs>" +
            "<use xlink:href='#lion'/><use href='#lion' x='50'/>" +
            "</svg>";

        var clean = _svc.SanitizeSvg(svg);

        clean.Should().NotBeNull();
        clean!.Should().Contain("<defs");
        System.Text.RegularExpressions.Regex.Matches(clean, "<use").Count.Should().Be(2);
    }

    [Test]
    public void SanitizeSvg_UseWithExternalHref_IsRemoved()
    {
        var svg =
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<use xlink:href='http://evil.example/doc.svg#x'/>" +
            "<use href='data:image/svg+xml;base64,AAAA'/>" +
            "<use/>" +
            "<rect width='10' height='10'/>" +
            "</svg>";

        var clean = _svc.SanitizeSvg(svg);

        clean.Should().NotBeNull();
        clean!.Should().NotContain("<use");
        clean.Should().NotContain("evil.example");
        clean.Should().Contain("<rect");
    }

    [Test]
    public void SanitizeSvg_Utf8BomPrefix_IsAccepted()
    {
        // Realne pliki logo z Windows bywają zapisane z BOM — U+FEFF na początku stringa
        // wywalał parser XML i cały (poprawny) SVG był odrzucany, a obraz znikał z edytora.
        var svg = "\uFEFF<svg xmlns='http://www.w3.org/2000/svg'><rect width='10' height='10'/></svg>";
        var clean = _svc.SanitizeSvg(svg);
        clean.Should().NotBeNull();
        clean!.Should().Contain("<rect");
    }

    [Test]
    public void SanitizeSvg_BenignIllustratorDoctype_IsAccepted()
    {
        // Starsze eksporty (Illustrator) niosą DOCTYPE bez encji — DtdProcessing.Ignore pomija go
        // zamiast odrzucać cały plik.
        var svg =
            "<?xml version='1.0' encoding='utf-8'?>" +
            "<!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG 1.1//EN\" \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\">" +
            "<svg xmlns='http://www.w3.org/2000/svg'><rect width='10' height='10'/></svg>";
        var clean = _svc.SanitizeSvg(svg);
        clean.Should().NotBeNull();
        clean!.Should().Contain("<rect");
    }

    [Test]
    public void SanitizeSvg_BlocksXxeViaProhibitedDtd()
    {
        var xxe =
            "<?xml version='1.0'?>" +
            "<!DOCTYPE svg [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]>" +
            "<svg xmlns='http://www.w3.org/2000/svg'><text>&xxe;</text></svg>";

        // DtdProcessing.Ignore → DTD pomijany bez przetwarzania, encja pozostaje niezdefiniowana
        // → parser rzuca → null (odrzucone), bez odczytu pliku.
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

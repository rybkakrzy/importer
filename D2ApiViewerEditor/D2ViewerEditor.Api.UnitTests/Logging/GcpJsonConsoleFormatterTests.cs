using System.Text.Json;
using D2ViewerEditor.Api.Logging;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Logging;

/// <summary>
/// Formatter logów do Google Cloud Logging: mapowanie LogLevel → `severity` (kluczowe, bo bez tego
/// pola Cloud Logging pokazuje wszystko jako INFO, także wyjątki).
/// </summary>
[TestFixture]
public class GcpJsonConsoleFormatterTests
{
    private static JsonDocument WriteAndParse(LogLevel level, string message, Exception? ex = null)
    {
        var formatter = new GcpJsonConsoleFormatter();
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            level, "Test.Category", new EventId(0), message, ex, (s, _) => s);

        using var sw = new StringWriter();
        formatter.Write(in entry, scopeProvider: null, sw);
        return JsonDocument.Parse(sw.ToString());
    }

    [TestCase(LogLevel.Information, "INFO")]
    [TestCase(LogLevel.Warning, "WARNING")]
    [TestCase(LogLevel.Error, "ERROR")]
    [TestCase(LogLevel.Critical, "CRITICAL")]
    [TestCase(LogLevel.Debug, "DEBUG")]
    [TestCase(LogLevel.Trace, "DEBUG")]
    public void Write_MapsLogLevelToGcpSeverity(LogLevel level, string expectedSeverity)
    {
        using var doc = WriteAndParse(level, "wiadomość testowa");

        doc.RootElement.GetProperty("severity").GetString().Should().Be(expectedSeverity);
        doc.RootElement.GetProperty("message").GetString().Should().Contain("wiadomość testowa");
    }

    [Test]
    public void Write_Exception_IsErrorSeverity_AndCarriesStackTraceInMessage()
    {
        var ex = new InvalidOperationException("coś poszło nie tak");

        using var doc = WriteAndParse(LogLevel.Error, "Wystąpił wyjątek", ex);

        doc.RootElement.GetProperty("severity").GetString().Should().Be("ERROR");
        doc.RootElement.GetProperty("message").GetString().Should().Contain("InvalidOperationException");
        doc.RootElement.GetProperty("exceptionType").GetString().Should().Be("System.InvalidOperationException");
    }

    [Test]
    public void Write_EmitsSingleJsonLine()
    {
        var formatter = new GcpJsonConsoleFormatter();
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "linia", null, (s, _) => s);

        using var sw = new StringWriter();
        formatter.Write(in entry, null, sw);

        sw.ToString().TrimEnd('\r', '\n').Should().NotContain("\n"); // jedna linia = jeden wpis w Cloud Logging
    }
}

using System.Text.Json;
using D2ViewerEditor.Api.Logging;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Logging;

/// <summary>
/// Structured JSON log formatter (ELK + GCP): LogLevel → `severity` (GCP) and `level`, plus
/// service/environment enrichment and scope key/values (correlationId, business context).
/// </summary>
[TestFixture]
public class GcpJsonConsoleFormatterTests
{
    private static GcpJsonConsoleFormatter Formatter(string service = "TestSvc", string env = "TST") =>
        new(Options.Create(new StructuredLogFormatterOptions { Service = service, Environment = env }));

    private static JsonDocument WriteAndParse(LogLevel level, string message, Exception? ex = null)
    {
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            level, "Test.Category", new EventId(0), message, ex, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, scopeProvider: null, sw);
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
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "linia", null, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, null, sw);

        sw.ToString().TrimEnd('\r', '\n').Should().NotContain("\n"); // jedna linia = jeden wpis
    }

    [Test]
    public void Write_EnrichesWithServiceEnvironmentAndLevel()
    {
        using var doc = WriteAndParse(LogLevel.Information, "ping");

        doc.RootElement.GetProperty("service").GetString().Should().Be("TestSvc");
        doc.RootElement.GetProperty("environment").GetString().Should().Be("TST");
        doc.RootElement.GetProperty("level").GetString().Should().Be("Information");
        doc.RootElement.GetProperty("timestamp").GetString().Should().NotBeNullOrEmpty();
    }

    [Test]
    public void Write_SerializesScopeKeyValues_ForCorrelation()
    {
        var scopeProvider = new LoggerExternalScopeProvider();
        using var _ = scopeProvider.Push(new Dictionary<string, object>
        {
            ["correlationId"] = "corr-123",
            ["masterId"] = "m-1"
        });

        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "msg", null, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, scopeProvider, sw);
        using var doc = JsonDocument.Parse(sw.ToString());

        doc.RootElement.GetProperty("correlationId").GetString().Should().Be("corr-123");
        doc.RootElement.GetProperty("masterId").GetString().Should().Be("m-1");
    }
}

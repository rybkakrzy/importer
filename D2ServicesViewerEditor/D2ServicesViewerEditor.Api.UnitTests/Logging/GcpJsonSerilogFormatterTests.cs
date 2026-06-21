using System.Text.Json;
using D2ServicesViewerEditor.Api.Logging;
using FluentAssertions;
using NUnit.Framework;
using Serilog.Events;
using Serilog.Parsing;

namespace D2ServicesViewerEditor.Api.UnitTests.Logging;

[TestFixture]
public class GcpJsonSerilogFormatterTests
{
    private static readonly MessageTemplateParser TemplateParser = new();

    [TestCase(LogEventLevel.Information, "INFO")]
    [TestCase(LogEventLevel.Warning, "WARNING")]
    [TestCase(LogEventLevel.Error, "ERROR")]
    [TestCase(LogEventLevel.Fatal, "CRITICAL")]
    [TestCase(LogEventLevel.Debug, "DEBUG")]
    [TestCase(LogEventLevel.Verbose, "DEBUG")]
    public void Format_MapsSerilogLevelToGcpSeverity(LogEventLevel level, string expectedSeverity)
    {
        var formatter = new GcpJsonSerilogFormatter("D2Services", "UAT");
        var logEvent = new LogEvent(
            DateTimeOffset.UtcNow,
            level,
            exception: null,
            TemplateParser.Parse("Message"),
            []);

        using var sw = new StringWriter();
        formatter.Format(logEvent, sw);
        using var json = JsonDocument.Parse(sw.ToString());

        json.RootElement.GetProperty("severity").GetString().Should().Be(expectedSeverity);
        json.RootElement.GetProperty("level").GetString().Should().Be(level.ToString());
        json.RootElement.GetProperty("service").GetString().Should().Be("D2Services");
        json.RootElement.GetProperty("environment").GetString().Should().Be("UAT");
        json.RootElement.GetProperty("message").GetString().Should().Be("Message");
    }

    [Test]
    public void Format_WhenSourceContextPresent_IncludesCategory()
    {
        var formatter = new GcpJsonSerilogFormatter();
        var logEvent = new LogEvent(
            DateTimeOffset.UtcNow,
            LogEventLevel.Information,
            exception: null,
            TemplateParser.Parse("Test"),
            [new LogEventProperty("SourceContext", new ScalarValue("App.Category"))]);

        using var sw = new StringWriter();
        formatter.Format(logEvent, sw);
        using var json = JsonDocument.Parse(sw.ToString());

        json.RootElement.GetProperty("category").GetString().Should().Be("App.Category");
    }

    [Test]
    public void Format_WhenExceptionPresent_IncludesExceptionTypeAndStackInMessage()
    {
        var formatter = new GcpJsonSerilogFormatter();
        var exception = new InvalidOperationException("boom");
        var logEvent = new LogEvent(
            DateTimeOffset.UtcNow,
            LogEventLevel.Error,
            exception,
            TemplateParser.Parse("Request failed"),
            []);

        using var sw = new StringWriter();
        formatter.Format(logEvent, sw);
        using var json = JsonDocument.Parse(sw.ToString());

        json.RootElement.GetProperty("exceptionType").GetString().Should().Be(typeof(InvalidOperationException).FullName);
        json.RootElement.GetProperty("message").GetString().Should().Contain("Request failed").And.Contain("InvalidOperationException");
    }

    [Test]
    public void Format_WhenStructuredPropertiesPresent_EmitsThemAsTopLevelFields()
    {
        var formatter = new GcpJsonSerilogFormatter();
        var logEvent = new LogEvent(
            DateTimeOffset.UtcNow,
            LogEventLevel.Information,
            exception: null,
            TemplateParser.Parse("Done {statusCode}"),
            [
                new LogEventProperty("statusCode", new ScalarValue(200)),
                new LogEventProperty("correlationId", new ScalarValue("corr-1")),
                new LogEventProperty("TraceId", new ScalarValue("trace-1")),
                new LogEventProperty("SpanId", new ScalarValue("span-1"))
            ]);

        using var sw = new StringWriter();
        formatter.Format(logEvent, sw);
        using var json = JsonDocument.Parse(sw.ToString());

        json.RootElement.GetProperty("statusCode").GetInt32().Should().Be(200);
        json.RootElement.GetProperty("correlationId").GetString().Should().Be("corr-1");
        json.RootElement.GetProperty("traceId").GetString().Should().Be("trace-1");
        json.RootElement.GetProperty("spanId").GetString().Should().Be("span-1");
    }
}

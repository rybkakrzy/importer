using System.Diagnostics;
using System.Text.Json;
using D2ViewerEditor.Api.Logging;
using D2ViewerEditor.Api.Middleware;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
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

    private static GcpJsonConsoleFormatter Formatter(
        StructuredLogFormatterOptions options, IHttpContextAccessor? http = null) =>
        new(Options.Create(options), http);

    private static JsonDocument Write(
        GcpJsonConsoleFormatter formatter, LogLevel level, string message,
        Exception? ex = null, IExternalScopeProvider? scopes = null)
    {
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            level, "Test.Category", new EventId(0), message, ex, (s, _) => s);
        using var sw = new StringWriter();
        formatter.Write(in entry, scopes, sw);
        return JsonDocument.Parse(sw.ToString());
    }

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
        // Migration (ECS default ON): service/environment moved into the ECS nested `service` object;
        // flat `level`/`timestamp` are kept for back-compat.
        using var doc = WriteAndParse(LogLevel.Information, "ping");

        doc.RootElement.GetProperty("service").GetProperty("name").GetString().Should().Be("TestSvc");
        doc.RootElement.GetProperty("service").GetProperty("environment").GetString().Should().Be("TST");
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

    [Test]
    public void Write_TypeValueInScope_SerializesAsName_WithoutFallback()
    {
        var scopeProvider = new LoggerExternalScopeProvider();
        using var scope = scopeProvider.Push(new Dictionary<string, object>
        {
            ["requestType"] = typeof(GcpJsonConsoleFormatterTests) // boxed System.RuntimeType
        });

        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "msg", null, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, scopeProvider, sw);
        using var doc = JsonDocument.Parse(sw.ToString());

        doc.RootElement.TryGetProperty("serializationError", out _).Should().BeFalse();
        doc.RootElement.GetProperty("requestType").GetString()
            .Should().Be(typeof(GcpJsonConsoleFormatterTests).FullName);
    }

    [Test]
    public void Write_NestedTypeValue_SerializesAsName_WithoutFallback()
    {
        var scopeProvider = new LoggerExternalScopeProvider();
        using var scope = scopeProvider.Push(new Dictionary<string, object>
        {
            ["request"] = new { Handler = typeof(GcpJsonConsoleFormatterTests), Name = "x" }
        });

        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "msg", null, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, scopeProvider, sw);
        using var doc = JsonDocument.Parse(sw.ToString());

        doc.RootElement.TryGetProperty("serializationError", out _).Should().BeFalse();
        doc.RootElement.GetProperty("request").GetProperty("Handler").GetString()
            .Should().Be(typeof(GcpJsonConsoleFormatterTests).FullName);
    }

    [Test]
    public void Write_ExceptionObjectInScope_SerializesWithoutFallback()
    {
        Exception caught;
        try { throw new InvalidOperationException("boom"); }
        catch (Exception e) { caught = e; } // populated TargetSite (MethodBase), StackTrace …

        var scopeProvider = new LoggerExternalScopeProvider();
        using var scope = scopeProvider.Push(new Dictionary<string, object>
        {
            ["failure"] = caught
        });

        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<string>(
            LogLevel.Information, "Cat", new EventId(0), "msg", null, (s, _) => s);

        using var sw = new StringWriter();
        Formatter().Write(in entry, scopeProvider, sw);
        using var doc = JsonDocument.Parse(sw.ToString());

        doc.RootElement.TryGetProperty("serializationError", out _).Should().BeFalse();
        doc.RootElement.GetProperty("failure").GetProperty("Message").GetString().Should().Be("boom");
    }

    [Test]
    public void Write_Exception_EmitsStructuredEcsError_WithInnerChain()
    {
        var ex = new InvalidOperationException("outer", new ArgumentException("inner"));

        using var doc = WriteAndParse(LogLevel.Error, "failed", ex);

        var error = doc.RootElement.GetProperty("error");
        error.GetProperty("type").GetString().Should().Be("System.InvalidOperationException");
        error.GetProperty("message").GetString().Should().Be("outer");
        error.GetProperty("stack_trace").GetString().Should().Contain("ArgumentException"); // full ToString()
        error.GetProperty("inner")[0].GetProperty("type").GetString().Should().Be("System.ArgumentException");
        error.GetProperty("inner")[0].GetProperty("message").GetString().Should().Be("inner");
        // Diagnostic classification so a 500 can be triaged from the log alone (InvalidOperation → code).
        doc.RootElement.GetProperty("event").GetProperty("reason").GetString().Should().Be("code");
    }

    [Test]
    public void Write_MessageTemplate_PreservesTemplateAndStructuredProperties()
    {
        var scope = new LoggerExternalScopeProvider();
        var state = new[]
        {
            new KeyValuePair<string, object>("OrderId", 42),
            new KeyValuePair<string, object>("{OriginalFormat}", "Processed order {OrderId}")
        };
        var entry = new Microsoft.Extensions.Logging.Abstractions.LogEntry<IReadOnlyList<KeyValuePair<string, object>>>(
            LogLevel.Information, "Cat", new EventId(7, "Processed"), state, null,
            (s, _) => "Processed order 42");

        using var sw = new StringWriter();
        Formatter().Write(in entry, scope, sw);
        using var doc = JsonDocument.Parse(sw.ToString());

        doc.RootElement.GetProperty("OrderId").GetInt32().Should().Be(42);
        doc.RootElement.GetProperty("log").GetProperty("template").GetString().Should().Be("Processed order {OrderId}");
        doc.RootElement.GetProperty("eventId").GetInt32().Should().Be(7);
    }

    [Test]
    public void Write_WithActivity_EmitsGcpTraceFieldsAndEcsTrace()
    {
        var options = new StructuredLogFormatterOptions { Service = "Svc", Environment = "TST", ProjectId = "my-proj" };
        var activity = new Activity("op");
        activity.SetIdFormat(ActivityIdFormat.W3C);
        activity.Start();
        try
        {
            using var doc = Write(Formatter(options), LogLevel.Information, "msg");

            var traceId = activity.TraceId.ToString();
            doc.RootElement.GetProperty("logging.googleapis.com/trace").GetString()
                .Should().Be($"projects/my-proj/traces/{traceId}");
            doc.RootElement.GetProperty("logging.googleapis.com/spanId").GetString()
                .Should().Be(activity.SpanId.ToString());
            doc.RootElement.TryGetProperty("logging.googleapis.com/trace_sampled", out _).Should().BeTrue();
            doc.RootElement.GetProperty("trace").GetProperty("id").GetString().Should().Be(traceId);
            doc.RootElement.GetProperty("span").GetProperty("id").GetString().Should().Be(activity.SpanId.ToString());
        }
        finally
        {
            activity.Stop();
        }
    }

    [Test]
    public void Write_RedactsSensitiveProperties_CaseAndSeparatorInsensitive()
    {
        var scopeProvider = new LoggerExternalScopeProvider();
        using var _ = scopeProvider.Push(new Dictionary<string, object>
        {
            ["Password"] = "hunter2",
            ["accessToken"] = "jwt.value.here",
            ["OrderId"] = 99
        });

        using var doc = Write(Formatter(), LogLevel.Information, "msg", scopes: scopeProvider);

        doc.RootElement.GetProperty("Password").GetString().Should().Be("[REDACTED]");
        doc.RootElement.GetProperty("accessToken").GetString().Should().Be("[REDACTED]");
        doc.RootElement.GetProperty("OrderId").GetInt32().Should().Be(99); // non-sensitive untouched
    }

    [Test]
    public void Write_DoesNotLetScopeOverwriteSystemFields()
    {
        var scopeProvider = new LoggerExternalScopeProvider();
        using var _ = scopeProvider.Push(new Dictionary<string, object>
        {
            ["severity"] = "HACKED",
            ["message"] = "HACKED",
            ["trace"] = "HACKED"
        });

        using var doc = Write(Formatter(), LogLevel.Error, "real message", scopes: scopeProvider);

        doc.RootElement.GetProperty("severity").GetString().Should().Be("ERROR");
        doc.RootElement.GetProperty("message").GetString().Should().Be("real message");
    }

    [Test]
    public void Write_WithoutHttpContext_OmitsHttpFields()
    {
        using var doc = Write(Formatter(), LogLevel.Information, "background work"); // no accessor

        doc.RootElement.TryGetProperty("httpRequest", out _).Should().BeFalse();
        doc.RootElement.TryGetProperty("http", out _).Should().BeFalse();
    }

    [Test]
    public void Write_WithHttpContext_EmitsHttpAndUrl_AndMasksSensitiveQuery_AndCorrelationId()
    {
        var ctx = new DefaultHttpContext();
        ctx.Request.Method = "POST";
        ctx.Request.Path = "/api/v1/document/save";
        ctx.Request.QueryString = new QueryString("?token=secretjwt&page=2");
        ctx.Response.StatusCode = 500;
        ctx.Items[RequestObservabilityMiddleware.CorrelationIdItemKey] = "corr-xyz";
        var accessor = new HttpContextAccessor { HttpContext = ctx };

        using var doc = Write(Formatter(new StructuredLogFormatterOptions { Service = "Svc" }, accessor),
            LogLevel.Error, "500 happened");

        doc.RootElement.GetProperty("http").GetProperty("request").GetProperty("method").GetString().Should().Be("POST");
        doc.RootElement.GetProperty("http").GetProperty("response").GetProperty("status_code").GetInt32().Should().Be(500);
        doc.RootElement.GetProperty("url").GetProperty("path").GetString().Should().Be("/api/v1/document/save");
        doc.RootElement.GetProperty("url").GetProperty("query").GetString().Should().Be("token=[REDACTED]&page=2");
        doc.RootElement.GetProperty("correlation_id").GetString().Should().Be("corr-xyz");
        doc.RootElement.GetProperty("labels").GetProperty("correlation_id").GetString().Should().Be("corr-xyz");
        doc.RootElement.GetProperty("httpRequest").GetProperty("status").GetInt32().Should().Be(500);
    }

    [Test]
    public void Write_HandlesSpecialCharacters_AsValidJson()
    {
        var message = "quote \" newline \n unicode ☃ ą backslash \\";

        using var doc = WriteAndParse(LogLevel.Information, message); // throws if invalid JSON

        doc.RootElement.GetProperty("message").GetString().Should().Be(message);
    }
}

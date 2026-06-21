using D2ViewerEditor.Api.Middleware;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Middleware;

[TestFixture]
public class RequestObservabilityMiddlewareTests
{
    private static RequestObservabilityMiddleware Build(RequestDelegate next) =>
        new(next, NullLogger<RequestObservabilityMiddleware>.Instance);

    [Test]
    public async Task Sets_correlation_id_response_header_when_absent()
    {
        var context = new DefaultHttpContext();
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString()
            .Should().NotBeNullOrWhiteSpace();
        context.Items[RequestObservabilityMiddleware.CorrelationIdItemKey].Should().NotBeNull();
    }

    [Test]
    public async Task Echoes_incoming_correlation_id()
    {
        var context = new DefaultHttpContext();
        context.Request.Headers[RequestObservabilityMiddleware.CorrelationIdHeader] = "corr-abc";
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString()
            .Should().Be("corr-abc");
        context.Items[RequestObservabilityMiddleware.CorrelationIdItemKey].Should().Be("corr-abc");
    }

    [Test]
    public async Task Ignores_whitespace_correlation_id_and_generates_new_one()
    {
        var context = new DefaultHttpContext();
        context.Request.Headers[RequestObservabilityMiddleware.CorrelationIdHeader] = "   ";
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        var resolved = context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString();
        resolved.Should().NotBeNullOrWhiteSpace();
        resolved.Should().NotBe("   ");
    }

    [Test]
    public async Task Ignores_too_long_correlation_id_and_generates_new_one()
    {
        var tooLong = new string('a', 129);
        var context = new DefaultHttpContext();
        context.Request.Headers[RequestObservabilityMiddleware.CorrelationIdHeader] = tooLong;
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        var resolved = context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString();
        resolved.Should().NotBeNullOrWhiteSpace();
        resolved.Should().NotBe(tooLong);
    }
}

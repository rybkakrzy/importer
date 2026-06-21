using D2ServicesViewerEditor.Api.Middleware;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ServicesViewerEditor.Api.UnitTests.Middleware;

[TestFixture]
public class RequestObservabilityMiddlewareTests
{
    private static RequestObservabilityMiddleware Build(RequestDelegate next) =>
        new(next, NullLogger<RequestObservabilityMiddleware>.Instance);

    [Test]
    public async Task InvokeAsync_WhenHeaderMissing_GeneratesCorrelationIdAndSetsResponseHeader()
    {
        var context = new DefaultHttpContext();
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader]
            .ToString().Should().NotBeNullOrWhiteSpace();
        context.Items[RequestObservabilityMiddleware.CorrelationIdItemKey].Should().NotBeNull();
    }

    [Test]
    public async Task InvokeAsync_WhenIncomingHeaderPresent_EchoesCorrelationId()
    {
        var context = new DefaultHttpContext();
        context.Request.Headers[RequestObservabilityMiddleware.CorrelationIdHeader] = "corr-in";
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString().Should().Be("corr-in");
        context.Items[RequestObservabilityMiddleware.CorrelationIdItemKey].Should().Be("corr-in");
    }

    [Test]
    public async Task InvokeAsync_WhenHeaderTooLong_GeneratesNewCorrelationId()
    {
        var tooLong = new string('x', 129);
        var context = new DefaultHttpContext();
        context.Request.Headers[RequestObservabilityMiddleware.CorrelationIdHeader] = tooLong;
        var middleware = Build(_ => Task.CompletedTask);

        await middleware.InvokeAsync(context);

        context.Response.Headers[RequestObservabilityMiddleware.CorrelationIdHeader].ToString().Should().NotBe(tooLong);
    }
}

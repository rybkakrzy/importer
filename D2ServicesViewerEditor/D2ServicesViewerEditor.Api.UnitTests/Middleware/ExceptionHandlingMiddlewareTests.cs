using System.Net;
using System.Text.Json;
using D2ServicesViewerEditor.Api.Middleware;
using FluentAssertions;
using FluentValidation;
using FluentValidation.Results;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ServicesViewerEditor.Api.UnitTests.Middleware;

[TestFixture]
public class ExceptionHandlingMiddlewareTests
{
    private static async Task<(int statusCode, string contentType, JsonDocument body)> InvokeWith(Exception toThrow, string? correlationId = null)
    {
        var context = new DefaultHttpContext();
        var responseBody = new MemoryStream();
        context.Response.Body = responseBody;
        if (correlationId is not null)
            context.Items[RequestObservabilityMiddleware.CorrelationIdItemKey] = correlationId;

        RequestDelegate next = _ => throw toThrow;
        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        await middleware.InvokeAsync(context);

        responseBody.Position = 0;
        using var reader = new StreamReader(responseBody);
        var text = await reader.ReadToEndAsync();
        return (context.Response.StatusCode, context.Response.ContentType ?? string.Empty, JsonDocument.Parse(text));
    }

    [Test]
    public async Task InvokeAsync_NoException_PassesThrough()
    {
        var context = new DefaultHttpContext();
        var called = false;
        RequestDelegate next = _ =>
        {
            called = true;
            return Task.CompletedTask;
        };

        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        await middleware.InvokeAsync(context);

        called.Should().BeTrue();
        context.Response.StatusCode.Should().Be(200);
    }

    [Test]
    public async Task InvokeAsync_ValidationException_Returns400ProblemDetails()
    {
        var ex = new ValidationException(new[]
        {
            new ValidationFailure("Name", "Required")
        });

        var (statusCode, contentType, body) = await InvokeWith(ex);

        statusCode.Should().Be((int)HttpStatusCode.BadRequest);
        contentType.Should().Be("application/problem+json");
        body.RootElement.GetProperty("title").GetString().Should().Be("Validation Error");
        body.RootElement.GetProperty("errors").GetProperty("Name").EnumerateArray().Select(x => x.GetString())
            .Should().Contain("Required");
    }

    [Test]
    public async Task InvokeAsync_ArgumentException_Returns400WithMessage()
    {
        var (statusCode, _, body) = await InvokeWith(new ArgumentException("bad input"));

        statusCode.Should().Be((int)HttpStatusCode.BadRequest);
        body.RootElement.GetProperty("detail").GetString().Should().Be("bad input");
    }

    [Test]
    public async Task InvokeAsync_UnauthorizedAccessException_Returns401()
    {
        var (statusCode, _, body) = await InvokeWith(new UnauthorizedAccessException());

        statusCode.Should().Be((int)HttpStatusCode.Unauthorized);
        body.RootElement.GetProperty("title").GetString().Should().Be("Unauthorized");
    }

    [Test]
    public async Task InvokeAsync_UnexpectedException_Returns500AndHidesInternalMessage()
    {
        var (statusCode, _, body) = await InvokeWith(new InvalidOperationException("internal details"));

        statusCode.Should().Be((int)HttpStatusCode.InternalServerError);
        body.RootElement.GetProperty("detail").GetString().Should().NotContain("internal details");
    }

    [Test]
    public async Task InvokeAsync_WhenCorrelationIdExists_AddsItToProblemDetails()
    {
        var (_, _, body) = await InvokeWith(new ArgumentException("bad"), correlationId: "corr-007");

        body.RootElement.GetProperty("correlationId").GetString().Should().Be("corr-007");
    }
}

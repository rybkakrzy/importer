using System.Net;
using System.Text.Json;
using FluentValidation;
using Microsoft.AspNetCore.Mvc;

namespace D2ServicesViewerEditor.Api.Middleware;

public class ExceptionHandlingMiddleware
{
    private readonly RequestDelegate _next;
    private readonly ILogger<ExceptionHandlingMiddleware> _logger;

    public ExceptionHandlingMiddleware(RequestDelegate next, ILogger<ExceptionHandlingMiddleware> logger)
    {
        _next = next;
        _logger = logger;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        try
        {
            await _next(context);
        }
        catch (Exception ex)
        {
            await HandleExceptionAsync(context, ex);
        }
    }

    private async Task HandleExceptionAsync(HttpContext context, Exception exception)
    {
        var (statusCode, problemDetails) = exception switch
        {
            ValidationException validationException => HandleValidationException(validationException),
            ArgumentException argumentException => HandleArgumentException(argumentException),
            KeyNotFoundException notFoundException => HandleNotFoundException(notFoundException),
            UnauthorizedAccessException => HandleUnauthorizedException(),
            OperationCanceledException => HandleCancelledException(),
            _ => HandleUnexpectedException(exception)
        };

        _logger.LogError(exception, "Wystąpił wyjątek: {Message}", exception.Message);

        if (context.Items.TryGetValue(RequestObservabilityMiddleware.CorrelationIdItemKey, out var correlationId)
            && correlationId is string id)
        {
            problemDetails.Extensions["correlationId"] = id;
        }

        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/problem+json";

        var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
        await context.Response.WriteAsync(
            JsonSerializer.Serialize(problemDetails, problemDetails.GetType(), options));
    }

    private static (int, ProblemDetails) HandleValidationException(ValidationException exception)
    {
        var errors = exception.Errors
            .GroupBy(e => e.PropertyName)
            .ToDictionary(
                g => g.Key,
                g => g.Select(e => e.ErrorMessage).ToArray()
            );

        return ((int)HttpStatusCode.BadRequest, new ValidationProblemDetails(errors)
        {
            Title = "Validation Error",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = "One or more validation errors occurred.",
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.1"
        });
    }

    private static (int, ProblemDetails) HandleArgumentException(ArgumentException exception)
    {
        return ((int)HttpStatusCode.BadRequest, new ProblemDetails
        {
            Title = "Bad Request",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = exception.Message,
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.1"
        });
    }

    private static (int, ProblemDetails) HandleNotFoundException(KeyNotFoundException exception)
    {
        return ((int)HttpStatusCode.NotFound, new ProblemDetails
        {
            Title = "Not Found",
            Status = (int)HttpStatusCode.NotFound,
            Detail = exception.Message,
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.4"
        });
    }

    private static (int, ProblemDetails) HandleUnauthorizedException()
    {
        return ((int)HttpStatusCode.Unauthorized, new ProblemDetails
        {
            Title = "Unauthorized",
            Status = (int)HttpStatusCode.Unauthorized,
            Detail = "Invalid or missing API key",
            Type = "https://tools.ietf.org/html/rfc7235#section-3.1"
        });
    }

    private static (int, ProblemDetails) HandleCancelledException()
    {
        return ((int)HttpStatusCode.BadRequest, new ProblemDetails
        {
            Title = "Request Cancelled",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = "The request was cancelled by the client"
        });
    }

    private static (int, ProblemDetails) HandleUnexpectedException(Exception exception)
    {
        return ((int)HttpStatusCode.InternalServerError, new ProblemDetails
        {
            Title = "Internal Server Error",
            Status = (int)HttpStatusCode.InternalServerError,
            Detail = "An unexpected error occurred. Please contact support if the problem persists.",
            Type = "https://tools.ietf.org/html/rfc7231#section-6.6.1"
        });
    }
}

using System.Net;
using System.Text.Json;
using FluentValidation;
using Microsoft.AspNetCore.Mvc;

namespace D2Tools.Api.Middleware;

/// <summary>
/// Middleware do globalnej obsługi wyjątków — mapuje wyjątki na ProblemDetails (RFC 7807)
/// </summary>
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
            OperationCanceledException => HandleCancelledException(),
            _ => HandleUnexpectedException(exception)
        };

        _logger.LogError(exception, "Wystąpił wyjątek: {Message}", exception.Message);

        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/problem+json";

        var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
        await context.Response.WriteAsync(JsonSerializer.Serialize(problemDetails, options));
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
            Title = "Błąd walidacji",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = "Jeden lub więcej błędów walidacji",
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.1"
        });
    }

    private static (int, ProblemDetails) HandleArgumentException(ArgumentException exception)
    {
        return ((int)HttpStatusCode.BadRequest, new ProblemDetails
        {
            Title = "Nieprawidłowe żądanie",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = exception.Message,
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.1"
        });
    }

    private static (int, ProblemDetails) HandleNotFoundException(KeyNotFoundException exception)
    {
        return ((int)HttpStatusCode.NotFound, new ProblemDetails
        {
            Title = "Nie znaleziono",
            Status = (int)HttpStatusCode.NotFound,
            Detail = exception.Message,
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.4"
        });
    }

    private static (int, ProblemDetails) HandleCancelledException()
    {
        return ((int)HttpStatusCode.BadRequest, new ProblemDetails
        {
            Title = "Żądanie anulowane",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = "Żądanie zostało anulowane przez klienta"
        });
    }

    private (int, ProblemDetails) HandleUnexpectedException(Exception exception)
    {
        _logger.LogCritical(exception, "Nieobsługiwany wyjątek");

        return ((int)HttpStatusCode.InternalServerError, new ProblemDetails
        {
            Title = "Błąd serwera",
            Status = (int)HttpStatusCode.InternalServerError,
            Detail = "Wystąpił nieoczekiwany błąd. Spróbuj ponownie później.",
            Type = "https://tools.ietf.org/html/rfc7231#section-6.6.1"
        });
    }
}

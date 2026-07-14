using System.Net;
using System.Text.Json;
using D2ViewerEditor.Application.Common;
using FluentValidation;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Middleware;

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

        // Surface the correlation id to the client so a failed call can be traced in Kibana.
        if (context.Items.TryGetValue(RequestObservabilityMiddleware.CorrelationIdItemKey, out var correlationId)
            && correlationId is string id)
        {
            problemDetails.Extensions["correlationId"] = id;
        }

        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/problem+json";

        var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
        // Serialize via the runtime type so derived members (e.g. ValidationProblemDetails.Errors)
        // are emitted — System.Text.Json otherwise honours the static type and would drop them.
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

        var problemDetails = new ValidationProblemDetails(errors)
        {
            Title = "Błąd walidacji",
            Status = (int)HttpStatusCode.BadRequest,
            Detail = "Jeden lub więcej błędów walidacji",
            Type = "https://tools.ietf.org/html/rfc7231#section-6.5.1"
        };

        // Wystaw stabilny kod maszynowy (np. DOCUMENT_CONTENT_EMPTY) obok listy błędów, aby GUI mogło
        // rozpoznać przypadek domenowy niezależnie od treści komunikatu — tak samo jak w odpowiedziach
        // { code, error }. Ignoruje domyślne kody FluentValidation, bierze pierwszy znany kod domenowy.
        var domainCode = exception.Errors
            .Select(e => e.ErrorCode)
            .FirstOrDefault(ErrorCodes.IsKnown);
        if (domainCode is not null)
            problemDetails.Extensions["code"] = domainCode;

        return ((int)HttpStatusCode.BadRequest, problemDetails);
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

using System.Diagnostics;
using Serilog.Context;

namespace D2ServicesViewerEditor.Api.Middleware;

/// <summary>
/// Resolves and propagates correlation id for each request, then opens scope/context properties
/// so every log line can be correlated in ELK/GCP.
/// </summary>
public sealed class RequestObservabilityMiddleware
{
    public const string CorrelationIdHeader = "X-Correlation-ID";
    public const string CorrelationIdItemKey = "CorrelationId";

    private readonly RequestDelegate _next;
    private readonly ILogger<RequestObservabilityMiddleware> _logger;

    public RequestObservabilityMiddleware(RequestDelegate next, ILogger<RequestObservabilityMiddleware> logger)
    {
        _next = next;
        _logger = logger;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var correlationId = ResolveCorrelationId(context);
        context.Items[CorrelationIdItemKey] = correlationId;
        context.Response.Headers[CorrelationIdHeader] = correlationId;

        using var scope = _logger.BeginScope(new Dictionary<string, object>
        {
            ["correlationId"] = correlationId,
            ["requestId"] = context.TraceIdentifier
        });

        using var correlationProperty = LogContext.PushProperty("correlationId", correlationId);
        using var requestProperty = LogContext.PushProperty("requestId", context.TraceIdentifier);

        await _next(context);
    }

    private static string ResolveCorrelationId(HttpContext context)
    {
        if (context.Request.Headers.TryGetValue(CorrelationIdHeader, out var header))
        {
            var value = header.ToString();
            if (!string.IsNullOrWhiteSpace(value) && value.Length <= 128)
                return value;
        }

        return Activity.Current?.TraceId.ToString() ?? Guid.NewGuid().ToString("N");
    }
}

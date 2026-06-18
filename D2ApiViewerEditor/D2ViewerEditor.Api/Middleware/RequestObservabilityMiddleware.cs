using System.Diagnostics;
using System.Security.Claims;

namespace D2ViewerEditor.Api.Middleware;

/// <summary>
/// Per-request observability: resolves/propagates a correlation id (request header or generated),
/// opens a logging scope so every log of the request carries it, and emits one structured access log
/// on completion (method, path, status, elapsed, userId). Level scales with the status code so 4xx/5xx
/// stand out in Kibana. No request/response bodies are logged.
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

        var stopwatch = Stopwatch.StartNew();
        try
        {
            await _next(context);
        }
        finally
        {
            stopwatch.Stop();
            var status = context.Response.StatusCode;
            var level = status >= 500 ? LogLevel.Error
                : status >= 400 ? LogLevel.Warning
                : LogLevel.Information;

            // userId resolved here (after authentication ran). Subject/oid only — never name/email (PII).
            _logger.Log(level,
                "HTTP {httpMethod} {httpPath} responded {statusCode} in {elapsedMs} ms (user={userId})",
                context.Request.Method,
                context.Request.Path.Value,
                status,
                stopwatch.ElapsedMilliseconds,
                ResolveUserId(context));
        }
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

    private static string ResolveUserId(HttpContext context)
    {
        var user = context.User;
        if (user?.Identity?.IsAuthenticated != true)
            return "anonymous";

        return user.FindFirst("oid")?.Value
               ?? user.FindFirst(ClaimTypes.NameIdentifier)?.Value
               ?? user.FindFirst("sub")?.Value
               ?? "unknown";
    }
}

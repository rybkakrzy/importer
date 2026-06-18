using D2ViewerEditor.Api.Middleware;

namespace D2ViewerEditor.Api.Extensions;

public static class MiddlewareExtensions
{
    public static IApplicationBuilder UseExceptionHandlingMiddleware(this IApplicationBuilder app)
    {
        return app.UseMiddleware<ExceptionHandlingMiddleware>();
    }

    public static IApplicationBuilder UseRequestObservability(this IApplicationBuilder app)
    {
        return app.UseMiddleware<RequestObservabilityMiddleware>();
    }
}

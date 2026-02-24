using D2Tools.Api.Middleware;

namespace D2Tools.Api.Extensions;

public static class MiddlewareExtensions
{
    public static IApplicationBuilder UseExceptionHandlingMiddleware(this IApplicationBuilder app)
    {
        return app.UseMiddleware<ExceptionHandlingMiddleware>();
    }
}

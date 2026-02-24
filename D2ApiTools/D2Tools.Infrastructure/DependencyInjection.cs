using D2Tools.Domain.Interfaces;
using D2Tools.Infrastructure.Services;
using Microsoft.Extensions.DependencyInjection;

namespace D2Tools.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IBarcodeGenerator, BarcodeGeneratorService>();
        services.AddScoped<IDocxToHtmlConverter, DocxToHtmlConverter>();
        services.AddScoped<IHtmlToDocxConverter, HtmlToDocxConverter>();
        services.AddScoped<IDigitalSignatureService, DigitalSignatureService>();

        return services;
    }
}

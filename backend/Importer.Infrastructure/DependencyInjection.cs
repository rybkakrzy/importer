using Importer.Domain.Interfaces;
using Importer.Infrastructure.Services;
using Microsoft.Extensions.DependencyInjection;

namespace Importer.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IBarcodeGenerator, BarcodeGeneratorService>();
        services.AddSingleton<IUploadSessionService, InMemoryUploadSessionService>();
        services.AddScoped<IDocxToHtmlConverter, DocxToHtmlConverter>();
        services.AddScoped<IHtmlToDocxConverter, HtmlToDocxConverter>();
        services.AddScoped<IDigitalSignatureService, DigitalSignatureService>();

        return services;
    }
}

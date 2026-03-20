using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Infrastructure.Persistence;
using D2ViewerEditor.Infrastructure.Persistence.Repositories;
using D2ViewerEditor.Infrastructure.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace D2ViewerEditor.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services, IConfiguration configuration)
    {
        // Document services
        services.AddSingleton<IBarcodeGenerator, BarcodeGeneratorService>();
        services.AddScoped<IDocxToHtmlConverter, DocxToHtmlConverter>();
        services.AddScoped<IHtmlToDocxConverter, HtmlToDocxConverter>();
        services.AddScoped<IDigitalSignatureService, DigitalSignatureService>();

        // Database
        var connectionString = configuration.GetConnectionString("DefaultConnection");
        if (!string.IsNullOrEmpty(connectionString))
        {
            services.AddDbContext<DocumentDbContext>(options =>
                options.UseNpgsql(connectionString));

            services.AddScoped<IDocumentRepository, DocumentRepository>();
        }

        return services;
    }
}

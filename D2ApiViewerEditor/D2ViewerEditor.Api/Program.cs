using D2ViewerEditor.Api.Extensions;
using D2ViewerEditor.Application;
using D2ViewerEditor.Infrastructure;
using Microsoft.OpenApi.Models;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// Załaduj secrets specyficzne dla środowiska (nie commituj do repo!)
var env = builder.Environment.EnvironmentName;
builder.Configuration.AddJsonFile($"appsettings.{env}.secrets.json", optional: true, reloadOnChange: false);

// Strukturalne logi JSON z polem `severity` → Cloud Logging pokazuje ERROR/CRITICAL poprawnie
// (zamiast wszystkiego jako INFO). Lokalnie (Development) zostaje czytelny formatter.
builder.AddGcpStructuredLogging();

// Add Architecture layers
builder.Services.AddApplication();
builder.Services.AddInfrastructure(builder.Configuration);

// Add API services
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "D2 Viewer Editor API",
        Version = "v1",
        Description = """
            REST API do zarządzania dokumentami DOCX — otwieranie, edycja, zapis wersji,
            podpisywanie cyfrowe oraz pobieranie historii wersji.
            """,
        Contact = new OpenApiContact
        {
            Name = "D2 Team",
            Email = "d2team@example.com"
        }
    });

    // Dołącz XML komentarze z kodu
    var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    options.IncludeXmlComments(xmlPath);

    // Grupowanie endpointów po tagach (controller name)
    options.TagActionsBy(api =>
    {
        api.ActionDescriptor.RouteValues.TryGetValue("controller", out var controller);
        return [api.GroupName ?? controller ?? "Default"];
    });
});

// Add CORS — originy z konfiguracji
var allowedOrigins = builder.Configuration.GetSection("Cors:AllowedOrigins").Get<string[]>() ?? ["http://localhost:4200"];
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAngularApp",
        policy =>
        {
            policy.WithOrigins(allowedOrigins)
                  .AllowAnyHeader()
                  .AllowAnyMethod();
        });
});

var app = builder.Build();

// Middleware pipeline — order matters
app.UseExceptionHandlingMiddleware();

// CORS musi być przed MapControllers
app.UseCors("AllowAngularApp");

// Swagger / Scalar — włączany per środowisko
var swaggerEnabled = app.Configuration.GetValue<bool>("Swagger:Enabled");
if (swaggerEnabled)
{
    app.UseSwagger();

    // Klasyczny Swagger UI pod /swagger
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "D2 Viewer Editor API v1");
        c.DocumentTitle = "D2 Viewer Editor API";
        c.DisplayRequestDuration();
        c.EnableDeepLinking();
        c.EnableFilter();
    });
}

app.MapControllers();

app.Run();

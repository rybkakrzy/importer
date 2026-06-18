using D2ViewerEditor.Api.Extensions;
using D2ViewerEditor.Api.Security;
using D2ViewerEditor.Application;
using D2ViewerEditor.Infrastructure;
using Microsoft.OpenApi.Models;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// Załaduj secrets specyficzne dla środowiska (nie commituj do repo!)
var env = builder.Environment.EnvironmentName;
builder.Configuration.AddJsonFile($"appsettings.{env}.secrets.json", optional: true, reloadOnChange: false);

// Large documents (DOCX/PDF carried as base64) exceed Kestrel's ~30 MB default → 413 on upload/finish.
// Raise the global cap (configurable) and relax slow-client rate limits for big transfers. Per-action
// [RequestSizeLimit] still applies where stricter; the Server header is dropped (minor hardening).
var maxRequestBodyBytes = builder.Configuration.GetValue<long?>("Kestrel:MaxRequestBodySizeBytes") ?? 150_000_000;
builder.WebHost.ConfigureKestrel(options =>
{
    options.AddServerHeader = false;
    options.Limits.MaxRequestBodySize = maxRequestBodyBytes;
    options.Limits.MinRequestBodyDataRate = null;
    options.Limits.MinResponseDataRate = null;
});

// Strukturalne logi JSON z polem `severity` → Cloud Logging pokazuje ERROR/CRITICAL poprawnie
// (zamiast wszystkiego jako INFO). Lokalnie (Development) zostaje czytelny formatter.
builder.AddGcpStructuredLogging();

// GCP Secret Manager (Qutas): inject AzureAd:ClientSecret from a managed secret when enabled.
// No-op locally (disabled by default) — never fails startup.
builder.AddEntraSecretFromGcp();

// Add Architecture layers
builder.Services.AddApplication();
builder.Services.AddInfrastructure(builder.Configuration);

// ── Authentication & authorization ──────────────────────────────────────────
// LOKALNY DEV: gdy `Auth:DevBypass=true` (i NIE Production) — pomijamy całkowicie Entra i używamy
// sztucznego, uwierzytelnionego principala (DevAuthHandler), żeby API działało bez tenanta. Domyślnie
// wyłączone; włączasz w LOKALNYM `appsettings.{ENV}.secrets.json` (poza repo). Nigdy na serwerze.
var devAuthBypass = builder.Configuration.GetValue<bool>("Auth:DevBypass")
                    && !builder.Environment.IsProduction();

if (devAuthBypass)
{
    builder.Services.AddDevBypassAuthentication();
}
else
{
    builder.Services.AddEntraIdAuthentication(builder.Configuration, builder.Environment);
}

// Role→resource authorization map (Qutas ResourcesProvider). Reads role names from AzureAdOptions
// (defaults Operator/Administrator in dev where the section isn't bound). Backend = source of truth
// for the Angular resource guard (GET /api/identity/resources).
builder.Services.AddScoped<ResourcesProvider>();

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

// Behind a TLS-terminating proxy (the app sees http) force the request scheme to https, so generated
// URLs and OIDC redirect URIs are https (D2WebCore pattern).
app.Use((context, next) =>
{
    context.Request.Scheme = "https";
    return next(context);
});

// Middleware pipeline — order matters. Observability is outermost so its correlation-id scope wraps
// every log (including exceptions) and the access log captures the final status code.
app.UseRequestObservability();
app.UseExceptionHandlingMiddleware();

// CORS musi być przed MapControllers
app.UseCors("AllowAngularApp");

// AuthN/AuthZ — po CORS, przed MapControllers
app.UseAuthentication();
app.UseAuthorization();

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

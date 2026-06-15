using D2ViewerEditor.Api.Extensions;
using D2ViewerEditor.Api.Security;
using D2ViewerEditor.Application;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Infrastructure;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Identity.Web;
using Microsoft.OpenApi.Models;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// Załaduj secrets specyficzne dla środowiska (nie commituj do repo!)
var env = builder.Environment.EnvironmentName;
builder.Configuration.AddJsonFile($"appsettings.{env}.secrets.json", optional: true, reloadOnChange: false);

// Strukturalne logi JSON z polem `severity` → Cloud Logging pokazuje ERROR/CRITICAL poprawnie
// (zamiast wszystkiego jako INFO). Lokalnie (Development) zostaje czytelny formatter.
builder.AddGcpStructuredLogging();

// GCP Secret Manager (Doc2): inject AzureAd:ClientSecret from a managed secret when enabled.
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
    builder.Services.AddAuthentication(DevAuthHandler.SchemeName)
        .AddScheme<AuthenticationSchemeOptions, DevAuthHandler>(DevAuthHandler.SchemeName, _ => { });
    builder.Services.AddAuthorization(options =>
    {
        // W trybie dev autoryzacja jest „otwarta" — każdy (sztuczny) użytkownik przechodzi, też admin.
        options.AddPolicy(AuthorizationPolicies.RequireAppEmployee, p => p.RequireAuthenticatedUser());
        options.AddPolicy(AuthorizationPolicies.RequireAppAdmin, p => p.RequireAuthenticatedUser());
    });
    builder.Services.AddHttpContextAccessor();
    builder.Services.AddScoped<ICurrentUserProvider, HttpHeaderCurrentUserProvider>();
    builder.Services.AddSingleton<IGraphUserService, DisabledGraphUserService>();
}
else
{
// ── Entra ID authentication & authorization ─────────────────────────────────
var azureAd = new AzureAdOptions();
builder.Configuration.GetSection(AzureAdOptions.SectionName).Bind(azureAd);
builder.Services.Configure<AzureAdOptions>(builder.Configuration.GetSection(AzureAdOptions.SectionName));

var keycloak = new KeycloakOptions();
builder.Configuration.GetSection(KeycloakOptions.SectionName).Bind(keycloak);

// Authentication: Entra ID access tokens validated via Microsoft.Identity.Web (Doc2/D2WebCore
// pattern) instead of raw JwtBearer. When Keycloak is enabled, a second JWT scheme runs alongside
// Entra and the active scheme is chosen per request by the token issuer (legacy coexistence).
if (keycloak.Enabled && !string.IsNullOrWhiteSpace(keycloak.Authority))
{
    // AddMicrosoftIdentityWebApi returns a specialized builder, so call each Add* on the base
    // AuthenticationBuilder as separate statements rather than chaining.
    var authBuilder = builder.Services.AddAuthentication(AuthSchemes.EntraOrKeycloak);
    authBuilder.AddPolicyScheme(AuthSchemes.EntraOrKeycloak, "Entra ID or Keycloak", options =>
    {
        options.ForwardDefaultSelector = ctx => AuthSchemes.SelectByIssuer(ctx, keycloak.Authority);
    });
    authBuilder.AddMicrosoftIdentityWebApi(builder.Configuration, AzureAdOptions.SectionName, AuthSchemes.Entra);
    authBuilder.AddJwtBearer(AuthSchemes.Keycloak, options =>
    {
        options.Authority = keycloak.Authority;
        options.Audience = keycloak.Audience;
        options.TokenValidationParameters.RoleClaimType = "roles";
        options.TokenValidationParameters.NameClaimType = "name";
    });
}
else
{
    builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
        .AddMicrosoftIdentityWebApi(builder.Configuration, AzureAdOptions.SectionName);
}

// Both native Entra app-role claims AND group→role mappings surface in the "roles" claim, so
// IsInRole / RequireRole work uniformly. PostConfigure (registered after AddMicrosoftIdentityWebApi)
// guarantees this wins over Identity.Web's own JwtBearer configuration.
builder.Services.PostConfigure<JwtBearerOptions>(JwtBearerDefaults.AuthenticationScheme, options =>
{
    options.TokenValidationParameters.RoleClaimType = "roles";
    options.TokenValidationParameters.NameClaimType = "name";
});

// Group→role mapping (Doc2): map the Entra "groups" claim onto application role claims.
builder.Services.Configure<RolesOptions>(builder.Configuration.GetSection(RolesOptions.SectionName));
builder.Services.AddScoped<IClaimsTransformation, Doc2ClaimsTransformer>();

builder.Services.AddAuthorization(options =>
{
    options.AddPolicy(AuthorizationPolicies.RequireAppEmployee, policy =>
        policy.RequireRole(azureAd.EmployeeRole, azureAd.AdminRole)); // admin also has app access
    options.AddPolicy(AuthorizationPolicies.RequireAppAdmin, policy =>
        policy.RequireRole(azureAd.AdminRole));
});

// Identity for document access control: production reads the Entra ID access token claims.
builder.Services.AddHttpContextAccessor();
builder.Services.AddScoped<ICurrentUserProvider, ClaimsCurrentUserProvider>();

// Microsoft Graph (Doc2): app-only user lookup. Active only when a client secret is available
// (from GCP Secret Manager) + ClientId/TenantId set; otherwise a disabled no-op is used so the
// API runs locally without Graph credentials.
if (!string.IsNullOrWhiteSpace(azureAd.ClientSecret)
    && !string.IsNullOrWhiteSpace(azureAd.ClientId)
    && !string.IsNullOrWhiteSpace(azureAd.TenantId))
{
    builder.Services.AddSingleton<IGraphUserService, GraphUserService>();
}
else
{
    builder.Services.AddSingleton<IGraphUserService, DisabledGraphUserService>();
}
}

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

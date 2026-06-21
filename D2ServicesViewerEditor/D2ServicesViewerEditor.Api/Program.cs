using D2ServicesViewerEditor.Api.Extensions;
using D2ServicesViewerEditor.Api.Logging;
using D2ViewerEditor.Application;
using D2ViewerEditor.Infrastructure;
using Microsoft.OpenApi.Models;
using Serilog;
using System.Reflection;

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .WriteTo.File("logs/d2services-.log", rollingInterval: RollingInterval.Day)
    .CreateBootstrapLogger();

try
{
    var builder = WebApplication.CreateBuilder(args);

    builder.Host.UseSerilog((context, services, configuration) =>
    {
        configuration
            .ReadFrom.Configuration(context.Configuration)
            .ReadFrom.Services(services)
            .Enrich.FromLogContext()
            .WriteTo.File("logs/d2services-.log", rollingInterval: RollingInterval.Day);

        // Konsola: w GCP/kontenerze strukturalny JSON (pole `severity` → ERROR/CRITICAL w Logs
        // Explorer); lokalnie (Development) czytelny tekst. Sterowanie configiem `Logging:UseGcpFormat`.
        var useGcp = context.Configuration.GetValue<bool?>("Logging:UseGcpFormat")
                     ?? !context.HostingEnvironment.IsDevelopment();
        if (useGcp)
            configuration.WriteTo.Console(new GcpJsonSerilogFormatter(
                context.HostingEnvironment.ApplicationName,
                context.HostingEnvironment.EnvironmentName));
        else
            configuration.WriteTo.Console();
    });

    builder.Configuration.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: true)
        .AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: true)
        .AddJsonFile("appsettings.secrets.json", optional: true, reloadOnChange: true);

    builder.Services.AddApplication();
    builder.Services.AddInfrastructure(builder.Configuration);

    // External API authenticates app-to-app (not per user). The shared Application layer
    // registers IDocumentAccessGuard (→ ICurrentUserProvider); provide a no-user system
    // provider so DI is valid. Ingest does not perform per-user document access checks.
    builder.Services.AddScoped<D2ViewerEditor.Application.Common.Security.ICurrentUserProvider,
        D2ViewerEditor.Application.Common.Security.SystemCurrentUserProvider>();

    // Limity uploadu: validator dopuszcza pliki do 100 MB — Kestrel i parser multipart
    // muszą zezwalać na co najmniej tyle (z drobnym zapasem na nagłówki i metadane).
    const long MaxUploadBytes = 110L * 1024 * 1024;
    builder.WebHost.ConfigureKestrel(options =>
    {
        options.Limits.MaxRequestBodySize = MaxUploadBytes;
    });
    builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
    {
        options.MultipartBodyLengthLimit = MaxUploadBytes;
        options.ValueLengthLimit = int.MaxValue;
    });

    builder.Services.AddControllers()
        .AddJsonOptions(options =>
        {
            // Enumy serializujemy jako nazwy (nie liczby) — dzięki temu Swashbuckle renderuje
            // w Swaggerze listę dozwolonych wartości (np. Saved/Editing/Sending/...), a nie
            // gołe "string". Format po drucie pozostaje stringowy, więc kontrakt bez zmian.
            options.JsonSerializerOptions.Converters.Add(
                new System.Text.Json.Serialization.JsonStringEnumConverter());
        });
    builder.Services.AddEndpointsApiExplorer();

    builder.Services.AddSwaggerGen(options =>
    {
        options.SwaggerDoc("v1", new OpenApiInfo
        {
            Title = "D2 Services Viewer Editor API",
            Version = "v1",
            Description = """
                REST API dla zewnętrznych systemów i integracji.
                Uwaga: To API jest przeznaczone dla aplikacji zewnętrznych, nie dla frontendu!

                Wymaga autoryzacji poprzez API Key lub OAuth2.
                """,
            Contact = new OpenApiContact
            {
                Name = "D2 Integration Team",
                Email = "integration@d2.example.com"
            }
        });

        var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
        var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);

        if (File.Exists(xmlPath))
        {
            options.IncludeXmlComments(xmlPath);
        }

        options.AddSecurityDefinition("ApiKey", new OpenApiSecurityScheme
        {
            Description = "API Key needed to access the endpoints. X-Api-Key: YOUR_API_KEY",
            In = ParameterLocation.Header,
            Name = "X-Api-Key",
            Type = SecuritySchemeType.ApiKey
        });

        options.AddSecurityRequirement(new OpenApiSecurityRequirement
        {
            {
                new OpenApiSecurityScheme
                {
                    Reference = new OpenApiReference
                    {
                        Type = ReferenceType.SecurityScheme,
                        Id = "ApiKey"
                    }
                },
                Array.Empty<string>()
            }
        });
    });

    builder.Services.AddCors(options =>
    {
        options.AddPolicy("ExternalIntegrations", policy =>
        {
            var allowedOrigins = builder.Configuration.GetSection("Cors:AllowedOrigins").Get<string[]>()
                ?? Array.Empty<string>();

            if (allowedOrigins.Length > 0)
            {
                policy.WithOrigins(allowedOrigins)
                    .AllowAnyMethod()
                    .AllowAnyHeader();
            }
            else
            {
                policy.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader();
            }
        });
    });

    builder.Services.AddHealthChecks();

    var app = builder.Build();

    app.UseRequestObservability();
    app.UseSerilogRequestLogging(options =>
    {
        options.EnrichDiagnosticContext = (diagnostics, context) =>
        {
            diagnostics.Set("httpMethod", context.Request.Method);
            diagnostics.Set("httpPath", context.Request.Path.Value ?? string.Empty);
            diagnostics.Set("statusCode", context.Response.StatusCode);
            diagnostics.Set("requestId", context.TraceIdentifier);

            if (context.Items.TryGetValue(D2ServicesViewerEditor.Api.Middleware.RequestObservabilityMiddleware.CorrelationIdItemKey, out var id)
                && id is string correlationId)
            {
                diagnostics.Set("correlationId", correlationId);
            }
        };
    });
    app.UseExceptionHandlingMiddleware();

    var swaggerEnabled = builder.Configuration.GetValue<bool>("Swagger:Enabled", true);

    if (swaggerEnabled)
    {
        app.UseSwagger();
        app.UseSwaggerUI(c =>
        {
            c.SwaggerEndpoint("/swagger/v1/swagger.json", "D2 Services Viewer Editor API v1");
            c.RoutePrefix = "swagger";
        });
    }

    app.UseHttpsRedirection();

    app.UseCors("ExternalIntegrations");

    app.UseAuthorization();

    app.MapControllers();
    app.MapHealthChecks("/health");

    Log.Information("D2 Services Viewer Editor API starting on {Url}", builder.Configuration["Urls"]);

    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}

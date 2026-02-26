using D2Tools.Api.Extensions;
using D2Tools.Application;
using D2Tools.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Załaduj secrets specyficzne dla środowiska (nie commituj do repo!)
var env = builder.Environment.EnvironmentName;
builder.Configuration.AddJsonFile($"appsettings.{env}.secrets.json", optional: true, reloadOnChange: false);

// Add Architecture layers
builder.Services.AddApplication();
builder.Services.AddInfrastructure();

// Add API services
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

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

// Swagger — włączany per środowisko
var swaggerEnabled = app.Configuration.GetValue<bool>("Swagger:Enabled");
if (swaggerEnabled)
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("AllowAngularApp");

app.UseHttpsRedirection();

app.MapControllers();

app.Run();

using System.Reflection;
using D2ExampleExternalApp.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.OpenApi.Models;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "D2 Example External App",
        Version = "v1",
        Description = """
            Przykładowa aplikacja ZEWNĘTRZNA, symulująca integratora D2 ViewerEditor.

            Przepływ:
            1. POST /api/integration/upload — wgrywasz plik DOCX; aplikacja przesyła go do
               D2ServicesViewerEditor (POST /api/v1/document) z returnUrl wskazującym na własny
               endpoint callback.
            2. Użytkownik edytuje dokument w edytorze i klika „Zakończ i wyślij".
            3. D2 odsyła gotowy plik na POST /api/integration/callback (multipart/form-data:
               file, masterId, versionId, corporateKey + nagłówki Idempotency-Key, X-Content-SHA256).
            4. GET /api/integration/received — podgląd tego, co wróciło.
            """
    });

    var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
        options.IncludeXmlComments(xmlPath);
});

builder.Services.Configure<D2ServicesOptions>(builder.Configuration.GetSection(D2ServicesOptions.SectionName));
builder.Services.Configure<ExternalAppOptions>(builder.Configuration.GetSection(ExternalAppOptions.SectionName));

builder.Services.AddHttpClient(D2ServicesOptions.HttpClientName, (sp, client) =>
{
    var opts = sp.GetRequiredService<IOptions<D2ServicesOptions>>().Value;
    client.BaseAddress = new Uri(opts.BaseUrl);
    client.Timeout = TimeSpan.FromSeconds(100);
})
.ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
{
    // DEMO/DEV: D2Services lokalnie używa self-signed certu i robi HTTPS redirect — nie
    // weryfikujemy łańcucha, żeby przykład działał out-of-the-box. NIE używać tak na produkcji.
    ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator,
    AllowAutoRedirect = true
});

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "D2 Example External App v1");
    c.RoutePrefix = "swagger";
});

app.MapControllers();

app.Run();

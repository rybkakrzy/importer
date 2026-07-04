using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;

namespace D2ViewerEditor.Api.IntegrationTests;

/// <summary>
/// Boots the real API pipeline (real controllers, real <c>[Authorize]</c> attributes and real
/// RequireAppOperator / RequireAppAdmin policies) but swaps Entra JWT bearer for
/// <see cref="TestAuthHandler"/> as the default scheme, so access control can be driven per-request
/// via headers. The host is a hermetic slice: no database, no GCS, no delivery worker, and Entra is
/// wired for real (NOT dev-bypass) so the genuine role policies are exercised.
/// </summary>
public sealed class AuthTestWebApplicationFactory : WebApplicationFactory<Program>
{
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.UseEnvironment("Development");

        // Minimal-hosting gotcha: the factory's ConfigureAppConfiguration runs AFTER Program.cs has
        // already read builder.Configuration (in AddInfrastructure / AddEntraIdAuthentication). Host
        // settings applied via UseSetting are visible early, so we use those instead.
        //
        // Real Entra wiring (genuine RequireRole policies), never the dev bypass; deterministic role
        // names; dummy ClientId/TenantId so Microsoft.Identity.Web option validation passes (the real
        // schemes are never invoked — TestAuthHandler handles auth); no DB / GCS / delivery worker.
        builder.UseSetting("Auth:DevBypass", "false");
        builder.UseSetting("AzureAd:OperatorRole", "Operator");
        builder.UseSetting("AzureAd:AdminRole", "Administrator");
        builder.UseSetting("AzureAd:ClientId", "11111111-1111-1111-1111-111111111111");
        builder.UseSetting("AzureAd:TenantId", "22222222-2222-2222-2222-222222222222");
        builder.UseSetting("ConnectionStrings:DefaultConnection", "");
        builder.UseSetting("GoogleCloudStorage:BucketName", "");
        builder.UseSetting("DeliveryWorker:Enabled", "false");
        builder.UseSetting("Swagger:Enabled", "false");

        // The document handlers' dependencies (IDocumentRepository, IDocumentStorageService) are not
        // registered in this DB-less slice. Those handlers are never constructed — authorization
        // short-circuits protected requests and the one 200-path endpoint (templates) has no
        // dependencies — so skip the eager build-time DI validation that would otherwise reject it.
        builder.UseDefaultServiceProvider(options => options.ValidateOnBuild = false);

        builder.ConfigureTestServices(services =>
        {
            services.AddAuthentication()
                .AddScheme<AuthenticationSchemeOptions, TestAuthHandler>(TestAuthHandler.SchemeName, _ => { });

            // Point every default scheme at the test handler so [Authorize] authenticates through it
            // and the real Entra schemes are never materialized.
            services.Configure<AuthenticationOptions>(options =>
            {
                options.DefaultScheme = TestAuthHandler.SchemeName;
                options.DefaultAuthenticateScheme = TestAuthHandler.SchemeName;
                options.DefaultChallengeScheme = TestAuthHandler.SchemeName;
            });
        });
    }
}

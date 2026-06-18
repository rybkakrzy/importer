using D2ViewerEditor.Application.Common.Security;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Identity.Web;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Wiring for the API's authentication/authorization. Two mutually exclusive modes:
/// local dev bypass (<see cref="AddDevBypassAuthentication"/>) and Entra ID
/// (<see cref="AddEntraIdAuthentication"/>). Mirrors the D2WebCore <c>ConfigureAuthentication</c>
/// shape but stays a pure resource server (no server-side OIDC/cookie web-app — SPA does interactive
/// login via MSAL; see ADR-0011/0012).
/// </summary>
public static class ConfigureAuthentication
{
    /// <summary>
    /// LOCAL DEV ONLY: a fake authenticated principal (DevAuthHandler) so the API runs without a
    /// real tenant. Authorization is "open" — every (synthetic) user passes both policies.
    /// </summary>
    public static IServiceCollection AddDevBypassAuthentication(this IServiceCollection services)
    {
        services.AddAuthentication(DevAuthHandler.SchemeName)
            .AddScheme<AuthenticationSchemeOptions, DevAuthHandler>(DevAuthHandler.SchemeName, _ => { });

        services.AddAuthorization(options =>
        {
            options.AddPolicy(AuthorizationPolicies.RequireAppOperator, p => p.RequireAuthenticatedUser());
            options.AddPolicy(AuthorizationPolicies.RequireAppAdmin, p => p.RequireAuthenticatedUser());
        });

        services.AddHttpContextAccessor();
        services.AddScoped<ICurrentUserProvider, HttpHeaderCurrentUserProvider>();
        services.AddSingleton<IGraphUserService, DisabledGraphUserService>();
        return services;
    }

    /// <summary>OIDC web-app scheme (server-side interactive sign-in, code flow + cookie).</summary>
    public const string WebAppScheme = "MyAzureAdScheme";

    /// <summary>
    /// Entra ID via Microsoft.Identity.Web: WebApi (JWT bearer, default scheme) for API access tokens
    /// AND WebApp (<see cref="WebAppScheme"/>: OpenID Connect code flow + cookie + downstream token
    /// acquisition) for server-side interactive sign-in. Plus group→role mapping, role policies,
    /// claims-based identity and (when configured) Microsoft Graph. Behind a corporate proxy
    /// (AzureAd:Proxy:Url, skipped when IS_LOCAL_DEV=true) the backchannel is routed through it.
    /// </summary>
    public static IServiceCollection AddEntraIdAuthentication(
        this IServiceCollection services,
        IConfiguration configuration,
        IHostEnvironment environment)
    {
        var azureAd = new AzureAdOptions();
        configuration.GetSection(AzureAdOptions.SectionName).Bind(azureAd);
        services.Configure<AzureAdOptions>(configuration.GetSection(AzureAdOptions.SectionName));

        // Detailed PII in Microsoft.IdentityModel logs — a debugging aid for token-validation failures.
        // NEVER in Production (leaks PII to logs). Env names are "DEV"/etc. (not "Development"), so we
        // gate on !IsProduction() rather than IsDevelopment().
        Microsoft.IdentityModel.Logging.IdentityModelEventSource.ShowPII = !environment.IsProduction();

        var isLocalDev = Environment.GetEnvironmentVariable("IS_LOCAL_DEV") == "true";

        // Enterprise forward proxy (D2WebCore): off locally (IS_LOCAL_DEV), otherwise routed through
        // AzureAd:Proxy:Url with GCS bypassed (and set as the process-wide default proxy).
        HttpClientHandler CreateHttpHandler()
        {
            if (isLocalDev)
                return new HttpClientHandler { UseProxy = false };

            var businessProxy = EntraBackchannel.CreateProxy(azureAd);
            if (businessProxy is null)
                return new HttpClientHandler();

            HttpClient.DefaultProxy = businessProxy;
            return new HttpClientHandler { UseProxy = true, Proxy = businessProxy };
        }

        // WebApi: Entra access-token validation (default JWT bearer scheme) for the SPA's API calls.
        services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
            .AddMicrosoftIdentityWebApi(
                jwtOptions =>
                {
                    configuration.Bind(AzureAdOptions.SectionName, jwtOptions);
                    if (!isLocalDev)
                        jwtOptions.BackchannelHttpHandler = CreateHttpHandler();
                    jwtOptions.IncludeErrorDetails = !environment.IsProduction();
                },
                identityOptions => configuration.Bind(AzureAdOptions.SectionName, identityOptions));

        // WebApp: server-side interactive sign-in (OIDC code flow + cookie) + downstream token
        // acquisition (Microsoft Graph / OBO). Non-default scheme, used only on explicit challenge.
        services.AddAuthentication()
            .AddMicrosoftIdentityWebApp(
                oidcOptions =>
                {
                    oidcOptions.BackchannelHttpHandler = CreateHttpHandler();
                    oidcOptions.SignInScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                    oidcOptions.Instance = azureAd.Instance;
                    oidcOptions.TenantId = azureAd.TenantId;
                    oidcOptions.ClientId = azureAd.ClientId;
                    oidcOptions.ClientSecret = azureAd.ClientSecret;
                    oidcOptions.ResponseType = OpenIdConnectResponseType.Code;
                    oidcOptions.NonceCookie.SecurePolicy = CookieSecurePolicy.Always;
                    oidcOptions.CorrelationCookie.SecurePolicy = CookieSecurePolicy.Always;
                    oidcOptions.Scope.Add("offline_access");
                    oidcOptions.Scope.Add("email");
                    oidcOptions.TokenValidationParameters.ValidateIssuerSigningKey = true;
                },
                configureCookieAuthenticationOptions: null,
                openIdConnectScheme: WebAppScheme,
                cookieScheme: null)
            .EnableTokenAcquisitionToCallDownstreamApi()
            .AddInMemoryTokenCaches();

        // Both native Entra app-role claims AND group→role mappings surface in the "roles" claim, so
        // IsInRole / RequireRole work uniformly. PostConfigure (after AddMicrosoftIdentityWebApi)
        // guarantees this wins over Identity.Web's own JwtBearer configuration.
        services.PostConfigure<JwtBearerOptions>(JwtBearerDefaults.AuthenticationScheme, options =>
        {
            options.TokenValidationParameters.RoleClaimType = "roles";
            options.TokenValidationParameters.NameClaimType = "name";
        });

        // Group→role mapping (Qutas): map the Entra "groups" claim onto application role claims.
        services.Configure<RolesOptions>(configuration.GetSection(RolesOptions.SectionName));
        services.AddScoped<IClaimsTransformation, ClaimsTransformer>();

        services.AddAuthorization(options =>
        {
            options.AddPolicy(AuthorizationPolicies.RequireAppOperator, policy =>
                policy.RequireRole(azureAd.OperatorRole, azureAd.AdminRole)); // admin also has app access
            options.AddPolicy(AuthorizationPolicies.RequireAppAdmin, policy =>
                policy.RequireRole(azureAd.AdminRole));
        });

        services.AddHttpContextAccessor();
        services.AddScoped<ICurrentUserProvider, ClaimsCurrentUserProvider>();

        // Microsoft Graph (Qutas): app-only user lookup. Active only with a client secret (GCP Secret
        // Manager) + ClientId/TenantId; otherwise a disabled no-op so the API runs locally without it.
        if (!string.IsNullOrWhiteSpace(azureAd.ClientSecret)
            && !string.IsNullOrWhiteSpace(azureAd.ClientId)
            && !string.IsNullOrWhiteSpace(azureAd.TenantId))
        {
            services.AddSingleton<IGraphUserService, GraphUserService>();
        }
        else
        {
            services.AddSingleton<IGraphUserService, DisabledGraphUserService>();
        }

        return services;
    }
}

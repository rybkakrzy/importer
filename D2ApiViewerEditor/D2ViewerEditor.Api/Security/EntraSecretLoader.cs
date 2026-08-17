using Google.Cloud.SecretManager.V1;

namespace D2ViewerEditor.Api.Security;

/// <summary>
/// Loads the Entra client secret from GCP Secret Manager at startup and injects it into
/// configuration as <c>AzureAd:ClientSecret</c> (Doc2 pattern — keeps the secret out of repo and
/// out of appsettings). No-op (and never throws) when disabled or unconfigured, so local/dev
/// without GCP credentials keeps working with token-validation-only Entra.
/// </summary>
public static class EntraSecretLoader
{
    public static void AddEntraSecretFromGcp(this WebApplicationBuilder builder)
    {
        var options = new GcpSecretManagerOptions();
        builder.Configuration.GetSection(GcpSecretManagerOptions.SectionName).Bind(options);

        if (!options.Enabled
            || string.IsNullOrWhiteSpace(options.ProjectId)
            || string.IsNullOrWhiteSpace(options.EntraSecretName))
        {
            return;
        }

        var logger = LoggerFactory.Create(b => b.AddConsole()).CreateLogger("EntraSecretLoader");

        try
        {
            var client = SecretManagerServiceClient.Create();
            var versionName = new SecretVersionName(options.ProjectId, options.EntraSecretName, "latest");
            var result = client.AccessSecretVersion(versionName);
            var secret = result.Payload.Data.ToStringUtf8();

            if (!string.IsNullOrWhiteSpace(secret))
            {
                builder.Configuration["AzureAd:ClientSecret"] = secret;
                logger.LogInformation(
                    "Entra client secret loaded from GCP Secret Manager ({Secret}).", options.EntraSecretName);
            }
        }
        catch (Exception ex)
        {
            // Never fail startup on secret retrieval — the API still validates tokens without a
            // secret; only downstream Graph calls (which need the secret) are unavailable.
            logger.LogError(ex,
                "Could not load Entra client secret from GCP Secret Manager ({Secret}). " +
                "Token validation continues; Microsoft Graph features are disabled.",
                options.EntraSecretName);
        }
    }
}

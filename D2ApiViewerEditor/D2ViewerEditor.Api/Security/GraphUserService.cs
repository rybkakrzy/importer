using Azure.Identity;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace D2ViewerEditor.Api.Security;

/// <summary>User profile fields read from Microsoft Graph (Doc2 <c>GraphUserService</c>).</summary>
public sealed record GraphUserInfo(
    string? Id,
    string? UserPrincipalName,
    string? Mail,
    string? DisplayName,
    string? GivenName,
    string? Surname,
    string? JobTitle,
    string? Department);

/// <summary>Looks up Entra ID users via Microsoft Graph. One implementation calls Graph
/// (app-only), the other is a disabled no-op for environments without a client secret.</summary>
public interface IGraphUserService
{
    Task<GraphUserInfo?> FindUserAsync(string query, CancellationToken cancellationToken = default);
}

/// <summary>
/// Microsoft Graph user lookup using app-only (client-credentials) auth — Doc2 pattern.
/// Registered only when AzureAd ClientId/TenantId/ClientSecret are present (secret comes from
/// GCP Secret Manager); otherwise <see cref="DisabledGraphUserService"/> is used.
/// </summary>
public sealed class GraphUserService : IGraphUserService
{
    private static readonly string[] Scopes = { "https://graph.microsoft.com/.default" };
    private static readonly string[] SelectFields =
    {
        "id", "userPrincipalName", "mail", "displayName", "givenName", "surname", "jobTitle", "department"
    };

    private readonly GraphServiceClient _graph;
    private readonly ILogger<GraphUserService> _logger;

    public GraphUserService(IOptions<AzureAdOptions> options, ILogger<GraphUserService> logger)
    {
        var o = options.Value;
        var credential = new ClientSecretCredential(o.TenantId, o.ClientId, o.ClientSecret);
        _graph = new GraphServiceClient(credential, Scopes);
        _logger = logger;
    }

    public async Task<GraphUserInfo?> FindUserAsync(string query, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(query))
            return null;

        // Escape single quotes for the OData filter (defence-in-depth; query is admin-only).
        var q = query.Trim().Replace("'", "''");

        try
        {
            var page = await _graph.Users.GetAsync(rc =>
            {
                rc.QueryParameters.Filter =
                    $"startswith(userPrincipalName,'{q}') or startswith(mail,'{q}') or startswith(displayName,'{q}')";
                rc.QueryParameters.Select = SelectFields;
                rc.QueryParameters.Top = 1;
            }, cancellationToken);

            var u = page?.Value?.FirstOrDefault();
            if (u is null)
                return null;

            return new GraphUserInfo(
                u.Id, u.UserPrincipalName, u.Mail, u.DisplayName,
                u.GivenName, u.Surname, u.JobTitle, u.Department);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Microsoft Graph user lookup failed for query '{Query}'.", query);
            return null;
        }
    }
}

/// <summary>No-op Graph lookup for environments without a client secret (local/dev).</summary>
public sealed class DisabledGraphUserService : IGraphUserService
{
    public Task<GraphUserInfo?> FindUserAsync(string query, CancellationToken cancellationToken = default)
        => Task.FromResult<GraphUserInfo?>(null);
}

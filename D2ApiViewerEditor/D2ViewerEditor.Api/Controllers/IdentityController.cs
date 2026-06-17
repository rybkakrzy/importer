using D2ViewerEditor.Api.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Controllers;

/// <summary>
/// Identity / authorization endpoints. <c>users</c> is an admin-only Microsoft Graph lookup;
/// <c>resources</c> returns the resource names the signed-in user may access (frontend authorization
/// gate). Backend is the source of truth — Angular guards are UX only.
/// </summary>
public class IdentityController : BaseApiController
{
    private readonly IGraphUserService _graphUsers;
    private readonly ResourcesProvider _resources;

    public IdentityController(IGraphUserService graphUsers, ResourcesProvider resources)
    {
        _graphUsers = graphUsers;
        _resources = resources;
    }

    /// <summary>Looks up a user in Entra ID by UPN / mail / display name (prefix match). Admin only.
    /// When Graph is not configured (no client secret) the lookup returns 404.</summary>
    /// <param name="query">Search term (start of UPN, mail or display name).</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    [HttpGet("users")]
    [Authorize(Policy = AuthorizationPolicies.RequireAppAdmin)]
    [ProducesResponseType(typeof(GraphUserInfo), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> FindUser([FromQuery] string query, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(query))
            return BadRequest(new { error = "Parametr 'query' jest wymagany." });

        var user = await _graphUsers.FindUserAsync(query, cancellationToken);
        return user is null ? NotFound() : Ok(user);
    }

    /// <summary>
    /// Resource names (frontend route names) the signed-in user may access, derived from their app
    /// roles. The Angular resource guard calls this and gates navigation by route path. Available to
    /// any application user (Operator or Administrator); a user with no app role gets an empty list.
    /// </summary>
    [HttpGet("resources")]
    [Authorize(Policy = AuthorizationPolicies.RequireAppOperator)]
    [ProducesResponseType(typeof(IEnumerable<string>), StatusCodes.Status200OK)]
    public IActionResult GetResources() => Ok(_resources.GetForUser(User));
}

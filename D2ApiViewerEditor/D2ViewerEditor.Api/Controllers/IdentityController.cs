using D2ViewerEditor.Api.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Controllers;

/// <summary>
/// Identity / directory endpoints backed by Microsoft Graph (Doc2 <c>IdentityController</c>).
/// Admin-only. When Graph is not configured (no client secret) the lookup returns 404.
/// </summary>
[Authorize(Policy = AuthorizationPolicies.RequireAppAdmin)]
public class IdentityController : BaseApiController
{
    private readonly IGraphUserService _graphUsers;

    public IdentityController(IGraphUserService graphUsers) => _graphUsers = graphUsers;

    /// <summary>Looks up a user in Entra ID by UPN / mail / display name (prefix match).</summary>
    /// <param name="query">Search term (start of UPN, mail or display name).</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    [HttpGet("users")]
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
}

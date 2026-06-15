using D2ViewerEditor.Domain.Common;
using MediatR;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Controllers;

/// <summary>
/// Bazowy kontroler API z obsługą wzorca Result.
/// Wymaga uwierzytelnienia (Entra ID) dla wszystkich akcji — cała aplikacja jest za logowaniem.
/// Endpointy administracyjne dokładają politykę <see cref="Security.AuthorizationPolicies.RequireAppAdmin"/>.
/// </summary>
[ApiController]
[Route("api/[controller]")]
[Authorize]
public abstract class BaseApiController : ControllerBase
{
    private IMediator? _mediator;

    protected IMediator Mediator =>
        _mediator ??= HttpContext.RequestServices.GetRequiredService<IMediator>();

    /// <summary>
    /// Mapuje Result&lt;T&gt; na odpowiedni HTTP response
    /// </summary>
    protected IActionResult HandleResult<T>(Result<T> result)
    {
        return result.Match<IActionResult>(
            onSuccess: value => Ok(value),
            onFailure: error => BadRequest(new ProblemDetails
            {
                Title = "Błąd walidacji",
                Detail = error,
                Status = StatusCodes.Status400BadRequest
            })
        );
    }

    /// <summary>
    /// Mapuje Result (void) na odpowiedni HTTP response
    /// </summary>
    protected IActionResult HandleResult(Result result)
    {
        return result.Match<IActionResult>(
            onSuccess: () => NoContent(),
            onFailure: error => BadRequest(new ProblemDetails
            {
                Title = "Błąd walidacji",
                Detail = error,
                Status = StatusCodes.Status400BadRequest
            })
        );
    }
}

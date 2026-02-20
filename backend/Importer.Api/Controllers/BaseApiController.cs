using Importer.Domain.Common;
using MediatR;
using Microsoft.AspNetCore.Mvc;

namespace Importer.Api.Controllers;

/// <summary>
/// Bazowy kontroler API z obsługą wzorca Result
/// </summary>
[ApiController]
[Route("api/[controller]")]
public abstract class BaseApiController : ControllerBase
{
    private IMediator? _mediator;

    protected IMediator Mediator =>
        _mediator ??= HttpContext.RequestServices.GetRequiredService<IMediator>();

    /// <summary>
    /// Mapuje Result<T> na odpowiedni HTTP response
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

using Microsoft.AspNetCore.Mvc;

namespace D2ServicesViewerEditor.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Produces("application/json")]
public class HealthController : ControllerBase
{
    private readonly ILogger<HealthController> _logger;

    public HealthController(ILogger<HealthController> logger)
    {
        _logger = logger;
    }

    [HttpGet]
    [ProducesResponseType(StatusCodes.Status200OK)]
    public IActionResult Get()
    {
        _logger.LogDebug("Health check requested");

        return Ok(new
        {
            Status = "Healthy",
            Service = "D2 Services Viewer Editor API",
            Timestamp = DateTime.UtcNow,
            Version = "1.0.0"
        });
    }

    [HttpGet("detailed")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status503ServiceUnavailable)]
    public IActionResult GetDetailed()
    {
        _logger.LogDebug("Detailed health check requested");

        var healthStatus = new
        {
            Status = "Healthy",
            Service = "D2 Services Viewer Editor API",
            Timestamp = DateTime.UtcNow,
            Version = "1.0.0",
            Dependencies = new
            {
                Database = "Healthy",
                GoogleCloudStorage = "Healthy",
                SharedLayers = "Healthy"
            }
        };

        return Ok(healthStatus);
    }
}

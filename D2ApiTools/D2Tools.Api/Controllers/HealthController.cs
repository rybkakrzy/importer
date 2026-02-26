using Microsoft.AspNetCore.Mvc;

namespace D2Tools.Api.Controllers;

/// <summary>
/// Health check — endpoint do sprawdzania dostępności API + informacje o buildzie
/// </summary>
[ApiController]
[Route("api/[controller]")]
[Tags("Health")]
public class HealthController : ControllerBase
{
    private readonly IConfiguration _configuration;

    public HealthController(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    /// <summary>
    /// Sprawdza dostępność API i zwraca informacje o buildzie
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(HealthResponse), StatusCodes.Status200OK)]
    public IActionResult GetHealth()
    {
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version?.ToString() ?? "1.0.0";
        var buildDate = System.IO.File.GetLastWriteTimeUtc(assembly.Location);
        var environmentName = _configuration["BuildInfo:Environment"] ?? "UNKNOWN";

        return Ok(new HealthResponse
        {
            Status = "healthy",
            Environment = environmentName,
            BuildNumber = version,
            BuildDate = buildDate.ToString("dd/MM/yyyy HH:mm"),
            Timestamp = DateTime.UtcNow
        });
    }
}

/// <summary>
/// Model odpowiedzi health check
/// </summary>
public class HealthResponse
{
    public string Status { get; set; } = string.Empty;
    public string Environment { get; set; } = string.Empty;
    public string BuildNumber { get; set; } = string.Empty;
    public string BuildDate { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; }
}

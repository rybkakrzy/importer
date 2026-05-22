using Microsoft.AspNetCore.Mvc;

namespace D2ServicesViewerEditor.Api.Controllers;

[ApiController]
[Route("api/v1/[controller]")]
[Produces("application/json")]
public class DocumentController : ControllerBase
{
    private readonly ILogger<DocumentController> _logger;

    public DocumentController(ILogger<DocumentController> logger)
    {
        _logger = logger;
    }

    [HttpGet("{documentId}")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    public async Task<IActionResult> GetDocument(Guid documentId)
    {
        _logger.LogInformation("External integration: GetDocument requested for {DocumentId}", documentId);

        return Ok(new
        {
            DocumentId = documentId,
            Title = "Sample Document",
            CreatedAt = DateTime.UtcNow.AddDays(-10),
            Status = "Draft",
            Message = "This is a placeholder. Implement actual logic using Application layer services."
        });
    }

    [HttpPost]
    [ProducesResponseType(StatusCodes.Status201Created)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    public async Task<IActionResult> CreateDocument([FromBody] CreateDocumentRequest request)
    {
        _logger.LogInformation("External integration: CreateDocument requested with title: {Title}", request.Title);

        var documentId = Guid.NewGuid();
        return CreatedAtAction(
            nameof(GetDocument),
            new { documentId },
            new
            {
                DocumentId = documentId,
                Title = request.Title,
                CreatedAt = DateTime.UtcNow,
                Message = "Placeholder: Implement actual creation logic"
            }
        );
    }

    [HttpPatch("{documentId}/status")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    public async Task<IActionResult> UpdateDocumentStatus(
        Guid documentId,
        [FromBody] UpdateStatusRequest request)
    {
        _logger.LogInformation(
            "External integration: UpdateDocumentStatus for {DocumentId} to {Status}",
            documentId,
            request.NewStatus);

        return Ok(new
        {
            DocumentId = documentId,
            NewStatus = request.NewStatus,
            UpdatedAt = DateTime.UtcNow,
            Message = "Placeholder: Implement actual update logic"
        });
    }
}

public record CreateDocumentRequest(string Title, string? Content);

public record UpdateStatusRequest(string NewStatus);

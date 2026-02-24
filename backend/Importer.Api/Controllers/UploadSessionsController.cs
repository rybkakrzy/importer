using Importer.Domain.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace Importer.Api.Controllers;

/// <summary>
/// Kontroler API do zarządzania sesjami importu
/// </summary>
[Route("api/[controller]")]
public class UploadSessionsController : BaseApiController
{
    private readonly IUploadSessionService _sessionService;

    public UploadSessionsController(IUploadSessionService sessionService)
    {
        _sessionService = sessionService;
    }

    [HttpGet]
    public IActionResult GetSessions()
    {
        var sessions = _sessionService.GetAllSessions();
        return Ok(sessions.Select(s => new
        {
            id = s.Id,
            fileName = s.FileName,
            status = s.Status,
            startedAt = s.StartedAt,
            completedAt = s.CompletedAt,
            filesCount = s.FilesCount,
            totalSize = s.TotalSize,
            errorMessage = s.ErrorMessage
        }));
    }
}

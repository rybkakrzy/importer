using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Controllers;

/// <summary>
/// Kontroler do zarządzania dokumentami i wersjami
/// </summary>
public class DocumentStorageController : BaseApiController
{
    /// <summary>
    /// Upload nowego dokumentu (pierwszy zapis)
    /// </summary>
    /// <param name="request">Dane dokumentu z contentem</param>
    /// <returns>GUID mastera i GUID wersji</returns>
    [HttpPost("upload")]
    [ProducesResponseType(typeof(UploadDocumentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UploadDocument([FromBody] UploadDocumentRequest request)
    {
        var command = new UploadDocumentCommand(
            Content: request.Content,
            FileName: request.Name,
            MimeType: request.MimeType,
            CreatedBy: request.CreatedBy ?? "System"
        );

        var result = await Mediator.Send(command);

        return result.IsSuccess
            ? Ok(result.Value)
            : BadRequest(result.Error);
    }

    /// <summary>
    /// Zapis nowej wersji dokumentu (każde zapisanie z GUI)
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <param name="request">Nowa zawartość dokumentu</param>
    /// <returns>GUID nowej wersji</returns>
    [HttpPost("{masterId:guid}/save")]
    [ProducesResponseType(typeof(SaveDocumentVersionResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> SaveDocumentVersion(Guid masterId, [FromBody] SaveDocumentVersionRequest request)
    {
        var command = new SaveDocumentVersionCommand(
            MasterId: masterId,
            Content: request.Content,
            CreatedBy: request.CreatedBy ?? "System"
        );

        var result = await Mediator.Send(command);

        return result.IsSuccess
            ? Ok(result.Value)
            : BadRequest(result.Error);
    }

    /// <summary>
    /// Pobranie aktywnej wersji dokumentu (dla GUI)
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <returns>Dokument z aktywną wersją</returns>
    [HttpGet("{masterId:guid}")]
    [ProducesResponseType(typeof(DocumentDto), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDocument(Guid masterId)
    {
        var query = new GetDocumentQuery(masterId);
        var result = await Mediator.Send(query);

        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(result.Error);
    }

    /// <summary>
    /// Pobranie listy wszystkich wersji dokumentu (dla historii)
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <returns>Lista wersji (metadane bez contentu)</returns>
    [HttpGet("{masterId:guid}/versions")]
    [ProducesResponseType(typeof(List<DocumentVersionDto>), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDocumentVersions(Guid masterId)
    {
        var query = new GetDocumentVersionsQuery(masterId);
        var result = await Mediator.Send(query);

        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(result.Error);
    }

    /// <summary>
    /// Przywrócenie wybranej wersji (cofnięcie się do poprzedniej wersji)
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <param name="versionId">GUID wersji do przywrócenia</param>
    /// <returns>Potwierdzenie operacji</returns>
    [HttpPost("{masterId:guid}/restore/{versionId:guid}")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> RestoreDocumentVersion(Guid masterId, Guid versionId)
    {
        var command = new RestoreDocumentVersionCommand(
            MasterId: masterId,
            VersionId: versionId
        );

        var result = await Mediator.Send(command);

        return result.IsSuccess
            ? Ok(new { Message = "Wersja została przywrócona", VersionId = versionId })
            : BadRequest(result.Error);
    }
}

// Request DTOs
public record UploadDocumentRequest(string Name, string MimeType, byte[] Content, string? CreatedBy);
public record SaveDocumentVersionRequest(byte[] Content, string? CreatedBy);

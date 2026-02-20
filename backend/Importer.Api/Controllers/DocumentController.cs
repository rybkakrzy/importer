using Importer.Application.Features.Documents.Commands.SaveDocument;
using Importer.Application.Features.Documents.Commands.UploadImage;
using Importer.Application.Features.Documents.Queries.GetNewDocument;
using Importer.Application.Features.Documents.Queries.GetTemplate;
using Importer.Application.Features.Documents.Queries.GetTemplates;
using Importer.Application.Features.Documents.Queries.OpenDocument;
using Importer.Domain.Models;
using Microsoft.AspNetCore.Mvc;

namespace Importer.Api.Controllers;

/// <summary>
/// Kontroler API dla edytora dokumentów Word
/// </summary>
[Route("api/[controller]")]
public class DocumentController : BaseApiController
{
    /// <summary>
    /// Wczytuje dokument DOCX i zwraca HTML
    /// </summary>
    [HttpPost("open")]
    [RequestSizeLimit(50 * 1024 * 1024)]
    public async Task<IActionResult> OpenDocument(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "Nie przesłano pliku" });

        var extension = Path.GetExtension(file.FileName).ToLower();
        if (extension != ".docx")
            return BadRequest(new { error = "Obsługiwane są tylko pliki DOCX" });

        var query = new OpenDocumentQuery(file.OpenReadStream(), file.FileName);
        var result = await Mediator.Send(query);
        return HandleResult(result);
    }

    /// <summary>
    /// Zapisuje HTML jako dokument DOCX
    /// </summary>
    [HttpPost("save")]
    public async Task<IActionResult> SaveDocument([FromBody] SaveDocumentRequest request)
    {
        var command = new SaveDocumentCommand(
            request.Html,
            request.OriginalFileName,
            request.Metadata,
            request.Header,
            request.Footer
        );

        var result = await Mediator.Send(command);

        return result.Match<IActionResult>(
            onSuccess: saveResult => File(
                saveResult.DocxBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                saveResult.FileName),
            onFailure: error => BadRequest(new { error })
        );
    }

    /// <summary>
    /// Tworzy nowy pusty dokument
    /// </summary>
    [HttpGet("new")]
    public async Task<IActionResult> NewDocument()
    {
        var result = await Mediator.Send(new GetNewDocumentQuery());
        return HandleResult(result);
    }

    /// <summary>
    /// Eksportuje dokument do PDF (konwersja przez DOCX)
    /// </summary>
    [HttpPost("export-pdf")]
    public IActionResult ExportToPdf([FromBody] SaveDocumentRequest request)
    {
        return StatusCode(501, new { error = "Eksport do PDF nie jest jeszcze zaimplementowany. Użyj funkcji Zapisz jako DOCX." });
    }

    /// <summary>
    /// Wgrywa obraz i zwraca Base64
    /// </summary>
    [HttpPost("upload-image")]
    [RequestSizeLimit(10 * 1024 * 1024)]
    public async Task<IActionResult> UploadImage(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "Nie przesłano pliku" });

        var command = new UploadImageCommand(
            file.OpenReadStream(),
            file.FileName,
            file.ContentType,
            file.Length
        );

        var result = await Mediator.Send(command);
        return HandleResult(result);
    }

    /// <summary>
    /// Pobiera listę szablonów dokumentów
    /// </summary>
    [HttpGet("templates")]
    public async Task<IActionResult> GetTemplates()
    {
        var result = await Mediator.Send(new GetTemplatesQuery());
        return HandleResult(result);
    }

    /// <summary>
    /// Pobiera szablon dokumentu
    /// </summary>
    [HttpGet("templates/{templateId}")]
    public async Task<IActionResult> GetTemplate(string templateId)
    {
        var result = await Mediator.Send(new GetTemplateQuery(templateId));
        return HandleResult(result);
    }
}

using D2Tools.Application.Features.Documents.Commands.SaveDocument;
using D2Tools.Application.Features.Documents.Commands.SignDocument;
using D2Tools.Application.Features.Documents.Commands.UploadImage;
using D2Tools.Application.Features.Documents.Queries.GetNewDocument;
using D2Tools.Application.Features.Documents.Queries.GetTemplate;
using D2Tools.Application.Features.Documents.Queries.GetTemplates;
using D2Tools.Application.Features.Documents.Queries.OpenDocument;
using D2Tools.Domain.Interfaces;
using D2Tools.Domain.Models;
using Microsoft.AspNetCore.Mvc;

namespace D2Tools.Api.Controllers;

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

    /// <summary>
    /// Podpisuje dokument certyfikatem cyfrowym
    /// </summary>
    [HttpPost("sign")]
    public async Task<IActionResult> SignDocument([FromBody] SignDocumentRequest request)
    {
        var command = new SignDocumentCommand(
            request.Html,
            request.OriginalFileName,
            request.Metadata,
            request.Header,
            request.Footer,
            request.CertificateBase64,
            request.CertificatePassword,
            request.SignerName,
            request.SignerTitle,
            request.SignerEmail,
            request.SignatureReason
        );

        var result = await Mediator.Send(command);

        return result.Match<IActionResult>(
            onSuccess: signResult => File(
                signResult.DocxBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                signResult.FileName),
            onFailure: error => BadRequest(new { error })
        );
    }

    /// <summary>
    /// Weryfikuje podpisy cyfrowe w dokumencie DOCX
    /// </summary>
    [HttpPost("verify-signatures")]
    [RequestSizeLimit(50 * 1024 * 1024)]
    public IActionResult VerifySignatures(IFormFile file, [FromServices] IDigitalSignatureService signatureService)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "Nie przesłano pliku" });

        try
        {
            using var stream = file.OpenReadStream();
            var signatures = signatureService.VerifySignatures(stream);
            return Ok(signatures);
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = $"Błąd weryfikacji podpisów: {ex.Message}" });
        }
    }
}

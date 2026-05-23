using System.Text.Json;
using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using MediatR;
using Microsoft.AspNetCore.Mvc;

namespace D2ServicesViewerEditor.Api.Controllers;

[ApiController]
[Route("api/v1/[controller]")]
[Produces("application/json")]
public class DocumentController : ControllerBase
{
    private readonly ILogger<DocumentController> _logger;
    private readonly IMediator _mediator;

    public DocumentController(ILogger<DocumentController> logger, IMediator mediator)
    {
        _logger = logger;
        _mediator = mediator;
    }

    /// <summary>
    /// Przyjmuje plik (DOCX lub PDF) od aplikacji zewnętrznej wraz z metadanymi.
    /// Dla DOCX zwraca { MasterId, VersionId } (VersionId = wersja edytowalna, na której pracuje edytor).
    /// Dla PDF zwraca tylko { MasterId } (PDF nie tworzy wersji edytowalnej).
    /// </summary>
    [HttpPost]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(110L * 1024 * 1024)]
    [RequestFormLimits(MultipartBodyLengthLimit = 110L * 1024 * 1024, ValueLengthLimit = int.MaxValue)]
    [ProducesResponseType(typeof(CreateDocumentResponse), StatusCodes.Status201Created)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    public async Task<IActionResult> CreateDocument(
        [FromForm] CreateDocumentRequest request,
        CancellationToken cancellationToken)
    {
        if (request.File is null || request.File.Length == 0)
            return BadRequest(new { error = "Plik jest wymagany" });

        if (!Enum.TryParse<DocumentClassification>(request.Classification, ignoreCase: true, out var classification))
            return BadRequest(new { error = "Klasyfikacja musi być jedną z: C1, C2, C3, C4" });

        var mimeType = ResolveMimeType(request.File);
        var isDocx = string.Equals(mimeType, IngestExternalDocumentCommandHandler.DocxMimeType, StringComparison.OrdinalIgnoreCase);
        var isPdf = string.Equals(mimeType, IngestExternalDocumentCommandHandler.PdfMimeType, StringComparison.OrdinalIgnoreCase);

        if (!isDocx && !isPdf)
            return BadRequest(new { error = "Wspierane są tylko pliki DOCX i PDF" });

        // ReturnUrl wymagany dla DOCX — bez niego edytor nie będzie wiedział gdzie odesłać plik po „Zakończ".
        // Dla PDF opcjonalny (dziś niewykorzystywany, zostawiony pod przyszłe scenariusze np. podpis cyfrowy).
        if (isDocx && string.IsNullOrWhiteSpace(request.ReturnUrl))
            return BadRequest(new { error = "ReturnUrl jest wymagany dla plików DOCX" });

        await using var stream = new MemoryStream((int)Math.Min(request.File.Length, int.MaxValue));
        await request.File.CopyToAsync(stream, cancellationToken);
        var content = stream.ToArray();

        var metadataJson = JsonSerializer.Serialize(new
        {
            returnUrl = string.IsNullOrWhiteSpace(request.ReturnUrl) ? null : request.ReturnUrl,
            classification = classification.ToString()
        });

        var createdBy = Request.Headers.TryGetValue("X-Created-By", out var headerValue)
            && !string.IsNullOrWhiteSpace(headerValue)
                ? headerValue.ToString()
                : "external-integration";

        var command = new IngestExternalDocumentCommand(
            Content: content,
            FileName: request.File.FileName,
            MimeType: mimeType,
            CreatedBy: createdBy,
            Metadata: metadataJson
        );

        var result = await _mediator.Send(command, cancellationToken);

        if (!result.IsSuccess)
        {
            _logger.LogWarning("Ingest dokumentu nie powiódł się: {Error}", result.Error);
            return BadRequest(new { error = result.Error });
        }

        var value = result.Value!;
        var response = new CreateDocumentResponse(value.MasterId, value.VersionId);

        _logger.LogInformation(
            "Zewnętrzny ingest dokumentu OK: MasterId={MasterId}, VersionId={VersionId}, Mime={Mime}, Classification={Classification}",
            response.MasterId, response.VersionId, mimeType, classification);

        return CreatedAtAction(nameof(GetDocument), new { documentId = response.MasterId }, response);
    }

    [HttpGet("{documentId}")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    public Task<IActionResult> GetDocument(Guid documentId)
    {
        _logger.LogInformation("External integration: GetDocument requested for {DocumentId}", documentId);

        return Task.FromResult<IActionResult>(Ok(new
        {
            DocumentId = documentId,
            Message = "Placeholder: implement read flow"
        }));
    }

    private static string ResolveMimeType(IFormFile file)
    {
        if (!string.IsNullOrWhiteSpace(file.ContentType))
            return file.ContentType;

        var ext = Path.GetExtension(file.FileName)?.ToLowerInvariant();
        return ext switch
        {
            ".docx" => IngestExternalDocumentCommandHandler.DocxMimeType,
            ".pdf" => IngestExternalDocumentCommandHandler.PdfMimeType,
            _ => "application/octet-stream"
        };
    }
}

public class CreateDocumentRequest
{
    /// <summary>Plik DOCX lub PDF</summary>
    public IFormFile? File { get; set; }

    /// <summary>
    /// URL, na który należy odesłać plik po zakończeniu edycji.
    /// Wymagany dla DOCX, opcjonalny dla PDF.
    /// </summary>
    public string? ReturnUrl { get; set; }

    /// <summary>Klasyfikacja dokumentu: C1, C2, C3 lub C4 (obligatoryjna)</summary>
    public string Classification { get; set; } = string.Empty;
}

/// <summary>
/// Odpowiedź uploadu. VersionId ustawione tylko dla DOCX (id wersji edytowalnej).
/// Dla PDF pole VersionId jest null.
/// </summary>
public record CreateDocumentResponse(Guid MasterId, Guid? VersionId);

public enum DocumentClassification
{
    C1 = 1,
    C2 = 2,
    C3 = 3,
    C4 = 4
}

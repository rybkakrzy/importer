using System.Text.Json;
using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UnlockDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;
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

        // Klasyfikacja jest OPCJONALNA. Brak/pusta → przyjmujemy bez klasyfikacji (null w metadanych).
        // Podana wartość musi być poprawna (C1..C4) — łapiemy literówki, ale nie wymuszamy obecności.
        DocumentClassification? classification = null;
        if (!string.IsNullOrWhiteSpace(request.Classification))
        {
            if (!Enum.TryParse<DocumentClassification>(request.Classification, ignoreCase: true, out var parsed))
                return BadRequest(new { error = "Jeśli podana, klasyfikacja musi być jedną z: C1, C2, C3, C4" });
            classification = parsed;
        }

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
            classification = classification?.ToString(),
            // Domain rule: missing field / non-true ⇒ false. Persisted only when explicitly true
            // so a stored false vs. missing is indistinguishable to the parser (both ⇒ blocked).
            userDownload = request.UserDownload == true ? (bool?)true : null
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

    // [HttpGet("{documentId}")]
    // [ProducesResponseType(StatusCodes.Status200OK)]
    // [ProducesResponseType(StatusCodes.Status404NotFound)]
    // [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    // public Task<IActionResult> GetDocument(Guid documentId)
    // {
    //     _logger.LogInformation("External integration: GetDocument requested for {DocumentId}", documentId);
    //     return Task.FromResult<IActionResult>(Ok(new
    //     {
    //         DocumentId = documentId,
    //         Message = "Placeholder: implement read flow"
    //     }));
    // }

    /// <summary>
    /// Aktualizuje URL, na który system odeśle plik po zakończeniu edycji ("Zakończ i wyślij").
    /// URL przechowywany w metadanych dokumentu master (obok klasyfikacji). Operacja jest
    /// idempotentna — ten sam URL w kolejnym wywołaniu nie zmienia stanu. Brak po stronie
    /// dokumentu w stanie wysyłki / wysłanym / nieudanej wysyłce (409 Conflict).
    /// </summary>
    [HttpPut("{masterId:guid}/callback-url")]
    [Consumes("application/json")]
    [ProducesResponseType(StatusCodes.Status204NoContent)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status409Conflict)]
    public async Task<IActionResult> UpdateCallbackUrl(
        [FromRoute] Guid masterId,
        [FromBody] UpdateCallbackUrlRequest request,
        CancellationToken cancellationToken)
    {
        var result = await _mediator.Send(
            new UpdateCallbackUrlCommand(masterId, request?.Url),
            cancellationToken);

        if (result.IsSuccess)
        {
            // The URL itself is intentionally NOT logged — it may carry an integration token.
            _logger.LogInformation("Callback URL updated for document {MasterId}.", masterId);
            return NoContent();
        }
        if (result.IsNotFound)
            return NotFound(new { error = result.Error });

        // Domain-level rejection (terminal/in-flight state) → 409 keeps it distinct from a
        // validation error (URL format mismatch → 400).
        var conflict = result.Error != null
            && result.Error.StartsWith("Nie można zaktualizować callback URL", StringComparison.OrdinalIgnoreCase);
        return conflict
            ? Conflict(new { error = result.Error })
            : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Odblokowuje dokument przetrzymywany w edytorze. W tej domenie status
    /// <c>Editing</c> jest sygnałem "trzymany przez edytora" (frontend pokazuje go jako
    /// <c>lockedByOther</c>), więc unlock = przejście <c>Editing → Saved</c>.
    /// Idempotentne dla <c>Saved</c> (200 OK, changed=false). Blokowane dla stanów
    /// wysyłkowych (409 Conflict).
    /// </summary>
    [HttpPost("{masterId:guid}/unlock")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(UnlockDocumentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status409Conflict)]
    public async Task<IActionResult> UnlockDocument(
        [FromRoute] Guid masterId,
        [FromBody] UnlockDocumentRequest? request,
        CancellationToken cancellationToken)
    {
        var result = await _mediator.Send(
            new UnlockDocumentCommand(masterId, request?.Reason),
            cancellationToken);

        if (result.IsSuccess)
            return Ok(result.Value);
        if (result.IsNotFound)
            return NotFound(new { error = result.Error });
        return Conflict(new { error = result.Error });
    }

    /// <summary>
    /// Zwraca aktualny stan dokumentu (status cyklu życia, aktywna wersja, najnowsze zadanie
    /// wysyłki, flaga <c>HasCallbackUrl</c>). Status jest na poziomie master — wersje nie
    /// mają własnego statusu, więc endpoint identyfikuje dokument jedynie przez
    /// <c>masterGuid</c>. Pełny <c>callbackUrl</c> NIE jest zwracany (może zawierać token).
    /// </summary>
    [HttpGet("{masterId:guid}/status")]
    [ProducesResponseType(typeof(DocumentStatusDto), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDocumentStatus(
        [FromRoute] Guid masterId,
        CancellationToken cancellationToken)
    {
        var result = await _mediator.Send(new GetDocumentStatusQuery(masterId), cancellationToken);
        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(new { error = result.Error });
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

    /// <summary>Klasyfikacja dokumentu: C1, C2, C3 lub C4. OPCJONALNA — brak/pusta = bez klasyfikacji.</summary>
    public string? Classification { get; set; }

    /// <summary>
    /// Opcjonalna flaga zezwalająca użytkownikowi na pobranie edytowanego pliku na komputer.
    /// Domyślnie <c>false</c> — brak pola / <c>null</c> / <c>false</c> blokuje pobieranie.
    /// Tylko <c>true</c> aktywuje menu „Pobierz dokument" w edytorze.
    /// <para>Niezależne od <c>ReturnUrl</c>: zwrotka do systemu źródłowego i pobranie przez
    /// użytkownika to dwa odrębne mechanizmy.</para>
    /// </summary>
    public bool? UserDownload { get; set; }
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

/// <summary>Request body for PUT /api/v1/document/{masterId}/callback-url.</summary>
public record UpdateCallbackUrlRequest(string? Url);

/// <summary>Request body for POST /api/v1/document/{masterId}/unlock (Reason optional, audited).</summary>
public record UnlockDocumentRequest(string? Reason);

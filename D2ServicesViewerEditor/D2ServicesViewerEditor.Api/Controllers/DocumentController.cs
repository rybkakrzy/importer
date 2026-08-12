using System.Text.Json;
using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UnlockDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;
using D2ViewerEditor.Domain.Entities;
using MediatR;
using Microsoft.AspNetCore.Mvc;

namespace D2ServicesViewerEditor.Api.Controllers;

[ApiController]
[Route("api/v1/[controller]")]
[Produces("application/json")]
public class DocumentController : ControllerBase
{
    // Allowlist rozszerzeń przyjmowanych na granicy zaufania żyje jako JAWNY switch na
    // literałach w CreateDocument (wymóg czytelności dla SAST). Rozszerzenie jest źródłem
    // prawdy, deklarowany Content-Type musi się z nim zgadzać (klient może go dowolnie podać).
    private static readonly char[] PathSeparators = { '/', '\\', ':' };

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

        DocumentClassification? classification = null;
        if (!string.IsNullOrWhiteSpace(request.Classification))
        {
            if (!Enum.TryParse<DocumentClassification>(request.Classification, ignoreCase: true, out var parsed))
                return BadRequest(new { error = "Jeśli podana, klasyfikacja musi być jedną z: C1, C2, C3, C4" });
            classification = parsed;
        }

        // Nazwa pliku pochodzi od klienta — bierzemy wyłącznie segment nazwy (bez ścieżki),
        // a rozszerzenie musi należeć do allowlisty ZANIM cokolwiek wczytamy z formularza.
        // Walidacja JAWNYM switchem na literałach (nie słownikiem): to wzorzec, który silniki
        // SAST rozpoznają jako sanityzację rozszerzenia (finding „Dangerous File Extension").
        var sanitizedBaseName = Path.GetFileNameWithoutExtension(SanitizeFileName(request.File.FileName));

        string canonicalExtension;
        string mimeType;
        switch (Path.GetExtension(SanitizeFileName(request.File.FileName)).ToLowerInvariant())
        {
            case ".docx":
                canonicalExtension = ".docx";
                mimeType = IngestExternalDocumentCommandHandler.DocxMimeType;
                break;
            case ".pdf":
                canonicalExtension = ".pdf";
                mimeType = IngestExternalDocumentCommandHandler.PdfMimeType;
                break;
            default:
                return BadRequest(new { error = "Wspierane są tylko pliki DOCX i PDF" });
        }

        // Nazwa przekazywana dalej jest SKŁADANA NA SERWERZE: baza po sanityzacji + kanoniczne
        // rozszerzenie z literału powyżej — rozszerzenie klienta nigdy nie płynie do zapisu.
        // (Sam zapis i tak idzie pod GUID-em wersji; nazwa jest wyłącznie metadaną wyświetlania.)
        var fileName = sanitizedBaseName + canonicalExtension;

        if (!IsDeclaredContentTypeAcceptable(request.File.ContentType, mimeType))
            return BadRequest(new { error = "Deklarowany typ MIME nie zgadza się z rozszerzeniem pliku" });

        var isDocx = string.Equals(mimeType, IngestExternalDocumentCommandHandler.DocxMimeType, StringComparison.OrdinalIgnoreCase);

        if (isDocx && string.IsNullOrWhiteSpace(request.ReturnUrl))
            return BadRequest(new { error = "ReturnUrl jest wymagany dla plików DOCX" });

        await using var stream = new MemoryStream((int)Math.Min(request.File.Length, int.MaxValue));
        await request.File.CopyToAsync(stream, cancellationToken);
        var content = stream.ToArray();

        var metadataJson = JsonSerializer.Serialize(new
        {
            returnUrl = string.IsNullOrWhiteSpace(request.ReturnUrl) ? null : request.ReturnUrl,
            classification = classification?.ToString(),
            userDownload = request.UserDownload == true ? (bool?)true : null,
            // Inverse-default flag: zapisujemy TYLKO jawne false (ukrycie UI zapisu w edytorze).
            // Brak / true → null → edytor pokazuje autosave i przycisk „Zapisz" jak dotąd.
            showSaveState = request.ShowSaveState == false ? (bool?)false : null
        });

        var createdBy = Request.Headers.TryGetValue("X-Created-By", out var headerValue)
            && !string.IsNullOrWhiteSpace(headerValue)
                ? headerValue.ToString()
                : "external-integration";

        var command = new IngestExternalDocumentCommand(
            Content: content,
            FileName: fileName,
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

        return StatusCode(StatusCodes.Status201Created, response);
    }

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
            _logger.LogInformation("Callback URL updated for document {MasterId}.", masterId);
            return NoContent();
        }
        if (result.IsNotFound)
            return NotFound(new { error = result.Error });

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
    /// Zwraca wyłącznie status cyklu życia dokumentu (na poziomie master). W tej, zewnętrznej
    /// integracji odpowiedź jest celowo zawężona do <c>masterId</c> + <c>status</c> — bez
    /// danych o aktywnej wersji, dostawie czy flagach (<c>HasCallbackUrl</c>/<c>UserDownload</c>).
    /// Pod spodem nadal używane jest współdzielone <see cref="GetDocumentStatusQuery"/>, ale
    /// jego pełny <see cref="DocumentStatusDto"/> jest tu rzutowany na okrojoną odpowiedź,
    /// więc bogatszy kontrakt D2ApiViewerEditor pozostaje bez zmian.
    /// </summary>
    [HttpGet("{masterId:guid}/status")]
    [ProducesResponseType(typeof(DocumentStatusResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status401Unauthorized)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDocumentStatus(
        [FromRoute] Guid masterId,
        CancellationToken cancellationToken)
    {
        var result = await _mediator.Send(new GetDocumentStatusQuery(masterId), cancellationToken);
        if (!result.IsSuccess)
            return NotFound(new { error = result.Error });

        var status = result.Value!;
        // DocumentStatusDto.Status to nazwa enuma (document.Status.ToString()) — rzutujemy ją
        // z powrotem na DocumentStatus, by w Swaggerze odpowiedź pokazywała listę opcji enuma.
        var statusEnum = Enum.Parse<DocumentStatus>(status.Status);
        return Ok(new DocumentStatusResponse(status.MasterId, statusEnum));
    }

    // Nazwa pliku z multipartu bywa pełną ścieżką (klient/legacy przeglądarki). Odcinamy wszystko
    // do ostatniego separatora — niezależnie od platformy hosta, bo Path.GetFileName na Linuksie
    // nie traktuje '\' jako separatora.
    private static string SanitizeFileName(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return string.Empty;

        var lastSeparator = fileName.LastIndexOfAny(PathSeparators);
        return (lastSeparator < 0 ? fileName : fileName[(lastSeparator + 1)..]).Trim();
    }

    // Content-Type jest deklaracją klienta: akceptujemy go tylko wtedy, gdy zgadza się z typem
    // wynikającym z rozszerzenia. Brak deklaracji nie blokuje uploadu (rozszerzenie już przeszło
    // allowlistę), a spójność z magic-bytes weryfikuje IFileUploadSecurityService w handlerze.
    private static bool IsDeclaredContentTypeAcceptable(string? declaredContentType, string expectedMimeType) =>
        string.IsNullOrWhiteSpace(declaredContentType)
        || string.Equals(declaredContentType, expectedMimeType, StringComparison.OrdinalIgnoreCase);
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

    /// <summary>
    /// Czy w edytorze ma być widoczny stan zapisu (sekcja autozapisu + osobny przycisk „Zapisz").
    /// Domyślnie <c>true</c>. Tylko jawne <c>false</c> ukrywa oba elementy. Brak pola / <c>null</c>
    /// / <c>true</c> → zachowanie jak dotychczas (oba widoczne).
    /// </summary>
    public bool? ShowSaveState { get; set; }
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

/// <summary>
/// Okrojona odpowiedź GET /api/v1/document/{masterId}/status dla integracji zewnętrznej —
/// świadomie zawiera wyłącznie identyfikator master i status cyklu życia. Pełny kontrakt
/// (wersja, dostawa, flagi) udostępnia <c>DocumentStatusDto</c> w D2ApiViewerEditor.
/// </summary>
/// <param name="MasterId">Identyfikator dokumentu master.</param>
/// <param name="Status">Status cyklu życia dokumentu (serializowany jako nazwa enuma).</param>
public record DocumentStatusResponse(Guid MasterId, DocumentStatus Status);

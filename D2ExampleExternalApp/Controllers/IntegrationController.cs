using System.Net.Http.Headers;
using System.Security.Cryptography;
using D2ExampleExternalApp.Configuration;
using D2ExampleExternalApp.Integration;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace D2ExampleExternalApp.Controllers;

/// <summary>
/// Symulacja aplikacji zewnętrznej integrującej się z D2 ViewerEditor.
/// Wysyłka dokumentu do D2 + odbiór gotowego pliku z powrotem (returnUrl).
/// </summary>
[ApiController]
[Route("api/integration")]
[Produces("application/json")]
public class IntegrationController : ControllerBase
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<IntegrationController> _logger;
    private readonly D2ServicesOptions _d2;
    private readonly ExternalAppOptions _app;

    public IntegrationController(
        IHttpClientFactory httpClientFactory,
        ILogger<IntegrationController> logger,
        IOptions<D2ServicesOptions> d2,
        IOptions<ExternalAppOptions> app)
    {
        _httpClientFactory = httpClientFactory;
        _logger = logger;
        _d2 = d2.Value;
        _app = app.Value;
    }

    /// <summary>
    /// Krok 1: wgrywasz plik (DOCX) — aplikacja przesyła go do D2ServicesViewerEditor
    /// (POST /api/v1/document) z returnUrl wskazującym z powrotem na własny endpoint callback.
    /// Zwraca surową odpowiedź D2 (201 + { masterId, versionId }).
    /// </summary>
    [HttpPost("upload")]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(StatusCodes.Status201Created)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UploadAndSend([FromForm] UploadRequest request, CancellationToken ct)
    {
        if (request.File is null || request.File.Length == 0)
            return BadRequest(new { error = "Plik jest wymagany" });

        var returnUrl = $"{_app.PublicBaseUrl.TrimEnd('/')}/api/integration/callback";

        // Budujemy multipart/form-data dokładnie wg kontraktu D2Services CreateDocument:
        // File (wymagane), ReturnUrl (wymagane dla DOCX), Classification/UserDownload (opcjonalne).
        using var form = new MultipartFormDataContent();
        await using var fileStream = request.File.OpenReadStream();
        var fileContent = new StreamContent(fileStream);
        fileContent.Headers.ContentType = new MediaTypeHeaderValue(
            string.IsNullOrWhiteSpace(request.File.ContentType)
                ? "application/octet-stream"
                : request.File.ContentType);
        form.Add(fileContent, "File", request.File.FileName);
        form.Add(new StringContent(returnUrl), "ReturnUrl");
        if (!string.IsNullOrWhiteSpace(request.Classification))
            form.Add(new StringContent(request.Classification), "Classification");
        if (request.UserDownload.HasValue)
            form.Add(new StringContent(request.UserDownload.Value ? "true" : "false"), "UserDownload");

        var client = _httpClientFactory.CreateClient(D2ServicesOptions.HttpClientName);
        using var req = new HttpRequestMessage(HttpMethod.Post, "/api/v1/document") { Content = form };
        req.Headers.TryAddWithoutValidation("X-Created-By", "D2ExampleExternalApp");
        if (!string.IsNullOrWhiteSpace(_d2.ApiKey))
            req.Headers.TryAddWithoutValidation("X-Api-Key", _d2.ApiKey);

        HttpResponseMessage resp;
        try
        {
            resp = await client.SendAsync(req, ct);
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Nie udało się połączyć z D2Services pod {BaseUrl}", _d2.BaseUrl);
            return StatusCode(StatusCodes.Status502BadGateway,
                new { error = $"Nie udało się połączyć z D2Services ({_d2.BaseUrl}): {ex.Message}" });
        }

        var body = await resp.Content.ReadAsStringAsync(ct);
        _logger.LogInformation("D2Services ingest -> HTTP {Status}, returnUrl={ReturnUrl}, body={Body}",
            (int)resp.StatusCode, returnUrl, body);

        // Przekazujemy odpowiedź D2 1:1 (status + body), żeby było widać masterId/versionId.
        return new ContentResult
        {
            StatusCode = (int)resp.StatusCode,
            Content = body,
            ContentType = "application/json"
        };
    }

    /// <summary>
    /// Krok 3 (returnUrl): D2 odsyła tu gotowy plik po „Zakończ i wyślij".
    /// Kontrakt = multipart/form-data: file, masterId, versionId, corporateKey
    /// + nagłówki Idempotency-Key i X-Content-SHA256. Zapisujemy plik, weryfikujemy hash
    /// i zwracamy 200 (2xx jest WYMAGANE — inaczej D2 ponawia wysyłkę).
    /// </summary>
    /// <remarks>
    /// UWAGA dot. nazw pól: D2 wysyła pole z użytkownikiem jako <c>corporateKey</c>
    /// (a nie <c>userCk</c>). Jeśli Twój endpoint nazywa je inaczej — zmień nazwę na <c>corporateKey</c>.
    /// </remarks>
    [HttpPost("callback")]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> ReceiveCallback([FromForm] CallbackRequest request, CancellationToken ct)
    {
        var file = request.File;
        var masterId = request.MasterId;
        var versionId = request.VersionId;
        var corporateKey = request.CorporateKey;

        if (file is null || file.Length == 0)
            return BadRequest(new { error = "Brak części 'file' w multipart/form-data" });

        var idempotencyKey = Request.Headers["Idempotency-Key"].ToString();
        var expectedSha = Request.Headers["X-Content-SHA256"].ToString();

        Directory.CreateDirectory(_app.ReceivedFilesPath);
        var safeName = $"{Sanitize(masterId)}_{Sanitize(versionId)}_{Sanitize(file.FileName)}";
        var savedPath = Path.GetFullPath(Path.Combine(_app.ReceivedFilesPath, safeName));

        byte[] bytes;
        await using (var ms = new MemoryStream())
        {
            await file.CopyToAsync(ms, ct);
            bytes = ms.ToArray();
        }
        await System.IO.File.WriteAllBytesAsync(savedPath, bytes, ct);

        var computedSha = Convert.ToHexString(SHA256.HashData(bytes));
        var hashMatches = string.IsNullOrEmpty(expectedSha)
            || string.Equals(computedSha, expectedSha, StringComparison.OrdinalIgnoreCase);

        var record = new ReceivedDocument(
            MasterId: masterId,
            VersionId: versionId,
            CorporateKey: corporateKey,
            FileName: file.FileName,
            SizeBytes: bytes.Length,
            IdempotencyKey: string.IsNullOrEmpty(idempotencyKey) ? null : idempotencyKey,
            ExpectedSha256: string.IsNullOrEmpty(expectedSha) ? null : expectedSha,
            ComputedSha256: computedSha,
            HashMatches: hashMatches,
            ReceivedAtUtc: DateTime.UtcNow,
            SavedPath: savedPath);

        var isNew = ReceivedDocumentStore.Add(record);

        _logger.LogInformation(
            "Callback odebrany: master={Master} version={Version} corporateKey={Ck} size={Size}B hashOk={HashOk} idem={Idem} duplicate={Dup}",
            masterId, versionId, corporateKey, bytes.Length, hashMatches, idempotencyKey, !isNew);

        if (!hashMatches)
        {
            // 200 i tak zwracamy (przyjęliśmy plik), ale sygnalizujemy rozjazd hasha w body.
            _logger.LogWarning("Hash mismatch: expected={Expected} computed={Computed}", expectedSha, computedSha);
        }

        return Ok(new
        {
            received = true,
            duplicate = !isNew,
            masterId,
            versionId,
            corporateKey,
            fileName = file.FileName,
            sizeBytes = bytes.Length,
            idempotencyKey,
            hashMatches,
            savedPath
        });
    }

    /// <summary>Podgląd dokumentów odebranych w callbacku (najnowsze pierwsze).</summary>
    [HttpGet("received")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    public IActionResult ListReceived() => Ok(ReceivedDocumentStore.Snapshot());

    /// <summary>
    /// Pomocniczo: odpytuje status dokumentu w D2 (GET /api/v1/document/{masterId}/status).
    /// Przydatne, by zobaczyć przejścia Saved → Editing → Sending → Sent.
    /// </summary>
    [HttpGet("status/{masterId:guid}")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    public async Task<IActionResult> GetStatus([FromRoute] Guid masterId, CancellationToken ct)
    {
        var client = _httpClientFactory.CreateClient(D2ServicesOptions.HttpClientName);
        using var req = new HttpRequestMessage(HttpMethod.Get, $"/api/v1/document/{masterId}/status");
        if (!string.IsNullOrWhiteSpace(_d2.ApiKey))
            req.Headers.TryAddWithoutValidation("X-Api-Key", _d2.ApiKey);

        using var resp = await client.SendAsync(req, ct);
        var body = await resp.Content.ReadAsStringAsync(ct);
        return new ContentResult
        {
            StatusCode = (int)resp.StatusCode,
            Content = body,
            ContentType = "application/json"
        };
    }

    private static string Sanitize(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "unknown";
        var invalid = Path.GetInvalidFileNameChars();
        return new string(value.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
    }
}

/// <summary>
/// Formularz uploadu. ReturnUrl NIE jest tu podawany — aplikacja zewnętrzna sama wie,
/// pod jaki swój endpoint D2 ma oddzwonić (budujemy go z ExternalApp:PublicBaseUrl).
/// </summary>
public class UploadRequest
{
    /// <summary>Plik DOCX do wysłania do edytora.</summary>
    public IFormFile? File { get; set; }

    /// <summary>Opcjonalna klasyfikacja: C1, C2, C3 lub C4.</summary>
    public string? Classification { get; set; }

    /// <summary>Opcjonalnie: czy użytkownik może pobrać edytowany plik na dysk.</summary>
    public bool? UserDownload { get; set; }
}

/// <summary>
/// Części zwrotki przyjmowane na returnUrl. Nazwy właściwości odpowiadają (case-insensitive)
/// częściom multipart wysyłanym przez D2: <c>file</c>, <c>masterId</c>, <c>versionId</c>,
/// <c>corporateKey</c>.
/// </summary>
public class CallbackRequest
{
    /// <summary>Gotowy plik DOCX (część multipart o nazwie <c>file</c>).</summary>
    public IFormFile? File { get; set; }

    /// <summary>GUID dokumentu master.</summary>
    public string? MasterId { get; set; }

    /// <summary>GUID wersji edytowalnej.</summary>
    public string? VersionId { get; set; }

    /// <summary>Corporate key użytkownika kończącego edycję (D2 wysyła to jako <c>corporateKey</c>).</summary>
    public string? CorporateKey { get; set; }
}

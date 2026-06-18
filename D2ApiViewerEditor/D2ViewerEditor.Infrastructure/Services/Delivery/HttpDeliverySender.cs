using System.Net;
using System.Net.Http.Headers;
using D2ViewerEditor.Domain.Interfaces;
using Microsoft.Extensions.Logging;

namespace D2ViewerEditor.Infrastructure.Services.Delivery;

/// <summary>
/// Wysyłka dokumentu na returnUrl przez HTTP POST jako multipart/form-data (pole pliku "file").
/// Idempotency-Key = deliveryId pozwala odbiorcy deduplikować przy at-least-once.
/// Klasyfikuje wynik na sukces / retryable / permanent.
/// </summary>
public class HttpDeliverySender : IDeliverySender
{
    private static readonly HashSet<int> PermanentStatusCodes = new() { 400, 401, 403, 404, 405, 422 };

    // Finish-and-send dostarcza zawsze edytowalny DOCX (PDF jest tylko do podglądu, nie ma wersji
    // edytowalnej → nie da się go „Zakończyć"), więc nazwa pliku i typ są stałe i poprawne dla tego flow.
    private const string FileFieldName = "file";
    private const string FileName = "document.docx";
    private const string DocxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private const string MasterIdFieldName = "masterId";
    private const string VersionIdFieldName = "versionId";
    private const string CorporateKeyFieldName = "corporateKey";

    private readonly HttpClient _http;
    private readonly ILogger<HttpDeliverySender> _logger;

    public HttpDeliverySender(HttpClient http, ILogger<HttpDeliverySender> logger)
    {
        _http = http;
        _logger = logger;
    }

    public async Task<DeliveryResult> SendAsync(DeliveryDispatch dispatch, CancellationToken cancellationToken = default)
    {
        // multipart/form-data: plik w polu "file" (z nazwą i typem). MultipartFormDataContent przejmuje
        // własność części i zwolni ByteArrayContent przy dispose — nie dispose'ujemy go osobno.
        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(dispatch.Content);
        fileContent.Headers.ContentType = new MediaTypeHeaderValue(DocxContentType);
        content.Add(fileContent, FileFieldName, FileName);

        // Identifying fields alongside the file (backward compatible — the "file" part is unchanged).
        // corporateKey may be absent → empty value so the field set stays predictable for the recipient.
        content.Add(new StringContent(dispatch.MasterId.ToString()), MasterIdFieldName);
        content.Add(new StringContent(dispatch.VersionId.ToString()), VersionIdFieldName);
        content.Add(new StringContent(dispatch.CorporateKey ?? string.Empty), CorporateKeyFieldName);

        using var request = new HttpRequestMessage(HttpMethod.Post, dispatch.RecipientUrl) { Content = content };
        request.Headers.TryAddWithoutValidation("Idempotency-Key", dispatch.DeliveryId.ToString());
        request.Headers.TryAddWithoutValidation("X-Content-SHA256", dispatch.Sha256);

        // Technical log: IDs + size only, never the file content. corporateKey is a business-context
        // identifier (not a secret); logged as presence flag to stay conservative.
        _logger.LogInformation(
            "Sending delivery {DeliveryId} (master={MasterId} version={VersionId}, {SizeBytes} B, corporateKey={HasCorporateKey})",
            dispatch.DeliveryId, dispatch.MasterId, dispatch.VersionId, dispatch.Content.Length,
            string.IsNullOrEmpty(dispatch.CorporateKey) ? "absent" : "present");

        try
        {
            using var response = await _http.SendAsync(request, cancellationToken);
            if (response.IsSuccessStatusCode)
                return DeliveryResult.Succeeded();

            var code = (int)response.StatusCode;
            var reason = $"HTTP {code} {response.ReasonPhrase}";

            return PermanentStatusCodes.Contains(code)
                ? DeliveryResult.Permanent(reason)
                : DeliveryResult.Retryable(reason);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw; // shutdown — nie traktuj jako błąd wysyłki, zadanie wróci po wygaśnięciu lease
        }
        catch (TaskCanceledException)
        {
            return DeliveryResult.Retryable("Timeout wysyłki");
        }
        catch (HttpRequestException ex)
        {
            return DeliveryResult.Retryable(ex.Message);
        }
    }
}

using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Application.Features.Documents.Commands.DownloadEditedDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.AbortSend;
using D2ViewerEditor.Application.Features.Documents.Commands.ContinueDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.CancelDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.RequeueDelivery;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;
using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveriesByStatus;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDeliveryStatus;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDocumentVersion;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentBaseContent;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersionContent;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocuments;
using D2ViewerEditor.Api.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace D2ViewerEditor.Api.Controllers;

/// <summary>
/// Kontroler do zarządzania dokumentami i wersjami
/// </summary>
public class DocumentStorageController : BaseApiController
{
    /// <summary>
    /// Pobranie listy wszystkich dokumentów (dla administracji)
    /// </summary>
    /// <returns>Lista dokumentów bez contentu</returns>
    [HttpGet]
    [Authorize(Policy = AuthorizationPolicies.RequireAppAdmin)]
    [ProducesResponseType(typeof(List<DocumentListItemDto>), StatusCodes.Status200OK)]
    public async Task<IActionResult> GetDocuments([FromQuery] int skip = 0, [FromQuery] int take = 200)
    {
        var query = new GetDocumentsQuery(skip, take);
        var result = await Mediator.Send(query);

        return result.IsSuccess
            ? Ok(result.Value)
            : BadRequest(result.Error);
    }

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
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(result.Error);
    }

    /// <summary>
    /// Nadpisanie istniejącej wersji w miejscu (auto-save edytora).
    /// Podmienia plik w GCS pod tym samym versionId — nie tworzy nowych wersji.
    /// Wersja oryginalna (v1) jest nietykalna (zwraca 400).
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <param name="versionId">GUID wersji do nadpisania (edytowalna)</param>
    /// <param name="request">Nowa zawartość dokumentu</param>
    [HttpPut("{masterId:guid}/versions/{versionId:guid}")]
    [ProducesResponseType(typeof(UpdateDocumentVersionResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UpdateDocumentVersion(
        Guid masterId, Guid versionId, [FromBody] SaveDocumentVersionRequest request)
    {
        var command = new UpdateDocumentVersionCommand(
            MasterId: masterId,
            VersionId: versionId,
            Content: request.Content
        );

        var result = await Mediator.Send(command);

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Pobranie metadanych dokumentu przysłanych przez aplikację zewnętrzną (returnUrl, classification).
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    [HttpGet("{masterId:guid}/metadata")]
    [ProducesResponseType(typeof(DocumentMetadataDto), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDocumentMetadata(Guid masterId)
    {
        var query = new GetDocumentMetadataQuery(masterId);
        var result = await Mediator.Send(query);

        if (result.IsForbidden)
            return StatusCode(StatusCodes.Status403Forbidden, new { error = result.Error });

        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(new { error = result.Error });
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

        if (result.IsForbidden)
            return StatusCode(StatusCodes.Status403Forbidden, new { error = result.Error });

        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(new { error = result.Error });
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
            : NotFound(new { error = result.Error });
    }

    /// <summary>
    /// Pobranie bazowego (oryginalnego) pliku dokumentu — pierwsza wersja po uploadzie
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <returns>Oryginalny plik z właściwym Content-Type</returns>
    [HttpGet("{masterId:guid}/download")]
    [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> DownloadBaseDocument(Guid masterId)
    {
        var query = new GetDocumentBaseContentQuery(masterId);
        var result = await Mediator.Send(query);

        if (result.IsForbidden)
            return StatusCode(StatusCodes.Status403Forbidden, new { error = result.Error });
        if (!result.IsSuccess)
            return NotFound(new { error = result.Error });

        var dto = result.Value!;
        return File(dto.Content, dto.MimeType, dto.FileName);
    }

    /// <summary>
    /// Pobranie fizycznego pliku konkretnej wersji dokumentu
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <param name="versionId">GUID wersji do pobrania</param>
    /// <returns>Fizyczny plik wersji dokumentu</returns>
    [HttpGet("{masterId:guid}/versions/{versionId:guid}/download")]
    [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> DownloadDocumentVersion(Guid masterId, Guid versionId)
    {
        var query = new GetDocumentVersionContentQuery(masterId, versionId);
        var result = await Mediator.Send(query);

        if (result.IsForbidden)
            return StatusCode(StatusCodes.Status403Forbidden, new { error = result.Error });
        if (!result.IsSuccess)
            return NotFound(new { error = result.Error });

        var dto = result.Value!;
        return File(dto.Content, dto.MimeType, dto.FileName);
    }

    /// <summary>
    /// Pobranie edytowanego pliku na komputer użytkownika ("Pobierz dokument"). Działa tylko,
    /// gdy metadane dokumentu mają <c>userDownload == true</c> (lokalny upload ustawia tę flagę
    /// automatycznie; dokumenty z aplikacji zewnętrznej muszą mieć ją jawnie). Pełne egzekwowanie
    /// po stronie backendu — frontend dodatkowo ukrywa pozycję menu.
    /// </summary>
    [HttpPost("{masterId:guid}/user-download")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status403Forbidden)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> DownloadEditedDocument(
        [FromRoute] Guid masterId,
        [FromBody] Domain.Models.SaveDocumentRequest request,
        CancellationToken cancellationToken)
    {
        var command = new DownloadEditedDocumentCommand(
            MasterId: masterId,
            Html: request?.Html ?? string.Empty,
            OriginalFileName: request?.OriginalFileName,
            Metadata: request?.Metadata,
            Header: request?.Header,
            Footer: request?.Footer,
            Margins: request?.Margins,
            PageSize: request?.PageSize);

        var result = await Mediator.Send(command, cancellationToken);

        if (result.IsSuccess)
            return File(result.Value!.DocxBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                result.Value.FileName);

        if (result.IsNotFound)
            return NotFound(new { error = result.Error });

        if (result.Error != null
            && result.Error.StartsWith(DownloadEditedDocumentCommandHandler.ForbiddenErrorPrefix, StringComparison.Ordinal))
        {
            return StatusCode(StatusCodes.Status403Forbidden, new
            {
                error = result.Error[(DownloadEditedDocumentCommandHandler.ForbiddenErrorPrefix.Length)..].Trim()
            });
        }

        return BadRequest(new { error = result.Error });
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
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(result.Error);
    }

    /// <summary>
    /// "Zakończ i wyślij": utrwala stan edytora, zamraża snapshot finalnego pliku, ustawia status
    /// "Zlecono do wysyłki" i wykonuje SYNCHRONICZNĄ pierwszą próbę dostarczenia na returnUrl.
    /// Sukces → "Wysłano"; błąd → "Błąd wysyłki" + zadanie czeka na decyzję użytkownika
    /// ("Przerwij" / "Kontynuuj wysyłkę w tle"). Wielokrotne kliknięcie jest idempotentne.
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    /// <param name="versionId">GUID wersji edytowalnej do sfinalizowania</param>
    /// <param name="request">Aktualna zawartość edytora</param>
    [HttpPost("{masterId:guid}/versions/{versionId:guid}/finish")]
    [ProducesResponseType(typeof(FinishAndSendResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> FinishAndSend(
        Guid masterId, Guid versionId, [FromBody] SaveDocumentVersionRequest request)
    {
        var command = new FinishAndSendDocumentCommand(
            MasterId: masterId,
            VersionId: versionId,
            Content: request.Content,
            CreatedBy: request.CreatedBy);

        var result = await Mediator.Send(command);

        if (result.IsNotFound)
            return NotFound(new { error = result.Error });
        if (result.IsFailure)
            return BadRequest(new { error = result.Error });

        return Ok(result.Value);
    }

    /// <summary>
    /// "Przerwij" po nieudanej pierwszej próbie wysyłki: anuluje zadanie (Cancelled) i ustawia
    /// dokument na "UzytkownikPrzerwałWysyłkę". Dokument zostaje edytowalny, nic nie idzie w tle.
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    [HttpPost("{masterId:guid}/abort-send")]
    [ProducesResponseType(typeof(AbortSendResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> AbortSend(Guid masterId)
    {
        var result = await Mediator.Send(new AbortSendCommand(masterId));

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// "Kontynuuj wysyłkę w tle" po nieudanej pierwszej próbie: przywraca zadanie do kolejki
    /// i ustawia dokument na "Zlecono do wysyłki". Dalej dostarcza je worker w tle.
    /// </summary>
    /// <param name="masterId">GUID mastera dokumentu</param>
    [HttpPost("{masterId:guid}/continue-delivery")]
    [ProducesResponseType(typeof(ContinueDeliveryResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> ContinueDelivery(Guid masterId)
    {
        var result = await Mediator.Send(new ContinueDeliveryCommand(masterId));

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Status zadania wysyłki (polling z GUI).
    /// </summary>
    /// <param name="deliveryId">GUID zadania wysyłki</param>
    [HttpGet("deliveries/{deliveryId:guid}")]
    [ProducesResponseType(typeof(DeliveryStatusDto), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    public async Task<IActionResult> GetDeliveryStatus(Guid deliveryId)
    {
        var result = await Mediator.Send(new GetDeliveryStatusQuery(deliveryId));

        return result.IsSuccess
            ? Ok(result.Value)
            : NotFound(new { error = result.Error });
    }

    /// <summary>
    /// Lista zadań wysyłki (monitoring / panel admina). Domyślnie zwraca WSZYSTKIE statusy.
    /// </summary>
    /// <param name="status">Pusty / "all" = wszystkie. Albo: Pending | Sending | RetryScheduled | Sent | FailedPermanently | DeadLettered</param>
    /// <param name="skip">Offset paginacji</param>
    /// <param name="take">Rozmiar strony</param>
    [HttpGet("deliveries")]
    [Authorize(Policy = AuthorizationPolicies.RequireAppAdmin)]
    [ProducesResponseType(typeof(IReadOnlyList<DeliveryListItemDto>), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetDeliveries(
        [FromQuery] string status = "", [FromQuery] int skip = 0, [FromQuery] int take = 100)
    {
        var result = await Mediator.Send(new GetDeliveriesByStatusQuery(status, skip, take));

        return result.IsSuccess
            ? Ok(result.Value)
            : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Ręczne ponowienie nieudanego zadania wysyłki (DeadLettered / FailedPermanently).
    /// </summary>
    /// <param name="deliveryId">GUID zadania wysyłki</param>
    [HttpPost("deliveries/{deliveryId:guid}/retry")]
    [Authorize(Policy = AuthorizationPolicies.RequireAppAdmin)]
    [ProducesResponseType(typeof(RequeueDeliveryResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> RetryDelivery(Guid deliveryId)
    {
        var result = await Mediator.Send(new RequeueDeliveryCommand(deliveryId));

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Ręczne anulowanie zadania wysyłki ("Anuluj") — dla zadań oczekujących/zaplanowanych
    /// (Pending / RetryScheduled). Przechodzi w stan końcowy Cancelled.
    /// </summary>
    /// <param name="deliveryId">GUID zadania wysyłki</param>
    [HttpPost("deliveries/{deliveryId:guid}/cancel")]
    [ProducesResponseType(typeof(CancelDeliveryResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CancelDelivery(Guid deliveryId)
    {
        var result = await Mediator.Send(new CancelDeliveryCommand(deliveryId));

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }

    /// <summary>
    /// Zmiana adresu odbiorcy (returnUrl/recipientUrl) zadania wysyłki — panel admina.
    /// Dozwolone dla zadań niewysłanych i nie w trakcie wysyłki.
    /// </summary>
    /// <param name="deliveryId">GUID zadania wysyłki</param>
    [HttpPut("deliveries/{deliveryId:guid}/recipient-url")]
    [ProducesResponseType(typeof(UpdateDeliveryRecipientUrlResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UpdateDeliveryRecipientUrl(
        Guid deliveryId, [FromBody] UpdateDeliveryRecipientUrlRequest request)
    {
        var result = await Mediator.Send(
            new UpdateDeliveryRecipientUrlCommand(deliveryId, request.RecipientUrl));

        return result.IsSuccess
            ? Ok(result.Value)
            : result.IsNotFound
                ? NotFound(new { error = result.Error })
                : BadRequest(new { error = result.Error });
    }
}

// Request DTOs
public record UploadDocumentRequest(string Name, string MimeType, byte[] Content, string? CreatedBy);
public record SaveDocumentVersionRequest(byte[] Content, string? CreatedBy);
public record UpdateDeliveryRecipientUrlRequest(string RecipientUrl);

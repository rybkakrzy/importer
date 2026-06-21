using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.AbortSend;

/// <summary>
/// „Przerwij" po nieudanej pierwszej próbie wysyłki: anuluje zadanie wysyłki (status Cancelled)
/// i ustawia dokument na "UzytkownikPrzerwałWysyłkę" (SendAborted). Dokument zostaje edytowalny,
/// nic nie jest wysyłane w tle. Identyfikowane przez masterId (status to fakt na poziomie master).
/// </summary>
public record AbortSendCommand(Guid MasterId) : IRequest<Result<AbortSendResult>>;

public record AbortSendResult(Guid MasterId, string DocumentStatus, string? DeliveryStatus);

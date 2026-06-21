using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.ContinueDelivery;

/// <summary>
/// „Kontynuuj wysyłkę w tle" po nieudanej pierwszej próbie: przywraca zadanie do kolejki
/// (Requeue → natychmiastowa próba) i ustawia dokument na "Zlecono do wysyłki" (Queued).
/// Dalej dostarcza je worker w tle (retry/backoff). Identyfikowane przez masterId.
/// </summary>
public record ContinueDeliveryCommand(Guid MasterId) : IRequest<Result<ContinueDeliveryResult>>;

public record ContinueDeliveryResult(Guid MasterId, string DocumentStatus, Guid DeliveryId, string DeliveryStatus);

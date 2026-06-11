using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.CancelDelivery;

/// <summary>
/// Ręczne anulowanie zadania wysyłki ("Anuluj") — dotyczy zadań oczekujących/zaplanowanych
/// (Pending / RetryScheduled). Zadanie przechodzi w stan końcowy Cancelled i nie jest dalej wysyłane.
/// </summary>
public record CancelDeliveryCommand(Guid DeliveryId) : IRequest<Result<CancelDeliveryResult>>;

public record CancelDeliveryResult(Guid DeliveryId, string Status);

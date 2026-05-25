using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.RequeueDelivery;

/// <summary>
/// Ręczne ponowienie zadania wysyłki w stanie nieudanym (DeadLettered / FailedPermanently).
/// </summary>
public record RequeueDeliveryCommand(Guid DeliveryId) : IRequest<Result<RequeueDeliveryResult>>;

public record RequeueDeliveryResult(Guid DeliveryId, string Status);

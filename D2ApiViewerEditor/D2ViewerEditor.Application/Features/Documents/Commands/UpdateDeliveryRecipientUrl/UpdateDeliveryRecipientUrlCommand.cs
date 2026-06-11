using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateDeliveryRecipientUrl;

/// <summary>
/// Zmiana adresu odbiorcy (returnUrl/recipientUrl) zadania wysyłki — panel admina „Pliki do wysłania".
/// Dozwolone dla zadań niewysłanych i nie w trakcie wysyłki (nie: Sent/Sending).
/// </summary>
public record UpdateDeliveryRecipientUrlCommand(Guid DeliveryId, string RecipientUrl)
    : IRequest<Result<UpdateDeliveryRecipientUrlResult>>;

public record UpdateDeliveryRecipientUrlResult(Guid DeliveryId, string RecipientUrl, string Status);

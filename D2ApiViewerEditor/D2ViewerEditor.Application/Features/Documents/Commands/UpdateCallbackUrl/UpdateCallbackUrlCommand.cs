using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;

/// <summary>
/// External-API command: set or update the URL the system will POST the finished file to
/// after the editor "Zakończ i wyślij" action. The URL lives in the master document's
/// metadata blob (alongside Classification) — once a DocumentDelivery has been queued,
/// the URL on the master is no longer the source of truth for that in-flight delivery
/// (the delivery captures a frozen RecipientUrl + GCS snapshot).
/// </summary>
public record UpdateCallbackUrlCommand(Guid MasterId, string? CallbackUrl) : IRequest<Result>;

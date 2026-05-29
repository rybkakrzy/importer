using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;

/// <summary>
/// External-API query: snapshot of the document's current lifecycle state, the active
/// version, and the latest delivery attempt (if any). Identified by masterGuid only —
/// status is a master-level fact in this domain (per-version status does not exist).
/// </summary>
public record GetDocumentStatusQuery(Guid MasterId) : IRequest<Result<DocumentStatusDto>>;

/// <summary>
/// Stable wire contract for external integrators. New fields may be added; existing
/// ones must not change meaning (per .ai/API_CONTRACTS).
/// </summary>
public record DocumentStatusDto(
    Guid MasterId,
    string Status,
    bool IsLocked,
    bool HasCallbackUrl,
    Guid? ActiveVersionId,
    int? ActiveVersionNumber,
    DateTime? ActiveVersionModifiedAt,
    DocumentStatusDeliveryDto? LatestDelivery
);

/// <summary>
/// Minimal delivery projection — enough for the source app to know whether the file
/// has been sent. The full RecipientUrl is intentionally not echoed back here (it may
/// embed a token from the original ingest call).
/// </summary>
public record DocumentStatusDeliveryDto(
    Guid DeliveryId,
    string Status,
    int AttemptCount,
    DateTime? LastAttemptAt,
    DateTime? NextAttemptAt,
    DateTime? DeadlineAt
);

using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UnlockDocument;

/// <summary>
/// External-API command: release a document held by the editor, i.e. transition
/// <c>Editing → Saved</c>. There is no separate user-lock entity in this domain — the
/// document-level <c>DocumentStatus.Editing</c> IS the "held by editor" signal (it's also
/// what the GUI surfaces as <c>lockedByOther</c>).
///
/// Idempotent: a document already in <c>Saved</c> is a no-op success. Forbidden during
/// delivery (Sending/Sent/DeliveryFailed) because those are owned by the delivery pipeline.
/// </summary>
public record UnlockDocumentCommand(Guid MasterId, string? Reason) : IRequest<Result<UnlockDocumentResult>>;

public record UnlockDocumentResult(Guid MasterId, bool Changed);

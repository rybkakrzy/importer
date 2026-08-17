using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.DeleteDocument;

/// <summary>
/// Panel administratora: TRWAŁE usunięcie dokumentu — bloby WSZYSTKICH wersji z magazynu (GCS)
/// oraz wpis z bazy (wersje i zadania wysyłki schodzą kaskadą — infra/sql: ON DELETE CASCADE).
/// Zablokowane w stanach pipeline'u wysyłki (Queued/Sending) — runner równolegle pracuje na
/// dokumencie; administrator najpierw anuluje/przerywa wysyłkę istniejącymi operacjami.
/// </summary>
public record DeleteDocumentCommand(Guid MasterId) : IRequest<Result<bool>>;

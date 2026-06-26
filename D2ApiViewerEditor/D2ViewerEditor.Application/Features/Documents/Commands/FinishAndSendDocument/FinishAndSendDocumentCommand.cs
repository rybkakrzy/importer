using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;

/// <summary>
/// Komenda "zakończ i wyślij": utrwala stan edytora, zamraża snapshot finalnego pliku, tworzy
/// zadanie wysyłki i wykonuje SYNCHRONICZNĄ pierwszą próbę dostarczenia na returnUrl. Wynik tej
/// próby (sukces/błąd) jest zwracany od razu, by GUI mogło pokazać "Wysłano" albo zaproponować
/// „Przerwij" / „Kontynuuj wysyłkę w tle".
/// </summary>
public record FinishAndSendDocumentCommand(
    Guid MasterId,
    Guid VersionId,
    byte[] Content,
    string? CreatedBy
) : IRequest<Result<FinishAndSendResult>>;

/// <summary>
/// Wynik kończenia dokumentu: identyfikator zadania, status zadania wysyłki, status dokumentu
/// po pierwszej próbie oraz flaga, czy dokument został dostarczony za pierwszym razem.
/// </summary>
public record FinishAndSendResult(
    Guid DeliveryId,
    string Status,
    string DocumentStatus,
    bool Delivered,
    string? Error = null
);

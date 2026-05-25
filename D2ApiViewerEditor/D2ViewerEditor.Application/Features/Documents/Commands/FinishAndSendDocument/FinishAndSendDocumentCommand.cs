using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;

/// <summary>
/// Komenda "zakończ i wyślij": utrwala stan edytora, zamraża snapshot finalnego pliku
/// i tworzy zadanie asynchronicznej wysyłki na returnUrl z metadanych dokumentu.
/// </summary>
public record FinishAndSendDocumentCommand(
    Guid MasterId,
    Guid VersionId,
    byte[] Content,
    string? CreatedBy
) : IRequest<Result<FinishAndSendResult>>;

/// <summary>
/// Wynik utworzenia (lub odnalezienia istniejącego) zadania wysyłki.
/// </summary>
public record FinishAndSendResult(
    Guid DeliveryId,
    string Status
);

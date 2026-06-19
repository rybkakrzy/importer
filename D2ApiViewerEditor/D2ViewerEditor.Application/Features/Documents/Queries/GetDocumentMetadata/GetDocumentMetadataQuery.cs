using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;

/// <summary>
/// Zwraca metadane przysłane przez aplikację zewnętrzną (returnUrl, classification).
/// Wykorzystywane m.in. w kroku „Zakończ" do odesłania pliku.
/// </summary>
public record GetDocumentMetadataQuery(Guid MasterId) : IRequest<Result<DocumentMetadataDto>>;

/// <summary>
/// Zdeserializowane metadane dokumentu.
/// </summary>
public record DocumentMetadataDto(
    Guid MasterId,
    string MimeType,
    string? ReturnUrl,
    string? Classification,
    // Per domain rule: only an explicit `true` allows the user-facing download action.
    // Missing field / non-true value ⇒ false. Surfaced to the GUI so the menu item can
    // be hidden; backend enforcement lives in the dedicated download command.
    bool UserDownload,
    // Inverse-default: visible unless the source app sent an explicit `false`. Missing
    // metadata ⇒ true (backward compatible). Drives visibility of the editor's save-state
    // UI (autosave section + manual "Zapisz" button).
    bool ShowSaveState
);

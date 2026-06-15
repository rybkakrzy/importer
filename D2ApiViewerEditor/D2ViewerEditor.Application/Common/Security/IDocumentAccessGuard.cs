namespace D2ViewerEditor.Application.Common.Security;

/// <summary>
/// Decides whether the current user may view a document, given its raw metadata JSON
/// (which may carry an <c>allowedCorporateKeys</c> allow-list). Single entry point used by
/// every read/content handler so the rule stays in one place.
/// </summary>
public interface IDocumentAccessGuard
{
    bool IsViewAllowed(string? metadataJson);
}

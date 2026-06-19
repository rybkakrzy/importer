using System.Text.Json;

namespace D2ViewerEditor.Application.Features.Documents.Common;

/// <summary>
/// Wire shape of the JSON blob stored in <c>documents.metadata</c>. The blob is set by
/// either the external ingest controller (form-data → JSON) or the local-upload handler
/// (system-controlled). Every property is optional in the wire format; missing fields
/// fall back to safe defaults via <see cref="Parse"/>.
///
/// <para><c>UserDownload</c> follows the domain rule: missing / null / non-true ⇒ false.
/// Only an explicit <c>true</c> enables the user-facing "download to disk" action.</para>
///
/// <para><c>ShowSaveState</c> is the inverse-default flag: missing / null / non-false ⇒ visible.
/// Only an explicit <c>false</c> hides the editor's save-state UI (autosave + manual "Zapisz").</para>
/// </summary>
public sealed record ExternalDocumentMetadata(
    string? ReturnUrl,
    string? Classification,
    bool? UserDownload,
    bool? ShowSaveState = null)
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public static ExternalDocumentMetadata Empty { get; } = new(null, null, null);

    /// <summary>
    /// Tolerant parser — malformed / null / empty metadata returns <see cref="Empty"/>
    /// instead of throwing, so a single bad row never breaks read endpoints.
    /// </summary>
    public static ExternalDocumentMetadata Parse(string? metadataJson)
    {
        if (string.IsNullOrWhiteSpace(metadataJson)) return Empty;
        try
        {
            return JsonSerializer.Deserialize<ExternalDocumentMetadata>(metadataJson, JsonOptions) ?? Empty;
        }
        catch (JsonException)
        {
            return Empty;
        }
    }

    public string Serialize() => JsonSerializer.Serialize(this, JsonOptions);

    /// <summary>
    /// Domain rule: download is only allowed when <c>UserDownload</c> is explicitly true.
    /// </summary>
    public bool IsUserDownloadAllowed => UserDownload == true;

    /// <summary>
    /// Inverse-default rule: the save-state UI stays visible unless <c>ShowSaveState</c> is
    /// explicitly false. Missing / null ⇒ visible (backward compatible with old metadata).
    /// </summary>
    public bool IsSaveStateVisible => ShowSaveState != false;
}

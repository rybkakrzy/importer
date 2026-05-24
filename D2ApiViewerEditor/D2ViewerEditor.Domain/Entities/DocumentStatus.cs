namespace D2ViewerEditor.Domain.Entities;

/// <summary>
/// Document lifecycle status (from ingest to delivery back to the source app).
/// Stored in the database as a string (see DocumentConfiguration).
/// </summary>
public enum DocumentStatus
{
    /// <summary>Ingested and stored (no editing yet).</summary>
    Saved,

    /// <summary>Editable version is being edited (save/auto-save from the editor).</summary>
    Editing,

    /// <summary>File is being sent back to the source app (returnUrl).</summary>
    Sending,

    /// <summary>Delivery to the recipient (returnUrl) failed — recipient not responding.</summary>
    DeliveryFailed,

    /// <summary>File has been sent back to the source app.</summary>
    Sent
}

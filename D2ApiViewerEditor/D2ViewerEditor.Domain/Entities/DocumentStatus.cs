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

    /// <summary>Delivery has been requested ("Zlecono do wysyłki") — the document is queued for
    /// sending but the actual transfer has not started yet (also used when the user chooses to
    /// continue delivery in the background).</summary>
    Queued,

    /// <summary>File is being sent back to the source app — first attempt in progress
    /// ("W trakcie wysyłki").</summary>
    Sending,

    /// <summary>Delivery to the recipient (returnUrl) failed — recipient not responding
    /// ("Błąd wysyłki").</summary>
    DeliveryFailed,

    /// <summary>File has been sent back to the source app ("Wysłano").</summary>
    Sent,

    /// <summary>The user aborted delivery after a failed first attempt ("UzytkownikPrzerwałWysyłkę").
    /// The matching delivery job is cancelled and the document stays editable.</summary>
    SendAborted
}

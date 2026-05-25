namespace D2ViewerEditor.Infrastructure.Services.Delivery;

/// <summary>
/// Konfiguracja workera wysyłki (sekcja "DeliveryWorker").
/// </summary>
public class DeliveryWorkerOptions
{
    public const string SectionName = "DeliveryWorker";

    /// <summary>Czy worker ma działać w tej instancji (pozwala wydzielić wysyłkę na dedykowany host).</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>Odstęp między cyklami pollingu, gdy kolejka jest pusta.</summary>
    public TimeSpan PollInterval { get; set; } = TimeSpan.FromSeconds(15);

    /// <summary>Maksymalna liczba zadań claimowanych w jednym batchu.</summary>
    public int BatchSize { get; set; } = 20;

    /// <summary>Maksymalna równoległość wysyłki w obrębie batcha.</summary>
    public int MaxConcurrency { get; set; } = 4;

    /// <summary>Lease claimu — po tym czasie zawieszone zadanie (Sending) może przejąć inny worker.</summary>
    public TimeSpan Lease { get; set; } = TimeSpan.FromMinutes(5);

    /// <summary>Okno ponawiania zadania (od utworzenia do dead-letter).</summary>
    public TimeSpan RetentionWindow { get; set; } = TimeSpan.FromHours(24);

    /// <summary>Timeout pojedynczej próby wysyłki HTTP.</summary>
    public TimeSpan HttpTimeout { get; set; } = TimeSpan.FromSeconds(30);
}

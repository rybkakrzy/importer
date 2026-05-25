namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Strategia wyliczania terminu kolejnej próby wysyłki.
/// </summary>
public interface IBackoffStrategy
{
    /// <summary>
    /// Zwraca moment kolejnej próby na podstawie liczby dotychczasowych prób.
    /// </summary>
    /// <param name="attemptCount">Liczba prób wykonanych do tej pory (>= 1).</param>
    /// <param name="now">Aktualny czas (UTC).</param>
    DateTime NextAttempt(int attemptCount, DateTime now);
}

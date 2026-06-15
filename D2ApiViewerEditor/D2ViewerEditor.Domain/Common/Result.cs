namespace D2ViewerEditor.Domain.Common;

/// <summary>
/// Wzorzec Result — enkapsuluje wynik operacji (sukces/porażka) bez rzucania wyjątków
/// </summary>
public class Result<T>
{
    private Result(T value)
    {
        IsSuccess = true;
        Value = value;
        Error = null;
        IsNotFound = false;
    }

    private Result(string error, bool isNotFound = false, bool isForbidden = false)
    {
        IsSuccess = false;
        Value = default;
        Error = error;
        IsNotFound = isNotFound;
        IsForbidden = isForbidden;
    }

    public bool IsSuccess { get; }
    public bool IsFailure => !IsSuccess;
    public bool IsNotFound { get; }
    public bool IsForbidden { get; }
    public T? Value { get; }
    public string? Error { get; }

    public static Result<T> Success(T value) => new(value);
    public static Result<T> Failure(string error) => new(error);
    public static Result<T> NotFound(string error = "Nie znaleziono dokumentu") => new(error, isNotFound: true);
    public static Result<T> Forbidden(string error = "Brak uprawnień do dokumentu") => new(error, isForbidden: true);

    public TResult Match<TResult>(Func<T, TResult> onSuccess, Func<string, TResult> onFailure)
        => IsSuccess ? onSuccess(Value!) : onFailure(Error!);
}

/// <summary>
/// Result bez wartości — dla operacji void
/// </summary>
public class Result
{
    private Result(bool isSuccess, string? error, bool isNotFound = false, bool isForbidden = false)
    {
        IsSuccess = isSuccess;
        Error = error;
        IsNotFound = isNotFound;
        IsForbidden = isForbidden;
    }

    public bool IsSuccess { get; }
    public bool IsFailure => !IsSuccess;
    public bool IsNotFound { get; }
    public bool IsForbidden { get; }
    public string? Error { get; }

    public static Result Success() => new(true, null);
    public static Result Failure(string error) => new(false, error);
    public static Result NotFound(string error = "Nie znaleziono dokumentu") => new(false, error, isNotFound: true);
    public static Result Forbidden(string error = "Brak uprawnień do dokumentu") => new(false, error, isForbidden: true);

    public TResult Match<TResult>(Func<TResult> onSuccess, Func<string, TResult> onFailure)
        => IsSuccess ? onSuccess() : onFailure(Error!);
}

namespace D2ViewerEditor.Application.Common.Security;

public enum ReturnUrlRejectionCode
{
    None = 0,
    Empty,
    TooLong,
    ContainsControlCharacters,
    DangerousCharacters,
    ProtocolRelative,
    InvalidAbsoluteUri,
    InvalidScheme,
    InsecureScheme,
    UserInfoNotAllowed,
    LoopbackNotAllowed,
    PrivateNetworkIpNotAllowed,
    HostNotAllowListed
}

public sealed record ReturnUrlValidationResult(
    bool IsValid,
    string? NormalizedUrl,
    ReturnUrlRejectionCode Code,
    string? Error)
{
    public static ReturnUrlValidationResult Success(string normalizedUrl) =>
        new(true, normalizedUrl, ReturnUrlRejectionCode.None, null);

    public static ReturnUrlValidationResult Failure(ReturnUrlRejectionCode code, string error) =>
        new(false, null, code, error);
}

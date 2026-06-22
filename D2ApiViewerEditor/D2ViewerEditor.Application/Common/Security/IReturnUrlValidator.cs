namespace D2ViewerEditor.Application.Common.Security;

public interface IReturnUrlValidator
{
    ReturnUrlValidationResult Validate(string? rawUrl);
}

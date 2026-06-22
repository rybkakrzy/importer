using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Application.Common.Security;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UploadImage;

public class UploadImageCommandHandler : IRequestHandler<UploadImageCommand, Result<ImageUploadResponse>>
{
    private readonly IFileUploadSecurityService _uploadSecurityService;

    public UploadImageCommandHandler(IFileUploadSecurityService uploadSecurityService)
    {
        _uploadSecurityService = uploadSecurityService;
    }

    public async Task<Result<ImageUploadResponse>> Handle(UploadImageCommand request, CancellationToken cancellationToken)
    {
        using var memoryStream = new MemoryStream();
        await request.FileStream.CopyToAsync(memoryStream, cancellationToken);
        var bytes = memoryStream.ToArray();

        var uploadValidation = await _uploadSecurityService.ValidateImageAsync(
            bytes,
            request.FileName,
            request.ContentType,
            cancellationToken);

        if (!uploadValidation.IsValid)
            return Result<ImageUploadResponse>.Failure($"Upload odrzucony ({uploadValidation.Code}): {uploadValidation.Error}");

        var base64 = Convert.ToBase64String(bytes);
        var normalizedMimeType = uploadValidation.NormalizedMimeType ?? request.ContentType;

        var response = new ImageUploadResponse(
            $"data:{normalizedMimeType};base64,{base64}",
            request.FileName,
            request.FileSize
        );

        return Result<ImageUploadResponse>.Success(response);
    }
}

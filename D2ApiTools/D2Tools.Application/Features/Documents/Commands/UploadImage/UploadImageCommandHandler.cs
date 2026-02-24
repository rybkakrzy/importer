using D2Tools.Domain.Common;
using MediatR;

namespace D2Tools.Application.Features.Documents.Commands.UploadImage;

public class UploadImageCommandHandler : IRequestHandler<UploadImageCommand, Result<ImageUploadResponse>>
{
    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"
    };

    public async Task<Result<ImageUploadResponse>> Handle(UploadImageCommand request, CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(request.FileName).ToLower();
        if (!AllowedExtensions.Contains(extension))
            return Result<ImageUploadResponse>.Failure("Niedozwolony format obrazu");

        using var memoryStream = new MemoryStream();
        await request.FileStream.CopyToAsync(memoryStream, cancellationToken);
        var base64 = Convert.ToBase64String(memoryStream.ToArray());

        var response = new ImageUploadResponse(
            $"data:{request.ContentType};base64,{base64}",
            request.FileName,
            request.FileSize
        );

        return Result<ImageUploadResponse>.Success(response);
    }
}

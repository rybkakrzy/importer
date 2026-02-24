using System.IO.Compression;
using Importer.Domain.Common;
using Importer.Domain.Interfaces;
using MediatR;

namespace Importer.Application.Features.FileUpload.Commands.ProcessZipFile;

public class ProcessZipFileCommandHandler : IRequestHandler<ProcessZipFileCommand, Result<ZipUploadResult>>
{
    private readonly IDocxToHtmlConverter _converter;
    private readonly IUploadSessionService _sessionService;

    public ProcessZipFileCommandHandler(IDocxToHtmlConverter converter, IUploadSessionService sessionService)
    {
        _converter = converter;
        _sessionService = sessionService;
    }

    public async Task<Result<ZipUploadResult>> Handle(ProcessZipFileCommand request, CancellationToken cancellationToken)
    {
        var sessionId = _sessionService.CreateSession(request.FileName);

        try
        {
            using var memoryStream = new MemoryStream();
            await request.FileStream.CopyToAsync(memoryStream, cancellationToken);
            memoryStream.Position = 0;

            var files = new List<ZipFileEntryDto>();
            long totalSize = 0;

            using var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read);
            foreach (var entry in archive.Entries)
            {
                if (string.IsNullOrEmpty(entry.Name)) continue;

                totalSize += entry.Length;
                string? content = null;

                if (entry.Name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    using var entryStream = entry.Open();
                    using var entryMemoryStream = new MemoryStream();
                    await entryStream.CopyToAsync(entryMemoryStream, cancellationToken);
                    entryMemoryStream.Position = 0;

                    var documentContent = _converter.Convert(entryMemoryStream);
                    content = documentContent.Html;
                }

                files.Add(new ZipFileEntryDto(entry.FullName, entry.Length, content));
            }

            _sessionService.CompleteSession(sessionId, files.Count, totalSize);
            var result = new ZipUploadResult(files, files.Count, totalSize);
            return Result<ZipUploadResult>.Success(result);
        }
        catch (Exception ex)
        {
            _sessionService.FailSession(sessionId, ex.Message);
            throw;
        }
    }
}

using Importer.Domain.Models;

namespace Importer.Domain.Interfaces;

public interface IUploadSessionService
{
    Guid CreateSession(string fileName);
    void CompleteSession(Guid sessionId, int filesCount, long totalSize);
    void FailSession(Guid sessionId, string errorMessage);
    IReadOnlyList<UploadSession> GetAllSessions();
}

using System.Collections.Concurrent;
using Importer.Domain.Interfaces;
using Importer.Domain.Models;

namespace Importer.Infrastructure.Services;

public class InMemoryUploadSessionService : IUploadSessionService
{
    private readonly ConcurrentDictionary<Guid, UploadSession> _sessions = new();

    public Guid CreateSession(string fileName)
    {
        var id = Guid.NewGuid();
        var session = new UploadSession(id, fileName, "Processing", DateTime.UtcNow, null, null, null, null);
        _sessions[id] = session;
        return id;
    }

    public void CompleteSession(Guid sessionId, int filesCount, long totalSize)
    {
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            _sessions[sessionId] = session with
            {
                Status = "Completed",
                CompletedAt = DateTime.UtcNow,
                FilesCount = filesCount,
                TotalSize = totalSize
            };
        }
    }

    public void FailSession(Guid sessionId, string errorMessage)
    {
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            _sessions[sessionId] = session with
            {
                Status = "Failed",
                CompletedAt = DateTime.UtcNow,
                ErrorMessage = errorMessage
            };
        }
    }

    public IReadOnlyList<UploadSession> GetAllSessions()
        => _sessions.Values.OrderByDescending(s => s.StartedAt).ToList();
}

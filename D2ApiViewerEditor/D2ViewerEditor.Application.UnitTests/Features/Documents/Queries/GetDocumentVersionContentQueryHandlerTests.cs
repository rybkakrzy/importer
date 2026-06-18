using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersionContent;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentVersionContentQueryHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private IDocumentRepository _repo = null!;
    private IDocumentStorageService _storage = null!;
    private IDocumentAccessGuard _accessGuard = null!;
    private GetDocumentVersionContentQueryHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _repo = Substitute.For<IDocumentRepository>();
        _storage = Substitute.For<IDocumentStorageService>();
        _accessGuard = Substitute.For<IDocumentAccessGuard>();
        _accessGuard.IsViewAllowed(Arg.Any<string?>()).Returns(true);
        _handler = new GetDocumentVersionContentQueryHandler(_repo, _storage, _accessGuard);
    }

    private static (Document doc, DocumentVersion version) DocWithVersion()
    {
        var doc = new Document(Guid.NewGuid(), "raport", DocxMime, "User", null);
        var v = doc.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User");
        return (doc, v);
    }

    [Test]
    public async Task Handle_ExistingVersion_DownloadsContentAndBuildsFileName()
    {
        var (doc, version) = DocWithVersion();
        var bytes = new byte[] { 9, 8, 7 };
        _repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        _storage.DownloadAsync(version.StoragePath, Arg.Any<CancellationToken>()).Returns(bytes);

        var result = await _handler.Handle(
            new GetDocumentVersionContentQuery(doc.Id, version.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Content.Should().BeEquivalentTo(bytes);
        result.Value.MimeType.Should().Be(DocxMime);
        result.Value.FileName.Should().Be($"raport_v{version.VersionNumber}.docx");
    }

    [Test]
    public async Task Handle_DocumentNotFound_ReturnsNotFound()
    {
        _repo.GetByIdWithVersionsAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var result = await _handler.Handle(
            new GetDocumentVersionContentQuery(Guid.NewGuid(), Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_VersionNotFound_ReturnsNotFound()
    {
        var (doc, _) = DocWithVersion();
        _repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);

        var result = await _handler.Handle(
            new GetDocumentVersionContentQuery(doc.Id, Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
        await _storage.DidNotReceive().DownloadAsync(Arg.Any<string>(), Arg.Any<CancellationToken>());
    }
}

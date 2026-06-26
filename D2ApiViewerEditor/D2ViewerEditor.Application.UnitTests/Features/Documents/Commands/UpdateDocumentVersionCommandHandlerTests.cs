using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Commands.UpdateDocumentVersion;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UpdateDocumentVersionCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private IDocumentRepository _repo = null!;
    private IDocumentStorageService _storage = null!;
    private ICurrentUserProvider _currentUser = null!;
    private UpdateDocumentVersionCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _repo = Substitute.For<IDocumentRepository>();
        _storage = Substitute.For<IDocumentStorageService>();
        _storage.UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns(ci => $"documents/{ci.ArgAt<Guid>(0)}");
        _currentUser = Substitute.For<ICurrentUserProvider>();
        _currentUser.CorporateKey.Returns("CORP-1");
        _handler = new UpdateDocumentVersionCommandHandler(_repo, _storage, _currentUser);
    }

    private static (Document doc, DocumentVersion v1, DocumentVersion v2) DocWithTwoVersions()
    {
        var doc = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", null);
        var v1 = doc.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User");
        var v2 = doc.AddVersion(Guid.NewGuid(), "documents/v2", 100, "User");
        return (doc, v1, v2);
    }

    [Test]
    public async Task Handle_EditableVersion_OverwritesInPlaceAndMarksEditing()
    {
        var (doc, _, v2) = DocWithTwoVersions();
        _repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        var content = new byte[] { 1, 2, 3, 4, 5 };

        var result = await _handler.Handle(
            new UpdateDocumentVersionCommand(doc.Id, v2.Id, content), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.VersionId.Should().Be(v2.Id);
        result.Value.VersionNumber.Should().Be(2);
        result.Value.SizeInBytes.Should().Be(content.Length);
        doc.Status.Should().Be(DocumentStatus.Editing);
        await _storage.Received(1).UploadAsync(v2.Id, content, DocxMime, Arg.Any<CancellationToken>());
        await _repo.Received(1).SaveChangesAsync(Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_EmptyContent_ReturnsFailure_WithoutTouchingRepo()
    {
        var result = await _handler.Handle(
            new UpdateDocumentVersionCommand(Guid.NewGuid(), Guid.NewGuid(), Array.Empty<byte>()),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        await _repo.DidNotReceive().GetByIdWithVersionsAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_DocumentNotFound_ReturnsNotFound()
    {
        _repo.GetByIdWithVersionsAsync(Arg.Any<Guid>(), Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var result = await _handler.Handle(
            new UpdateDocumentVersionCommand(Guid.NewGuid(), Guid.NewGuid(), new byte[] { 1 }),
            CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_VersionNotFound_ReturnsNotFound()
    {
        var (doc, _, _) = DocWithTwoVersions();
        _repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);

        var result = await _handler.Handle(
            new UpdateDocumentVersionCommand(doc.Id, Guid.NewGuid(), new byte[] { 1 }),
            CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_OriginalVersionV1_ReturnsFailure_NeverOverwrites()
    {
        var (doc, v1, _) = DocWithTwoVersions();
        _repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);

        var result = await _handler.Handle(
            new UpdateDocumentVersionCommand(doc.Id, v1.Id, new byte[] { 1, 2 }),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue(); // domena: v1 nietykalna → InvalidOperationException → Failure
        await _repo.DidNotReceive().SaveChangesAsync(Arg.Any<CancellationToken>());
    }
}

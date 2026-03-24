using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class RestoreDocumentVersionCommandHandlerTests
{
    private IDocumentRepository _documentRepository;
    private RestoreDocumentVersionCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _handler = new RestoreDocumentVersionCommandHandler(_documentRepository);
    }

    [Test]
    public async Task Handle_ValidVersionId_ShouldRestoreVersion()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "contract.pdf", "application/pdf", "Lawyer");
        
        var version1 = document.AddVersion(new byte[] { 1, 2 }, "Lawyer");
        var version2 = document.AddVersion(new byte[] { 3, 4 }, "Lawyer");
        var version3 = document.AddVersion(new byte[] { 5, 6 }, "Lawyer");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var command = new RestoreDocumentVersionCommand(
            MasterId: masterId,
            VersionId: version1.Id
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        version1.IsActive.Should().BeTrue();
        version2.IsActive.Should().BeFalse();
        version3.IsActive.Should().BeFalse();

        await _documentRepository.Received(1).UpdateAsync(document, Arg.Any<CancellationToken>());
        await _documentRepository.Received(1).SaveChangesAsync(Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_DocumentNotFound_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var versionId = Guid.NewGuid();

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var command = new RestoreDocumentVersionCommand(masterId, versionId);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_VersionNotFound_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "file.txt", "text/plain", "User");
        document.AddVersion(new byte[] { 1 }, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var nonExistentVersionId = Guid.NewGuid();
        var command = new RestoreDocumentVersionCommand(masterId, nonExistentVersionId);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Wersja");
        result.Error.Should().Contain("nie należy do dokumentu");
    }

    [Test]
    public async Task Handle_RestoreOldVersion_ShouldKeepAllVersionsInHistory()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "policy.doc", "application/msword", "Admin");
        
        var v1 = document.AddVersion(new byte[] { 1 }, "Admin");
        var v2 = document.AddVersion(new byte[] { 2 }, "Admin");
        var v3 = document.AddVersion(new byte[] { 3 }, "Admin");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var command = new RestoreDocumentVersionCommand(masterId, v1.Id);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        document.Versions.Should().HaveCount(3); // Wszystkie wersje zachowane
        document.GetActiveVersion()!.Id.Should().Be(v1.Id);
    }

    [Test]
    public async Task Handle_RepositoryException_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "test.pdf", "application/pdf", "User");
        var version = document.AddVersion(new byte[] { 1 }, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _documentRepository.When(x => x.SaveChangesAsync(Arg.Any<CancellationToken>()))
            .Do(x => throw new Exception("Database deadlock"));

        var command = new RestoreDocumentVersionCommand(masterId, version.Id);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Database deadlock");
    }

    [Test]
    public async Task Handle_RestoreMiddleVersion_ShouldWork()
    {
        // Arrange - Scenariusz: użytkownik edytował dokument 5 razy, chce wrócić do wersji 3
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "article.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Writer");
        
        var v1 = document.AddVersion(new byte[] { 1 }, "Writer"); // Draft
        var v2 = document.AddVersion(new byte[] { 2 }, "Writer"); // First revision
        var v3 = document.AddVersion(new byte[] { 3 }, "Writer"); // Good version ← chcemy wrócić
        var v4 = document.AddVersion(new byte[] { 4 }, "Writer"); // Bad edit
        var v5 = document.AddVersion(new byte[] { 5 }, "Writer"); // Worse edit

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var command = new RestoreDocumentVersionCommand(masterId, v3.Id);

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        v3.IsActive.Should().BeTrue();
        v5.IsActive.Should().BeFalse();
        document.Versions.Should().HaveCount(5); // Historia zachowana
    }
}

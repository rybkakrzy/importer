using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class SaveDocumentVersionCommandHandlerTests
{
    private IDocumentRepository _documentRepository;
    private IDocumentStorageService _storageService;
    private ICurrentUserProvider _currentUser;
    private SaveDocumentVersionCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _storageService = Substitute.For<IDocumentStorageService>();
        _storageService.UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns(ci => $"documents/{ci.ArgAt<Guid>(0)}");
        _currentUser = Substitute.For<ICurrentUserProvider>();
        _currentUser.CorporateKey.Returns("CORP-1");
        _handler = new SaveDocumentVersionCommandHandler(_documentRepository, _storageService, _currentUser);
    }

    [Test]
    public async Task Handle_ValidCommand_ShouldAddNewVersion()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var existingDocument = new Document(
            id: masterId,
            name: "test.pdf",
            mimeType: "application/pdf",
            createdBy: "User1"
        );
        existingDocument.AddVersion("documents/existing/v1", 100, "User1");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(existingDocument);

        var command = new SaveDocumentVersionCommand(
            MasterId: masterId,
            Content: new byte[] { 3, 4, 5 },
            CreatedBy: "User2"
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().NotBeNull();
        result.Value!.VersionNumber.Should().Be(2); // Druga wersja
        result.Value.VersionId.Should().NotBeEmpty();

        await _storageService.Received(1).UploadAsync(
            Arg.Any<Guid>(), Arg.Any<byte[]>(), "application/pdf", Arg.Any<CancellationToken>());
        // Handler celowo NIE woła UpdateAsync — encja jest śledzona, a pełny Update agregatu
        // wymusiłby UPDATE documents z created_at jako Kind=Unspecified (błąd Npgsql / timestamptz).
        await _documentRepository.DidNotReceive().UpdateAsync(Arg.Any<Document>(), Arg.Any<CancellationToken>());
        await _documentRepository.Received(1).SaveChangesAsync(Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_DocumentNotFound_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var command = new SaveDocumentVersionCommand(
            MasterId: masterId,
            Content: new byte[] { 1, 2, 3 },
            CreatedBy: "User"
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_MultipleVersions_ShouldIncrementVersionNumber()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "doc.txt", "text/plain", "User");
        document.AddVersion("documents/test/v1", 10, "User");
        document.AddVersion("documents/test/v2", 20, "User");
        document.AddVersion("documents/test/v3", 30, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var command = new SaveDocumentVersionCommand(
            MasterId: masterId,
            Content: new byte[] { 4, 5 },
            CreatedBy: "User"
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.VersionNumber.Should().Be(4);
        document.Versions.Should().HaveCount(4);
    }

    [Test]
    public async Task Handle_RepositoryException_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "test.doc", "application/msword", "User");
        document.AddVersion("documents/test/v1", 10, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _documentRepository.When(x => x.SaveChangesAsync(Arg.Any<CancellationToken>()))
            .Do(x => throw new Exception("Connection timeout"));

        var command = new SaveDocumentVersionCommand(masterId, new byte[] { 2 }, "User");

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Connection timeout");
    }

    [Test]
    public async Task Handle_NewVersionShouldBecomeActive()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Analyst");
        document.AddVersion("documents/test/v1", 100, "Analyst");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var command = new SaveDocumentVersionCommand(masterId, new byte[] { 4, 5, 6 }, "Analyst");

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        
        // Sprawdź że poprzednia wersja została dezaktywowana
        var activeVersions = document.Versions.Where(v => v.IsActive).ToList();
        activeVersions.Should().HaveCount(1);
        activeVersions.First().VersionNumber.Should().Be(2);
    }
}

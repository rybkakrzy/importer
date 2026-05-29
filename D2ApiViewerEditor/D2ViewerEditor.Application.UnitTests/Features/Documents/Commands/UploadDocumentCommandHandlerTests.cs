using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UploadDocumentCommandHandlerTests
{
    private IDocumentRepository _documentRepository;
    private IDocumentStorageService _storageService;
    private UploadDocumentCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _storageService = Substitute.For<IDocumentStorageService>();
        _storageService.UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns(ci => $"documents/{ci.ArgAt<Guid>(0)}");
        _handler = new UploadDocumentCommandHandler(_documentRepository, _storageService);
    }

    [Test]
    public async Task Handle_ValidCommand_ShouldCreateDocumentWithFirstVersion()
    {
        // Arrange
        var content = new byte[] { 1, 2, 3, 4, 5 };
        var command = new UploadDocumentCommand(
            Content: content,
            FileName: "test.pdf",
            MimeType: "application/pdf",
            CreatedBy: "TestUser"
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().NotBeNull();
        result.Value!.MasterId.Should().NotBeEmpty();
        result.Value.VersionId.Should().NotBeEmpty();
        result.Value.FileName.Should().Be("test.pdf");
        result.Value.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));

        await _storageService.Received(1).UploadAsync(
            Arg.Any<Guid>(), content, "application/pdf", Arg.Any<CancellationToken>());
        await _documentRepository.Received(1).AddAsync(
            Arg.Is<Document>(d => d.Name == "test.pdf" && d.MimeType == "application/pdf"),
            Arg.Any<CancellationToken>()
        );
        await _documentRepository.Received(1).SaveChangesAsync(Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_LocalUpload_ShouldAutomaticallySetUserDownloadTrue()
    {
        // Local upload: no external app and no return URL, so the only way to retrieve the
        // edited file is the in-browser download. The flag is set by the SYSTEM (not the
        // request) so a tampered upload cannot suppress it.
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1, 2, 3 },
            FileName: "local.docx",
            MimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            CreatedBy: "TestUser"
        );

        await _handler.Handle(command, CancellationToken.None);

        await _documentRepository.Received(1).AddAsync(
            Arg.Is<Document>(d => d.Metadata != null && d.Metadata.Contains("\"userDownload\":true")),
            Arg.Any<CancellationToken>()
        );
    }

    [Test]
    public async Task Handle_ValidCommand_ShouldCreateDocumentWithVersionNumber1()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1, 2, 3 },
            FileName: "document.docx",
            MimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            CreatedBy: "Admin"
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        
        // Weryfikacja że dokument został utworzony z pierwszą wersją
        await _documentRepository.Received(1).AddAsync(
            Arg.Is<Document>(d => d.Versions.Count == 1 && d.Versions.First().VersionNumber == 1),
            Arg.Any<CancellationToken>()
        );
    }

    [Test]
    public async Task Handle_RepositoryThrowsException_ShouldReturnFailure()
    {
        // Arrange
        var command = new UploadDocumentCommand(
            Content: new byte[] { 1 },
            FileName: "test.pdf",
            MimeType: "application/pdf",
            CreatedBy: "User"
        );

        _documentRepository.When(x => x.SaveChangesAsync(Arg.Any<CancellationToken>()))
            .Do(x => throw new Exception("Database error"));

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Database error");
    }

    [Test]
    public async Task Handle_EmptyContent_ShouldReturnFailure()
    {
        // Arrange - Domain logic powinien odrzucić pusty content
        var command = new UploadDocumentCommand(
            Content: Array.Empty<byte>(),
            FileName: "empty.txt",
            MimeType: "text/plain",
            CreatedBy: "System"
        );

        _storageService.UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns("documents/test/v1");

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
    }
}

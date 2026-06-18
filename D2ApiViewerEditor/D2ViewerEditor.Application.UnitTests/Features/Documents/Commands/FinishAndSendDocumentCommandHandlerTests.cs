using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class FinishAndSendDocumentCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private Mock<IDocumentStorageService> _storage = null!;
    private Mock<ICurrentUserProvider> _currentUser = null!;
    private FinishAndSendDocumentCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _storage = new Mock<IDocumentStorageService>();
        _currentUser = new Mock<ICurrentUserProvider>();
        _handler = new FinishAndSendDocumentCommandHandler(
            _documentRepo.Object, _deliveryRepo.Object, _storage.Object, _currentUser.Object);
    }

    private static Document BuildDocumentWithEditableVersion(out Guid versionId, string? metadata)
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata);
        document.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User"); // v1 (oryginał)
        var v2 = document.AddVersion(Guid.NewGuid(), "documents/v2", 100, "User"); // v2 (edytowalna)
        versionId = v2.Id;
        return document;
    }

    [Test]
    public async Task Handle_HappyPath_ShouldCreateDeliveryAndFreezeSnapshot()
    {
        var document = BuildDocumentWithEditableVersion(out var versionId,
            "{\"returnUrl\":\"https://example.com/cb\",\"classification\":\"C2\"}");

        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);
        _currentUser.SetupGet(u => u.CorporateKey).Returns("ACME-42");

        DocumentDelivery? created = null;
        _deliveryRepo.Setup(r => r.AddAsync(It.IsAny<DocumentDelivery>(), It.IsAny<CancellationToken>()))
            .Callback<DocumentDelivery, CancellationToken>((d, _) => created = d);

        var command = new FinishAndSendDocumentCommand(document.Id, versionId, new byte[] { 1, 2, 3 }, "User");
        var result = await _handler.Handle(command, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Status.Should().Be(nameof(DeliveryStatus.Pending));
        document.Status.Should().Be(DocumentStatus.Sending);

        // Delivery carries the identifying fields sent to the recipient: masterId, versionId, corporateKey.
        created.Should().NotBeNull();
        created!.DocumentId.Should().Be(document.Id);
        created.SourceVersionId.Should().Be(versionId);
        created.CorporateKey.Should().Be("ACME-42");

        _storage.Verify(s => s.UploadAsync(versionId, It.IsAny<byte[]>(), DocxMime, It.IsAny<CancellationToken>()), Times.Once);
        _storage.Verify(s => s.UploadRawAsync(It.Is<string>(n => n.StartsWith("deliveries/")), It.IsAny<byte[]>(), DocxMime, It.IsAny<CancellationToken>()), Times.Once);
        _deliveryRepo.Verify(r => r.AddAsync(It.IsAny<DocumentDelivery>(), It.IsAny<CancellationToken>()), Times.Once);
        _deliveryRepo.Verify(r => r.SaveChangesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Test]
    public async Task Handle_WhenActiveDeliveryExists_ShouldReturnExistingAndNotCreateNew()
    {
        var document = BuildDocumentWithEditableVersion(out var versionId,
            "{\"returnUrl\":\"https://example.com/cb\"}");
        var existing = DocumentDelivery.Create(Guid.NewGuid(), document.Id, versionId,
            "deliveries/x", 10, "h", "https://example.com/cb", "User", Guid.NewGuid(), TimeSpan.FromHours(24));

        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(existing);

        var command = new FinishAndSendDocumentCommand(document.Id, versionId, new byte[] { 1 }, "User");
        var result = await _handler.Handle(command, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DeliveryId.Should().Be(existing.Id);
        _deliveryRepo.Verify(r => r.AddAsync(It.IsAny<DocumentDelivery>(), It.IsAny<CancellationToken>()), Times.Never);
        _storage.Verify(s => s.UploadRawAsync(It.IsAny<string>(), It.IsAny<byte[]>(), It.IsAny<string>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Test]
    public async Task Handle_WithoutReturnUrl_ShouldFail()
    {
        var document = BuildDocumentWithEditableVersion(out var versionId, metadata: null);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);

        var command = new FinishAndSendDocumentCommand(document.Id, versionId, new byte[] { 1 }, "User");
        var result = await _handler.Handle(command, CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("returnUrl");
    }

    [Test]
    public async Task Handle_WhenDocumentMissing_ShouldReturnNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var command = new FinishAndSendDocumentCommand(Guid.NewGuid(), Guid.NewGuid(), new byte[] { 1 }, "User");
        var result = await _handler.Handle(command, CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }
}

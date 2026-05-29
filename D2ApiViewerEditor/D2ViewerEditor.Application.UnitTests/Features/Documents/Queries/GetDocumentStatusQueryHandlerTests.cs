using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentStatus;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentStatusQueryHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IDocumentDeliveryRepository> _deliveryRepo = null!;
    private GetDocumentStatusQueryHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _deliveryRepo = new Mock<IDocumentDeliveryRepository>();
        _handler = new GetDocumentStatusQueryHandler(_documentRepo.Object, _deliveryRepo.Object);
    }

    private static Document BuildDocWithVersion(string? metadata, out Guid versionId)
    {
        var doc = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata);
        var v = doc.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User");
        versionId = v.Id;
        return doc;
    }

    [Test]
    public async Task Handle_DocumentMissing_ReturnsNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_SavedDocumentNoCallbackNoDelivery_ReturnsBaselineDto()
    {
        var doc = BuildDocWithVersion(metadata: null, out var versionId);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        var dto = result.Value!;
        dto.MasterId.Should().Be(doc.Id);
        dto.Status.Should().Be("Saved");
        dto.IsLocked.Should().BeFalse();
        dto.HasCallbackUrl.Should().BeFalse();
        dto.ActiveVersionId.Should().Be(versionId);
        dto.ActiveVersionNumber.Should().Be(1);
        dto.LatestDelivery.Should().BeNull();
    }

    [Test]
    public async Task Handle_EditingDocument_ReportsIsLockedTrue()
    {
        var doc = BuildDocWithVersion(metadata: null, out _);
        doc.MarkEditing();
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.IsLocked.Should().BeTrue();
        result.Value.Status.Should().Be("Editing");
    }

    [Test]
    public async Task Handle_WithCallbackUrl_ReportsHasCallbackUrlTrue_WithoutEchoingTheUrl()
    {
        var doc = BuildDocWithVersion(
            metadata: "{\"returnUrl\":\"https://example.com/cb?token=secret\",\"classification\":\"C2\"}",
            out _);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.HasCallbackUrl.Should().BeTrue();
        // The DTO must not have a CallbackUrl property — guard against a future regression
        // where someone adds it back and leaks the URL (and embedded token).
        typeof(DocumentStatusDto).GetProperty("CallbackUrl").Should().BeNull();
    }

    [Test]
    public async Task Handle_MalformedMetadata_ReportsHasCallbackUrlFalse()
    {
        var doc = BuildDocWithVersion(metadata: "{not-json", out _);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.HasCallbackUrl.Should().BeFalse();
    }

    [TestCase(null, false, TestName = "metadata missing ⇒ userDownload false")]
    [TestCase("{}", false, TestName = "metadata empty ⇒ userDownload false")]
    [TestCase("{\"userDownload\":false}", false, TestName = "explicit false ⇒ false")]
    [TestCase("{\"userDownload\":true}", true, TestName = "explicit true ⇒ true")]
    public async Task Handle_UserDownload_IsMirroredFromMetadata(string? metadata, bool expected)
    {
        var doc = BuildDocWithVersion(metadata, out _);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync((DocumentDelivery?)null);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.UserDownload.Should().Be(expected);
    }

    [Test]
    public async Task Handle_WithActiveDelivery_IncludesDeliveryProjection()
    {
        var doc = BuildDocWithVersion(metadata: null, out var versionId);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var delivery = DocumentDelivery.Create(
            id: Guid.NewGuid(),
            documentId: doc.Id,
            sourceVersionId: versionId,
            snapshotObjectName: "deliveries/x",
            snapshotSizeBytes: 10,
            snapshotSha256: new string('a', 64),
            recipientUrl: "https://example.com/cb",
            createdBy: "User",
            correlationId: Guid.NewGuid(),
            retentionWindow: TimeSpan.FromHours(24));
        _deliveryRepo.Setup(r => r.GetActiveByDocumentIdAsync(doc.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(delivery);

        var result = await _handler.Handle(new GetDocumentStatusQuery(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.LatestDelivery.Should().NotBeNull();
        result.Value.LatestDelivery!.DeliveryId.Should().Be(delivery.Id);
        result.Value.LatestDelivery.Status.Should().Be("Pending");
        result.Value.LatestDelivery.AttemptCount.Should().Be(0);
        result.Value.LatestDelivery.DeadlineAt.Should().Be(delivery.DeadlineAt);
    }
}

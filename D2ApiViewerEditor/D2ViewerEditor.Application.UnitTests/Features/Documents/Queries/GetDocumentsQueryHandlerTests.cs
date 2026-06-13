using D2ViewerEditor.Application.Features.Documents.Queries.GetDocuments;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentsQueryHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private IDocumentRepository _repo = null!;
    private GetDocumentsQueryHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _repo = Substitute.For<IDocumentRepository>();
        _handler = new GetDocumentsQueryHandler(_repo);
    }

    [Test]
    public async Task Handle_MapsDocumentsToDtos_WithActiveVersion()
    {
        var doc = new Document(Guid.NewGuid(), "raport.docx", DocxMime, "User", null);
        doc.AddVersion(Guid.NewGuid(), "documents/v1", 100, "User");
        var v2 = doc.AddVersion(Guid.NewGuid(), "documents/v2", 120, "User"); // aktywna = ostatnia
        _repo.GetAllAsync(0, 200, Arg.Any<CancellationToken>()).Returns(new List<Document> { doc });

        var result = await _handler.Handle(new GetDocumentsQuery(), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        var dto = result.Value!.Single();
        dto.MasterId.Should().Be(doc.Id);
        dto.Name.Should().Be("raport.docx");
        dto.MimeType.Should().Be(DocxMime);
        dto.ActiveVersionId.Should().Be(v2.Id);
        dto.VersionNumber.Should().Be(2);
        dto.Status.Should().Be(doc.Status.ToString());
    }

    [Test]
    public async Task Handle_NoDocuments_ReturnsEmptyList()
    {
        _repo.GetAllAsync(Arg.Any<int>(), Arg.Any<int>(), Arg.Any<CancellationToken>())
            .Returns(new List<Document>());

        var result = await _handler.Handle(new GetDocumentsQuery(), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.Should().BeEmpty();
    }

    [Test]
    public async Task Handle_RepositoryThrows_ReturnsFailure()
    {
        _repo.GetAllAsync(Arg.Any<int>(), Arg.Any<int>(), Arg.Any<CancellationToken>())
            .Returns<IReadOnlyList<Document>>(_ => throw new Exception("db down"));

        var result = await _handler.Handle(new GetDocumentsQuery(), CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("db down");
    }
}

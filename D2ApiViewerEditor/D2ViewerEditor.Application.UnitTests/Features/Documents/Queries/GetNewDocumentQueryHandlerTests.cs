using D2ViewerEditor.Application.Features.Documents.Queries.GetNewDocument;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetNewDocumentQueryHandlerTests
{
    private GetNewDocumentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _handler = new GetNewDocumentQueryHandler();
    }

    [Test]
    public async Task Handle_ShouldReturnSuccessResult()
    {
        // Act
        var result = await _handler.Handle(new GetNewDocumentQuery(), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
    }

    [Test]
    public async Task Handle_ShouldReturnDocumentContentWithHtml()
    {
        // Act
        var result = await _handler.Handle(new GetNewDocumentQuery(), CancellationToken.None);

        // Assert
        result.Value!.Html.Should().NotBeNullOrEmpty();
        result.Value.Html.Should().Contain("document-content");
    }

    [Test]
    public async Task Handle_ShouldReturnDocumentWithMetadata()
    {
        // Act
        var result = await _handler.Handle(new GetNewDocumentQuery(), CancellationToken.None);

        // Assert
        result.Value!.Metadata.Should().NotBeNull();
        result.Value.Metadata.Title.Should().Be("Nowy dokument");
        result.Value.Metadata.Created.Should().BeCloseTo(DateTime.Now, TimeSpan.FromSeconds(5));
        result.Value.Metadata.Modified.Should().BeCloseTo(DateTime.Now, TimeSpan.FromSeconds(5));
    }

    [Test]
    public async Task Handle_ShouldReturnEmptyImagesList()
    {
        // Act
        var result = await _handler.Handle(new GetNewDocumentQuery(), CancellationToken.None);

        // Assert
        result.Value!.Images.Should().NotBeNull().And.BeEmpty();
    }
}

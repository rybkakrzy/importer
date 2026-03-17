using D2ViewerEditor.Application.Features.Documents.Queries.GetTemplates;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetTemplatesQueryHandlerTests
{
    private GetTemplatesQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _handler = new GetTemplatesQueryHandler();
    }

    [Test]
    public async Task Handle_ShouldReturnSuccessResult()
    {
        // Act
        var result = await _handler.Handle(new GetTemplatesQuery(), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
    }

    [Test]
    public async Task Handle_ShouldReturnFourTemplates()
    {
        // Act
        var result = await _handler.Handle(new GetTemplatesQuery(), CancellationToken.None);

        // Assert
        result.Value.Should().HaveCount(4);
    }

    [Test]
    public async Task Handle_ShouldContainBlankTemplate()
    {
        // Act
        var result = await _handler.Handle(new GetTemplatesQuery(), CancellationToken.None);

        // Assert
        result.Value.Should().Contain(t => t.Id == "blank");
    }

    [Test]
    public async Task Handle_ShouldContainAllExpectedIds()
    {
        // Act
        var result = await _handler.Handle(new GetTemplatesQuery(), CancellationToken.None);

        // Assert
        var ids = result.Value!.Select(t => t.Id).ToList();
        ids.Should().Contain("blank");
        ids.Should().Contain("letter");
        ids.Should().Contain("report");
        ids.Should().Contain("cv");
    }

    [Test]
    public async Task Handle_AllTemplates_ShouldHaveNonEmptyNames()
    {
        // Act
        var result = await _handler.Handle(new GetTemplatesQuery(), CancellationToken.None);

        // Assert
        result.Value!.Should().AllSatisfy(t =>
        {
            t.Name.Should().NotBeNullOrEmpty();
            t.Description.Should().NotBeNullOrEmpty();
        });
    }
}

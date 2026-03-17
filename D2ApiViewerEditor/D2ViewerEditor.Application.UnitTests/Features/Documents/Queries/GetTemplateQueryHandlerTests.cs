using D2ViewerEditor.Application.Features.Documents.Queries.GetTemplate;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetTemplateQueryHandlerTests
{
    private GetTemplateQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _handler = new GetTemplateQueryHandler();
    }

    [Test]
    [TestCase("blank")]
    [TestCase("letter")]
    [TestCase("report")]
    [TestCase("cv")]
    public async Task Handle_KnownTemplateId_ShouldReturnSuccess(string templateId)
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery(templateId), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().NotBeNullOrEmpty();
    }

    [Test]
    public async Task Handle_BlankTemplate_ShouldReturnSimpleHtml()
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery("blank"), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Contain("document-content");
    }

    [Test]
    public async Task Handle_LetterTemplate_ShouldContainSenderAndRecipient()
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery("letter"), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Contain("Nadawca");
        result.Value.Html.Should().Contain("Odbiorca");
    }

    [Test]
    public async Task Handle_ReportTemplate_ShouldContainWprowadzenie()
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery("report"), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Contain("Wprowadzenie");
        result.Value.Html.Should().Contain("Wnioski");
    }

    [Test]
    public async Task Handle_CvTemplate_ShouldContainDoswiadczenie()
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery("cv"), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Contain("Doświadczenie zawodowe");
    }

    [Test]
    public async Task Handle_UnknownTemplateId_ShouldReturnDefaultHtml()
    {
        // Act
        var result = await _handler.Handle(new GetTemplateQuery("nonexistent"), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Contain("document-content");
    }
}

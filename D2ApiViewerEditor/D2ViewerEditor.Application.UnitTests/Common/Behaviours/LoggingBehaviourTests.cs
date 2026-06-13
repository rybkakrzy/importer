using D2ViewerEditor.Application.Common.Behaviours;
using FluentAssertions;
using MediatR;
using Microsoft.Extensions.Logging;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common.Behaviours;

[TestFixture]
public class LoggingBehaviourTests
{
    public sealed record SampleRequest(string Name) : IRequest<string>;

    [Test]
    public async Task Handle_ReturnsResponseFromNext()
    {
        var logger = Substitute.For<ILogger<LoggingBehaviour<SampleRequest, string>>>();
        var behaviour = new LoggingBehaviour<SampleRequest, string>(logger);

        var result = await behaviour.Handle(
            new SampleRequest("x"),
            () => Task.FromResult("handled"),
            CancellationToken.None);

        result.Should().Be("handled");
    }

    [Test]
    public async Task Handle_LogsBeforeAndAfterHandling()
    {
        var logger = Substitute.For<ILogger<LoggingBehaviour<SampleRequest, string>>>();
        var behaviour = new LoggingBehaviour<SampleRequest, string>(logger);

        await behaviour.Handle(
            new SampleRequest("x"),
            () => Task.FromResult("handled"),
            CancellationToken.None);

        // One log on entry ("Handling") and one on exit ("Handled").
        logger.ReceivedCalls()
            .Count(c => c.GetMethodInfo().Name == nameof(ILogger.Log))
            .Should().Be(2);
    }

    [Test]
    public async Task Handle_DoesNotSwallowExceptionFromNext()
    {
        var logger = Substitute.For<ILogger<LoggingBehaviour<SampleRequest, string>>>();
        var behaviour = new LoggingBehaviour<SampleRequest, string>(logger);

        var act = async () => await behaviour.Handle(
            new SampleRequest("x"),
            () => throw new InvalidOperationException("boom"),
            CancellationToken.None);

        await act.Should().ThrowAsync<InvalidOperationException>().WithMessage("boom");
    }
}

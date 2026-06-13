using D2ViewerEditor.Application.Common.Behaviours;
using FluentAssertions;
using FluentValidation;
using FluentValidation.Results;
using MediatR;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common.Behaviours;

[TestFixture]
public class ValidationBehaviourTests
{
    public sealed record SampleRequest(string Name) : IRequest<string>;

    private static RequestHandlerDelegate<string> NextReturning(string value, Action? onCalled = null) =>
        () =>
        {
            onCalled?.Invoke();
            return Task.FromResult(value);
        };

    [Test]
    public async Task Handle_NoValidators_CallsNextAndReturnsResponse()
    {
        var behaviour = new ValidationBehaviour<SampleRequest, string>(
            Enumerable.Empty<IValidator<SampleRequest>>());
        var nextCalled = false;

        var result = await behaviour.Handle(
            new SampleRequest("ok"),
            NextReturning("handled", () => nextCalled = true),
            CancellationToken.None);

        result.Should().Be("handled");
        nextCalled.Should().BeTrue();
    }

    [Test]
    public async Task Handle_AllValidatorsPass_CallsNext()
    {
        var validator = Substitute.For<IValidator<SampleRequest>>();
        validator
            .ValidateAsync(Arg.Any<ValidationContext<SampleRequest>>(), Arg.Any<CancellationToken>())
            .Returns(new ValidationResult());

        var behaviour = new ValidationBehaviour<SampleRequest, string>(new[] { validator });
        var nextCalled = false;

        var result = await behaviour.Handle(
            new SampleRequest("ok"),
            NextReturning("handled", () => nextCalled = true),
            CancellationToken.None);

        result.Should().Be("handled");
        nextCalled.Should().BeTrue();
    }

    [Test]
    public async Task Handle_ValidatorFails_ThrowsValidationExceptionAndDoesNotCallNext()
    {
        var validator = Substitute.For<IValidator<SampleRequest>>();
        var failure = new ValidationFailure("Name", "Name is required");
        validator
            .ValidateAsync(Arg.Any<ValidationContext<SampleRequest>>(), Arg.Any<CancellationToken>())
            .Returns(new ValidationResult(new[] { failure }));

        var behaviour = new ValidationBehaviour<SampleRequest, string>(new[] { validator });
        var nextCalled = false;

        var act = async () => await behaviour.Handle(
            new SampleRequest(""),
            NextReturning("handled", () => nextCalled = true),
            CancellationToken.None);

        (await act.Should().ThrowAsync<ValidationException>())
            .Which.Errors.Should().ContainSingle(e => e.ErrorMessage == "Name is required");
        nextCalled.Should().BeFalse("the pipeline must short-circuit on validation failure");
    }

    [Test]
    public async Task Handle_AggregatesFailuresAcrossMultipleValidators()
    {
        var first = Substitute.For<IValidator<SampleRequest>>();
        first.ValidateAsync(Arg.Any<ValidationContext<SampleRequest>>(), Arg.Any<CancellationToken>())
            .Returns(new ValidationResult(new[] { new ValidationFailure("Name", "error A") }));

        var second = Substitute.For<IValidator<SampleRequest>>();
        second.ValidateAsync(Arg.Any<ValidationContext<SampleRequest>>(), Arg.Any<CancellationToken>())
            .Returns(new ValidationResult(new[] { new ValidationFailure("Name", "error B") }));

        var behaviour = new ValidationBehaviour<SampleRequest, string>(new[] { first, second });

        var act = async () => await behaviour.Handle(
            new SampleRequest(""),
            NextReturning("handled"),
            CancellationToken.None);

        (await act.Should().ThrowAsync<ValidationException>())
            .Which.Errors.Select(e => e.ErrorMessage)
            .Should().BeEquivalentTo(new[] { "error A", "error B" });
    }
}

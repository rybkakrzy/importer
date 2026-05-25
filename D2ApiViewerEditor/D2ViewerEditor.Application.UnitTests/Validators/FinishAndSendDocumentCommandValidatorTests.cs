using D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;
using D2ViewerEditor.Application.Validators.Documents;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Validators;

[TestFixture]
public class FinishAndSendDocumentCommandValidatorTests
{
    private FinishAndSendDocumentCommandValidator _validator = null!;

    [SetUp]
    public void SetUp() => _validator = new FinishAndSendDocumentCommandValidator();

    [Test]
    public void Validate_ValidCommand_ShouldPass()
    {
        var command = new FinishAndSendDocumentCommand(Guid.NewGuid(), Guid.NewGuid(), new byte[] { 1, 2 }, "User");

        _validator.Validate(command).IsValid.Should().BeTrue();
    }

    [Test]
    public void Validate_EmptyMasterId_ShouldFail()
    {
        var command = new FinishAndSendDocumentCommand(Guid.Empty, Guid.NewGuid(), new byte[] { 1 }, "User");

        _validator.Validate(command).IsValid.Should().BeFalse();
    }

    [Test]
    public void Validate_EmptyContent_ShouldFail()
    {
        var command = new FinishAndSendDocumentCommand(Guid.NewGuid(), Guid.NewGuid(), Array.Empty<byte>(), "User");

        _validator.Validate(command).IsValid.Should().BeFalse();
    }
}

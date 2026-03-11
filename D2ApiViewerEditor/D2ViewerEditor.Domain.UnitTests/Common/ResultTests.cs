using D2ViewerEditor.Domain.Common;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Common;

[TestFixture]
public class ResultTests
{
    [Test]
    public void Success_ShouldCreateSuccessResult()
    {
        // Arrange
        var value = "test value";

        // Act
        var result = Result<string>.Success(value);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.IsFailure.Should().BeFalse();
        result.Value.Should().Be(value);
        result.Error.Should().BeNull();
    }

    [Test]
    public void Failure_ShouldCreateFailureResult()
    {
        // Arrange
        var errorMessage = "Something went wrong";

        // Act
        var result = Result<string>.Failure(errorMessage);

        // Assert
        result.IsSuccess.Should().BeFalse();
        result.IsFailure.Should().BeTrue();
        result.Value.Should().BeNull();
        result.Error.Should().Be(errorMessage);
    }

    [Test]
    public void Match_WhenSuccess_ShouldExecuteOnSuccessCallback()
    {
        // Arrange
        var result = Result<int>.Success(42);
        var expectedOutput = "Number: 42";

        // Act
        var output = result.Match(
            onSuccess: value => $"Number: {value}",
            onFailure: error => $"Error: {error}"
        );

        // Assert
        output.Should().Be(expectedOutput);
    }

    [Test]
    public void Match_WhenFailure_ShouldExecuteOnFailureCallback()
    {
        // Arrange
        var errorMessage = "Invalid operation";
        var result = Result<int>.Failure(errorMessage);
        var expectedOutput = $"Error: {errorMessage}";

        // Act
        var output = result.Match(
            onSuccess: value => $"Number: {value}",
            onFailure: error => $"Error: {error}"
        );

        // Assert
        output.Should().Be(expectedOutput);
    }

    [Test]
    public void ResultWithoutValue_Success_ShouldCreateSuccessResult()
    {
        // Act
        var result = Result.Success();

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.IsFailure.Should().BeFalse();
        result.Error.Should().BeNull();
    }

    [Test]
    public void ResultWithoutValue_Failure_ShouldCreateFailureResult()
    {
        // Arrange
        var errorMessage = "Operation failed";

        // Act
        var result = Result.Failure(errorMessage);

        // Assert
        result.IsSuccess.Should().BeFalse();
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Be(errorMessage);
    }
}

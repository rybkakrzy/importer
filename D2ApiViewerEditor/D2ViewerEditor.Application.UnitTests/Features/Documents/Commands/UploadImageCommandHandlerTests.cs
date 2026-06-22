using D2ViewerEditor.Application.Features.Documents.Commands.UploadImage;
using D2ViewerEditor.Application.Common.Security;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class UploadImageCommandHandlerTests
{
    private IFileUploadSecurityService _uploadSecurityService;
    private UploadImageCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _uploadSecurityService = Substitute.For<IFileUploadSecurityService>();
        _uploadSecurityService.ValidateImageAsync(
                Arg.Any<byte[]>(),
                Arg.Any<string>(),
                Arg.Any<string>(),
                Arg.Any<CancellationToken>())
            .Returns(ci =>
            {
                var fileName = ci.ArgAt<string>(1);
                var contentType = ci.ArgAt<string>(2);
                var extension = Path.GetExtension(fileName).ToLowerInvariant();
                return extension is ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp" or ".webp"
                    ? UploadValidationResult.Success(contentType)
                    : UploadValidationResult.Failure(UploadRejectionCode.UnsupportedExtension, "Niedozwolony format obrazu.");
            });
        _handler = new UploadImageCommandHandler(_uploadSecurityService);
    }

    [Test]
    public async Task Handle_ValidJpgFile_ShouldReturnBase64Response()
    {
        // Arrange
        var fileContent = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }; // JPEG header
        var stream = new MemoryStream(fileContent);

        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: "photo.jpg",
            ContentType: "image/jpeg",
            FileSize: fileContent.Length
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Base64.Should().StartWith("data:image/jpeg;base64,");
        result.Value.FileName.Should().Be("photo.jpg");
        result.Value.Size.Should().Be(fileContent.Length);
    }

    [Test]
    public async Task Handle_ValidPngFile_ShouldReturnBase64Response()
    {
        // Arrange
        var fileContent = new byte[] { 0x89, 0x50, 0x4E, 0x47 }; // PNG header
        var stream = new MemoryStream(fileContent);

        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: "image.png",
            ContentType: "image/png",
            FileSize: fileContent.Length
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Base64.Should().StartWith("data:image/png;base64,");
    }

    [Test]
    [TestCase(".gif")]
    [TestCase(".bmp")]
    [TestCase(".webp")]
    [TestCase(".jpeg")]
    public async Task Handle_AllowedExtension_ShouldReturnSuccess(string extension)
    {
        // Arrange
        var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: $"file{extension}",
            ContentType: "image/jpeg",
            FileSize: 3
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
    }

    [Test]
    [TestCase(".pdf")]
    [TestCase(".docx")]
    [TestCase(".txt")]
    [TestCase(".exe")]
    public async Task Handle_DisallowedExtension_ShouldReturnFailure(string extension)
    {
        // Arrange
        var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: $"file{extension}",
            ContentType: "application/pdf",
            FileSize: 3
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Upload odrzucony");
    }

    [Test]
    public async Task Handle_ValidFile_ShouldContainCorrectBase64Data()
    {
        // Arrange
        var fileContent = new byte[] { 1, 2, 3, 4, 5 };
        var expectedBase64 = Convert.ToBase64String(fileContent);
        var stream = new MemoryStream(fileContent);

        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: "test.png",
            ContentType: "image/png",
            FileSize: fileContent.Length
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Base64.Should().Be($"data:image/png;base64,{expectedBase64}");
    }

    [Test]
    public async Task Handle_UpperCaseExtension_ShouldReturnSuccess()
    {
        // Arrange
        var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var command = new UploadImageCommand(
            FileStream: stream,
            FileName: "IMAGE.JPG",
            ContentType: "image/jpeg",
            FileSize: 3
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
    }
}

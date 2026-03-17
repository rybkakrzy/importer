using D2ViewerEditor.Application.Features.Documents.Commands.SignDocument;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NSubstitute;
using NSubstitute.ExceptionExtensions;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class SignDocumentCommandHandlerTests
{
    private IHtmlToDocxConverter _converter;
    private IDigitalSignatureService _signatureService;
    private SignDocumentCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _converter = Substitute.For<IHtmlToDocxConverter>();
        _signatureService = Substitute.For<IDigitalSignatureService>();
        _handler = new SignDocumentCommandHandler(_converter, _signatureService);
    }

    [Test]
    public async Task Handle_ValidCommand_ShouldReturnSignedDocxBytes()
    {
        // Arrange
        var signedBytes = new byte[] { 10, 20, 30 };
        var certBytes = new byte[] { 5, 6 };
        var certBase64 = Convert.ToBase64String(certBytes);

        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1, 2, 3 });

        _signatureService.SignDocument(
            Arg.Any<byte[]>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<string>(),
            Arg.Any<string?>(), Arg.Any<string?>(), Arg.Any<string?>())
            .Returns(signedBytes);

        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "signed.docx",
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: certBase64,
            CertificatePassword: "password",
            SignerName: "Jan Kowalski",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(signedBytes);
        result.Value.FileName.Should().Be("signed.docx");
    }

    [Test]
    public async Task Handle_WithoutOriginalFileName_ShouldUseDefaultName()
    {
        // Arrange
        var certBytes = new byte[] { 1 };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });
        _signatureService.SignDocument(Arg.Any<byte[]>(), Arg.Any<byte[]>(), Arg.Any<string>(),
            Arg.Any<string>(), Arg.Any<string?>(), Arg.Any<string?>(), Arg.Any<string?>())
            .Returns(new byte[] { 2 });

        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: Convert.ToBase64String(certBytes),
            CertificatePassword: "pass",
            SignerName: "Signer",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("dokument.docx");
    }

    [Test]
    public async Task Handle_FileNameWithoutDocxExtension_ShouldAppendExtension()
    {
        // Arrange
        var certBytes = new byte[] { 1 };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });
        _signatureService.SignDocument(Arg.Any<byte[]>(), Arg.Any<byte[]>(), Arg.Any<string>(),
            Arg.Any<string>(), Arg.Any<string?>(), Arg.Any<string?>(), Arg.Any<string?>())
            .Returns(new byte[] { 2 });

        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "plik",
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: Convert.ToBase64String(certBytes),
            CertificatePassword: "pass",
            SignerName: "Signer",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("plik.docx");
    }

    [Test]
    public async Task Handle_ConverterThrowsException_ShouldReturnFailure()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Throws(new InvalidOperationException("Błąd konwersji"));

        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: Convert.ToBase64String(new byte[] { 1 }),
            CertificatePassword: "pass",
            SignerName: "Signer",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Błąd konwersji");
    }

    [Test]
    public async Task Handle_SignatureServiceThrowsException_ShouldReturnFailure()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });
        _signatureService.SignDocument(Arg.Any<byte[]>(), Arg.Any<byte[]>(), Arg.Any<string>(),
            Arg.Any<string>(), Arg.Any<string?>(), Arg.Any<string?>(), Arg.Any<string?>())
            .Throws(new InvalidOperationException("Błąd podpisu"));

        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: Convert.ToBase64String(new byte[] { 1 }),
            CertificatePassword: "pass",
            SignerName: "Signer",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Błąd podpisu");
    }

    [Test]
    public async Task Handle_InvalidBase64Certificate_ShouldReturnFailure()
    {
        // Arrange
        var command = new SignDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null,
            CertificateBase64: "not-valid-base64!!!",
            CertificatePassword: "pass",
            SignerName: "Signer",
            SignerTitle: null,
            SignerEmail: null,
            SignatureReason: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().StartWith("Błąd podpisywania dokumentu:");
    }
}

using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Kontrakt serwisu podpisów cyfrowych
/// </summary>
public interface IDigitalSignatureService
{
    /// <summary>
    /// Podpisuje dokument DOCX certyfikatem X.509
    /// </summary>
    byte[] SignDocument(byte[] docxBytes, byte[] certificateBytes, string password,
                        string signerName, string? signerTitle = null,
                        string? signerEmail = null, string? reason = null);

    /// <summary>
    /// Weryfikuje podpisy cyfrowe w dokumencie DOCX
    /// </summary>
    List<DigitalSignatureInfo> VerifySignatures(Stream docxStream);
}

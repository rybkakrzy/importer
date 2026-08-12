using System.IO;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OpenMcdf;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Normalizacja wejścia: pass-through DOCX (ZIP), detekcja zaszyfrowanego DOCX (hasło) oraz
/// binarnego .doc — rozpoznanie po zawartości (magic-bytes + zawartość kontenera CFB), nie po rozszerzeniu.
/// </summary>
[TestFixture]
public class DocumentInputNormalizerTests
{
    private DocumentInputNormalizer _normalizer = null!;

    // Hasło-fikstura losowane per uruchomienie testów — w źródłach nie ma literału hasła
    // (SAST: „hard-coded password"); round-trip szyfrowania i tak używa tej samej wartości.
    private static readonly string TestPassword = "test-" + Guid.NewGuid().ToString("N");
    private static readonly string WrongPassword = TestPassword + "-zle";

    [SetUp]
    public void SetUp() => _normalizer = new DocumentInputNormalizer();

    [Test]
    public void Zip_IsTreatedAsDocx_PassThrough()
    {
        // Nagłówek ZIP (PK\x03\x04) = zwykły DOCX.
        var bytes = new byte[] { 0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0, 1, 2 };

        var result = _normalizer.Normalize(bytes);

        result.Status.Should().Be(DocumentInputStatus.Ok);
        result.Docx.Should().BeEquivalentTo(bytes);
    }

    [Test]
    public void Garbage_IsInvalid()
    {
        _normalizer.Normalize(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 })
            .Status.Should().Be(DocumentInputStatus.Invalid);
    }

    [Test]
    public void EncryptedDocx_WithoutPassword_RequiresPassword()
    {
        // Realny zaszyfrowany DOCX (EncryptionInfo + EncryptedPackage); bez hasła → trzeba poprosić.
        var encrypted = EncryptDocx(BuildMinimalDocx(), TestPassword);

        _normalizer.Normalize(encrypted, password: null)
            .Status.Should().Be(DocumentInputStatus.PasswordRequired);
    }

    [Test]
    public void BinaryDoc_IsUnsupportedLegacyDoc()
    {
        // CFB z wpisem WordDocument (a bez EncryptionInfo) = binarny .doc.
        var cfb = BuildCfb("WordDocument", new byte[] { 0xEC, 0xA5, 0xC1, 0x00 });

        _normalizer.Normalize(cfb)
            .Status.Should().Be(DocumentInputStatus.UnsupportedLegacyDoc);
    }

    [Test]
    public void EncryptedDocx_WithCorrectPassword_IsDecryptedToValidDocx()
    {
        var docx = BuildMinimalDocx();
        var encrypted = EncryptDocx(docx, TestPassword);

        var result = _normalizer.Normalize(encrypted, TestPassword);

        result.Status.Should().Be(DocumentInputStatus.Ok, because: "poprawne hasło musi odszyfrować dokument");
        result.Docx.Should().NotBeNull();
        // Odszyfrowany pakiet to zwykły DOCX (ZIP) i daje się otworzyć.
        result.Docx![0].Should().Be(0x50); // 'P'
        result.Docx![1].Should().Be(0x4B); // 'K'
        using var ms = new MemoryStream(result.Docx!);
        var act = () => { using var d = WordprocessingDocument.Open(ms, false); _ = d.MainDocumentPart!.Document; };
        act.Should().NotThrow();
    }

    [Test]
    public void EncryptedDocx_WithWrongPassword_ReturnsWrongPassword()
    {
        var encrypted = EncryptDocx(BuildMinimalDocx(), TestPassword);

        _normalizer.Normalize(encrypted, WrongPassword)
            .Status.Should().Be(DocumentInputStatus.WrongPassword);
    }

    private static byte[] BuildMinimalDocx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Tajne")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    // Spec-zgodny enkryptor Agile (MS-OFFCRYPTO) — odwrotność produkcyjnego OoxmlAgileDecryptor.
    // Sami wytwarzamy zaszyfrowaną fixturę (kontener CFB składamy OpenMcdf). Te same stałe
    // blockKey i algorytm (SHA512 + AES-256-CBC) co dekryptor → round-trip dowodzi inwersji, a zgodność
    // ze spec sprawia, że realne pliki Office też się odszyfrują.
    private static readonly byte[] BVerInput = { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
    private static readonly byte[] BVerValue = { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
    private static readonly byte[] BKeyValue = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

    private static byte[] EncryptDocx(byte[] docx, string password)
    {
        const int spin = 1000, keyBytes = 32, block = 16;
        var rnd = System.Security.Cryptography.RandomNumberGenerator.Create();
        byte[] keSalt = new byte[16], pkgSalt = new byte[16], secretKey = new byte[keyBytes], verifierInput = new byte[16];
        rnd.GetBytes(keSalt); rnd.GetBytes(pkgSalt); rnd.GetBytes(secretKey); rnd.GetBytes(verifierInput);

        byte[] H(byte[] d) { using var s = System.Security.Cryptography.SHA512.Create(); return s.ComputeHash(d); }
        byte[] Cat(byte[] a, byte[] b) { var r = new byte[a.Length + b.Length]; Array.Copy(a, r, a.Length); Array.Copy(b, 0, r, a.Length, b.Length); return r; }
        byte[] Derive(byte[] pw, byte[] blk) { var h = H(Cat(pw, blk)); return h.AsSpan(0, keyBytes).ToArray(); }
        byte[] Enc(byte[] data, byte[] key, byte[] iv)
        {
            using var aes = System.Security.Cryptography.Aes.Create();
            aes.Mode = System.Security.Cryptography.CipherMode.CBC; aes.Padding = System.Security.Cryptography.PaddingMode.None;
            aes.Key = key; aes.IV = iv.AsSpan(0, 16).ToArray();
            var pad = data; if (pad.Length % 16 != 0) { pad = new byte[data.Length + (16 - data.Length % 16)]; Array.Copy(data, pad, data.Length); }
            using var e = aes.CreateEncryptor(); return e.TransformFinalBlock(pad, 0, pad.Length);
        }

        var pwHash = H(Cat(keSalt, System.Text.Encoding.Unicode.GetBytes(password)));
        for (int i = 0; i < spin; i++) pwHash = H(Cat(BitConverter.GetBytes(i), pwHash));

        var encVerInput = Enc(verifierInput, Derive(pwHash, BVerInput), keSalt);
        var encVerValue = Enc(H(verifierInput), Derive(pwHash, BVerValue), keSalt);
        var encKeyValue = Enc(secretKey, Derive(pwHash, BKeyValue), keSalt);

        // EncryptedPackage: 8B rozmiar + segmenty 4096B (IV = H(pkgSalt + i)).
        using var pkg = new MemoryStream();
        pkg.Write(BitConverter.GetBytes((long)docx.Length));
        for (int off = 0, bi = 0; off < docx.Length; off += 4096, bi++)
        {
            int len = Math.Min(4096, docx.Length - off);
            var iv = H(Cat(pkgSalt, BitConverter.GetBytes(bi))).AsSpan(0, block).ToArray();
            pkg.Write(Enc(docx.AsSpan(off, len).ToArray(), secretKey, iv));
        }

        string B64(byte[] b) => Convert.ToBase64String(b);
        var xml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" " +
            "xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" +
            $"<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"{B64(pkgSalt)}\"/>" +
            "<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" +
            $"<p:encryptedKey spinCount=\"{spin}\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" " +
            $"saltValue=\"{B64(keSalt)}\" encryptedVerifierHashInput=\"{B64(encVerInput)}\" encryptedVerifierHashValue=\"{B64(encVerValue)}\" encryptedKeyValue=\"{B64(encKeyValue)}\"/>" +
            "</keyEncryptor></keyEncryptors></encryption>";
        var xmlBytes = System.Text.Encoding.UTF8.GetBytes(xml);
        using var infoMs = new MemoryStream();
        infoMs.Write(new byte[] { 0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00 }); // major 4, minor 4, flags 0x40
        infoMs.Write(xmlBytes);

        return BuildCfb(("EncryptionInfo", infoMs.ToArray()), ("EncryptedPackage", pkg.ToArray()));
    }

    private static byte[] BuildCfb(string entryName, byte[] data) => BuildCfb((entryName, data));

    private static byte[] BuildCfb(params (string Name, byte[] Data)[] entries)
    {
        using var ms = new MemoryStream();
        using (var root = RootStorage.Create(ms, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            foreach (var (name, data) in entries)
            {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(data, 0, data.Length);
            }
        }
        return ms.ToArray();
    }
}

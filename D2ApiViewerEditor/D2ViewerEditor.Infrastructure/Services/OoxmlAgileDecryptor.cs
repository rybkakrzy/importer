using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Dekrypcja pakietu OOXML zabezpieczonego hasłem — schemat „Agile Encryption" (MS-OFFCRYPTO §2.3.4.10+),
/// domyślny w Office 2010+. Czysto zarządzane (System.Security.Cryptography), bez zależności natywnych.
/// NPOI 2.8.0 dostarcza tylko interfejs dekryptora, ale NIE implementacji Agile (brak builderów) — stąd
/// własna implementacja. Standard/CryptoAPI (Office 2007) nie jest tu obsługiwany (zwraca Unsupported).
/// </summary>
public static class OoxmlAgileDecryptor
{
    public enum DecryptResult { Ok, WrongPassword, Unsupported, Invalid }

    // Stałe blockKey z MS-OFFCRYPTO (§2.3.4.12–14).
    private static readonly byte[] BlockVerifierHashInput = { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
    private static readonly byte[] BlockVerifierHashValue = { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
    private static readonly byte[] BlockKeyValue = { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

    /// <param name="encryptionInfo">Strumień EncryptionInfo z kontenera CFB.</param>
    /// <param name="encryptedPackage">Strumień EncryptedPackage z kontenera CFB.</param>
    /// <param name="docx">Odszyfrowany pakiet OOXML (ZIP) — gdy wynik == Ok.</param>
    public static DecryptResult TryDecrypt(byte[] encryptionInfo, byte[] encryptedPackage, string password, out byte[]? docx)
    {
        docx = null;
        if (encryptionInfo.Length < 8) return DecryptResult.Invalid;

        // Nagłówek: major(2)+minor(2)+flags(4). Agile = 4.4; reszta (np. 3.2/4.2) = Standard/CryptoAPI.
        var major = BitConverter.ToUInt16(encryptionInfo, 0);
        var minor = BitConverter.ToUInt16(encryptionInfo, 2);
        if (major != 4 || minor != 4)
            return DecryptResult.Unsupported;

        XElement enc;
        try
        {
            var xml = Encoding.UTF8.GetString(encryptionInfo, 8, encryptionInfo.Length - 8);
            enc = XElement.Parse(xml);
        }
        catch
        {
            return DecryptResult.Invalid;
        }

        XNamespace e = "http://schemas.microsoft.com/office/2006/encryption";
        XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        var keyData = enc.Element(e + "keyData");
        var encryptedKey = enc.Element(e + "keyEncryptors")?.Element(e + "keyEncryptor")?.Element(p + "encryptedKey");
        if (keyData == null || encryptedKey == null) return DecryptResult.Invalid;

        try
        {
            // --- keyData (do odszyfrowania pakietu) ---
            var pkgSalt = B64(keyData, "saltValue");
            var pkgBlockSize = Int(keyData, "blockSize");
            var pkgKeyBits = Int(keyData, "keyBits");
            var pkgHash = HashName(Str(keyData, "hashAlgorithm"));

            // --- encryptedKey (do wyprowadzenia klucza z hasła) ---
            var keSalt = B64(encryptedKey, "saltValue");
            var spinCount = Int(encryptedKey, "spinCount");
            var keKeyBits = Int(encryptedKey, "keyBits");
            var keHash = HashName(Str(encryptedKey, "hashAlgorithm"));
            var encVerifierInput = B64(encryptedKey, "encryptedVerifierHashInput");
            var encVerifierValue = B64(encryptedKey, "encryptedVerifierHashValue");
            var encKeyValue = B64(encryptedKey, "encryptedKeyValue");
            if (pkgHash == null || keHash == null) return DecryptResult.Unsupported;

            // Hasło → klucz bazowy: H(salt + UTF16LE(pwd)), potem spinCount iteracji H(i + H).
            var pwHash = Hash(keHash.Value, Concat(keSalt, Encoding.Unicode.GetBytes(password)));
            for (int i = 0; i < spinCount; i++)
                pwHash = Hash(keHash.Value, Concat(BitConverter.GetBytes(i), pwHash));

            int keKeyBytes = keKeyBits / 8;
            var verifierInputKey = DeriveKey(keHash.Value, pwHash, BlockVerifierHashInput, keKeyBytes);
            var verifierValueKey = DeriveKey(keHash.Value, pwHash, BlockVerifierHashValue, keKeyBytes);
            var keyValueKey = DeriveKey(keHash.Value, pwHash, BlockKeyValue, keKeyBytes);

            // Weryfikacja hasła: H(odszyfrowany verifierInput) == odszyfrowany verifierValue.
            var verifierInput = AesCbcDecrypt(encVerifierInput, verifierInputKey, keSalt);
            var verifierValue = AesCbcDecrypt(encVerifierValue, verifierValueKey, keSalt);
            var verifierInputHash = Hash(keHash.Value, verifierInput);
            if (!verifierInputHash.AsSpan(0, Math.Min(verifierInputHash.Length, verifierValue.Length))
                    .SequenceEqual(verifierValue.AsSpan(0, Math.Min(verifierInputHash.Length, verifierValue.Length))))
                return DecryptResult.WrongPassword;

            // Sekretny klucz pakietu.
            var secretKey = AesCbcDecrypt(encKeyValue, keyValueKey, keSalt);
            if (secretKey.Length > pkgKeyBits / 8)
                secretKey = secretKey.AsSpan(0, pkgKeyBits / 8).ToArray();

            // EncryptedPackage: 8 bajtów rozmiar (LE) + segmenty 4096 B (każdy własne IV = H(pkgSalt + i)).
            if (encryptedPackage.Length < 8) return DecryptResult.Invalid;
            long totalSize = BitConverter.ToInt64(encryptedPackage, 0);
            var output = new MemoryStream();
            const int segment = 4096;
            int offset = 8;
            int blockIndex = 0;
            while (offset < encryptedPackage.Length)
            {
                int len = Math.Min(segment, encryptedPackage.Length - offset);
                if (len % pkgBlockSize != 0) len -= len % pkgBlockSize; // segmenty są wielokrotnością bloku
                if (len <= 0) break;
                var iv = Hash(pkgHash.Value, Concat(pkgSalt, BitConverter.GetBytes(blockIndex)));
                if (iv.Length > pkgBlockSize) iv = iv.AsSpan(0, pkgBlockSize).ToArray();
                var chunk = encryptedPackage.AsSpan(offset, len).ToArray();
                output.Write(AesCbcDecrypt(chunk, secretKey, iv), 0, len);
                offset += segment;
                blockIndex++;
            }

            var all = output.ToArray();
            if (totalSize > 0 && totalSize <= all.Length)
                all = all.AsSpan(0, (int)totalSize).ToArray();
            docx = all;
            return DecryptResult.Ok;
        }
        catch
        {
            return DecryptResult.Invalid;
        }
    }

    /// <summary>deriveKey: H(pwHash + blockKey) przycięty/uzupełniony do keyBytes (MS-OFFCRYPTO §2.3.4.11).</summary>
    private static byte[] DeriveKey(HashAlgorithmName hash, byte[] pwHash, byte[] blockKey, int keyBytes)
    {
        var h = Hash(hash, Concat(pwHash, blockKey));
        if (h.Length >= keyBytes) return h.AsSpan(0, keyBytes).ToArray();
        var key = new byte[keyBytes];
        Array.Fill(key, (byte)0x36);
        Array.Copy(h, key, h.Length);
        return key;
    }

    private static byte[] AesCbcDecrypt(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        aes.Key = key;
        aes.IV = iv.Length == 16 ? iv : iv.AsSpan(0, 16).ToArray();
        using var dec = aes.CreateDecryptor();
        var padded = data;
        if (padded.Length % 16 != 0)
        {
            padded = new byte[data.Length + (16 - data.Length % 16)];
            Array.Copy(data, padded, data.Length);
        }
        return dec.TransformFinalBlock(padded, 0, padded.Length);
    }

    private static byte[] Hash(HashAlgorithmName name, byte[] data)
    {
        using var algo = IncrementalHash.CreateHash(name);
        algo.AppendData(data);
        return algo.GetHashAndReset();
    }

    private static HashAlgorithmName? HashName(string? a) => a?.ToUpperInvariant() switch
    {
        "SHA512" => HashAlgorithmName.SHA512,
        "SHA384" => HashAlgorithmName.SHA384,
        "SHA256" => HashAlgorithmName.SHA256,
        "SHA-512" => HashAlgorithmName.SHA512,
        "SHA-256" => HashAlgorithmName.SHA256,
        "SHA1" => HashAlgorithmName.SHA1,
        _ => null
    };

    private static byte[] Concat(byte[] a, byte[] b)
    {
        var r = new byte[a.Length + b.Length];
        Array.Copy(a, r, a.Length);
        Array.Copy(b, 0, r, a.Length, b.Length);
        return r;
    }

    private static string? Str(XElement el, string attr) => el.Attribute(attr)?.Value;
    private static int Int(XElement el, string attr) => int.Parse(el.Attribute(attr)!.Value);
    private static byte[] B64(XElement el, string attr) => Convert.FromBase64String(el.Attribute(attr)!.Value);
}

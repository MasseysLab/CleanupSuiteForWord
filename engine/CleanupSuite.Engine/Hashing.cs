using System.Security.Cryptography;

namespace MasseysLab.CleanupSuite.Engine;

public static class Hashing
{
    public static string Sha256Hex(ReadOnlySpan<byte> bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    public static string Sha256Utf8(string text) =>
        Sha256Hex(JsonSupport.StrictUtf8.GetBytes(text));

    public static bool IsLowercaseSha256(string value) =>
        value.Length == 64
        && value.All(character =>
            character is >= '0' and <= '9'
            || character is >= 'a' and <= 'f');
}

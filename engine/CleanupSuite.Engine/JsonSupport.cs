using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MasseysLab.CleanupSuite.Engine;

public static class JsonSupport
{
    public static readonly UTF8Encoding StrictUtf8 = new(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);

    public static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNameCaseInsensitive = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    public static readonly JsonSerializerOptions WriteOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static T DeserializeStrict<T>(byte[] bytes)
    {
        RejectUtf8Bom(bytes);
        string json = StrictUtf8.GetString(bytes);
        return JsonSerializer.Deserialize<T>(json, ReadOptions)
            ?? throw new InvalidDataException("JSON deserialized to null.");
    }

    public static byte[] SerializeUtf8<T>(T value) =>
        JsonSerializer.SerializeToUtf8Bytes(value, WriteOptions);

    public static void RejectUtf8Bom(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length >= 3
            && bytes[0] == 0xEF
            && bytes[1] == 0xBB
            && bytes[2] == 0xBF)
        {
            throw new InvalidDataException("UTF-8 byte-order marks are not allowed.");
        }
    }
}

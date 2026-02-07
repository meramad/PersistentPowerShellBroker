using System.Text;
using System.Text.Json;

namespace PersistentPowerShellBroker.Protocol;

public static class JsonLineCodec
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = null
    };

    public static async Task<T> ReadLineAsync<T>(Stream stream, CancellationToken cancellationToken)
    {
        using var reader = new StreamReader(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true);
        var line = await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
        if (line is null)
        {
            throw new EndOfStreamException("Expected one JSON line request but stream ended.");
        }

        var value = JsonSerializer.Deserialize<T>(line, SerializerOptions);
        if (value is null)
        {
            throw new JsonException("JSON payload deserialized to null.");
        }

        return value;
    }

    public static async Task WriteLineAsync<T>(Stream stream, T value, CancellationToken cancellationToken)
    {
        var json = JsonSerializer.Serialize(value, SerializerOptions);
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true);
        await writer.WriteAsync(json.AsMemory(), cancellationToken).ConfigureAwait(false);
        await writer.WriteLineAsync().ConfigureAwait(false);
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
    }
}

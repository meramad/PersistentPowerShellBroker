using System.Text.Json.Serialization;

namespace PersistentPowerShellBroker.Protocol;

public sealed class BrokerResponse
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("success")]
    public bool Success { get; init; }

    [JsonPropertyName("stdout")]
    public string Stdout { get; init; } = string.Empty;

    [JsonPropertyName("stderr")]
    public string Stderr { get; init; } = string.Empty;

    [JsonPropertyName("error")]
    public string? Error { get; init; }

    [JsonPropertyName("durationMs")]
    public int DurationMs { get; init; }
}

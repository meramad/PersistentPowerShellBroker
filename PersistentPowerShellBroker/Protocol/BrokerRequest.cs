using System.Text.Json;
using System.Text.Json.Serialization;

namespace PersistentPowerShellBroker.Protocol;

public sealed class BrokerRequest
{
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    [JsonPropertyName("kind")]
    public string Kind { get; init; } = string.Empty;

    [JsonPropertyName("command")]
    public string Command { get; init; } = string.Empty;

    [JsonPropertyName("args")]
    public JsonElement? Args { get; init; }

    [JsonPropertyName("timeoutMs")]
    public int? TimeoutMs { get; init; }

    [JsonPropertyName("clientName")]
    public string? ClientName { get; init; }

    [JsonPropertyName("clientPid")]
    public int? ClientPid { get; init; }
}

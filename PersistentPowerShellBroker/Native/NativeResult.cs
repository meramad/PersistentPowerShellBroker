namespace PersistentPowerShellBroker.Native;

public sealed class NativeResult
{
    public bool Success { get; init; }
    public string Stdout { get; init; } = string.Empty;
    public string Stderr { get; init; } = string.Empty;
    public string? Error { get; init; }
}

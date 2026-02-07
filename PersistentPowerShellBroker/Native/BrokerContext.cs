namespace PersistentPowerShellBroker.Native;

public sealed class BrokerContext
{
    public required string PipeName { get; init; }
    public required DateTimeOffset StartedAtUtc { get; init; }
    public required int ProcessId { get; init; }
    public required Action RequestStop { get; init; }
}

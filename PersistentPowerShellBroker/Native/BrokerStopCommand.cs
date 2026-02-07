namespace PersistentPowerShellBroker.Native;

public sealed class BrokerStopCommand : INativeCommand
{
    public string Name => "broker.stop";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, CancellationToken cancellationToken)
    {
        context.RequestStop();
        return Task.FromResult(new NativeResult
        {
            Success = true,
            Stdout = "stopping",
            Stderr = string.Empty,
            Error = null
        });
    }
}

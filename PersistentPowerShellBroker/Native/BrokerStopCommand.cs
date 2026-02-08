using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerStopCommand : INativeCommand
{
    public string Name => "broker.stop";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, Runspace runspace, CancellationToken cancellationToken)
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

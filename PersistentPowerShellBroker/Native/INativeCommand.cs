using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public interface INativeCommand
{
    string Name { get; }
    Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, Runspace runspace, CancellationToken cancellationToken);
}

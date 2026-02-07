using System.Reflection;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerInfoCommand : INativeCommand
{
    public string Name => "broker.info";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, CancellationToken cancellationToken)
    {
        var version = Assembly.GetEntryAssembly()?.GetName().Version?.ToString() ?? "unknown";
        var payload = new
        {
            version,
            pipeName = context.PipeName,
            startedAtUtc = context.StartedAtUtc,
            pid = context.ProcessId
        };

        return Task.FromResult(new NativeResult
        {
            Success = true,
            Stdout = JsonSerializer.Serialize(payload),
            Stderr = string.Empty,
            Error = null
        });
    }
}

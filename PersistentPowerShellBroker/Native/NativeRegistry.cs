namespace PersistentPowerShellBroker.Native;

public sealed class NativeRegistry
{
    private readonly Dictionary<string, INativeCommand> _commands;

    public NativeRegistry(IEnumerable<INativeCommand> commands)
    {
        _commands = commands.ToDictionary(command => command.Name, StringComparer.OrdinalIgnoreCase);
    }

    public bool TryGet(string name, out INativeCommand command) => _commands.TryGetValue(name, out command!);
}

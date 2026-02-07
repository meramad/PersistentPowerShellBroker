namespace PersistentPowerShellBroker.Util;

public sealed class ConsoleLogger
{
    private readonly object _gate = new();
    private readonly LogLevel _minimumLevel;

    public ConsoleLogger(LogLevel minimumLevel)
    {
        _minimumLevel = minimumLevel;
    }

    public void Info(string message) => Write(LogLevel.Info, message);

    public void Debug(string message) => Write(LogLevel.Debug, message);

    public void Error(string message) => Write(LogLevel.Info, $"ERROR {message}");

    private void Write(LogLevel level, string message)
    {
        if (level > _minimumLevel)
        {
            return;
        }

        lock (_gate)
        {
            Console.WriteLine(message);
        }
    }
}

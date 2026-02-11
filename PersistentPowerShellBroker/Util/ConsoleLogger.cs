namespace PersistentPowerShellBroker.Util;

public sealed class ConsoleLogger
{
    private readonly object _gate = new();
    private readonly LogLevel _minimumLevel;
    public LogLevel MinimumLevel => _minimumLevel;

    public ConsoleLogger(LogLevel minimumLevel)
    {
        _minimumLevel = minimumLevel;
    }

    public void Info(string message) => Write(LogLevel.Info, message);

    public void Debug(string message) => Write(LogLevel.Debug, message);

    public void Error(string message)
    {
        if (_minimumLevel == LogLevel.Silent)
        {
            return;
        }

        lock (_gate)
        {
            Console.WriteLine($"ERROR {message}");
        }
    }

    public void PrettyBlock(IEnumerable<string> lines)
    {
        if (_minimumLevel != LogLevel.Pretty)
        {
            return;
        }

        lock (_gate)
        {
            foreach (var line in lines)
            {
                Console.WriteLine(line);
            }
        }
    }

    private void Write(LogLevel level, string message)
    {
        if (!IsEnabled(level))
        {
            return;
        }

        lock (_gate)
        {
            Console.WriteLine(message);
        }
    }

    private bool IsEnabled(LogLevel level)
    {
        return _minimumLevel switch
        {
            LogLevel.Silent => false,
            LogLevel.Pretty => level == LogLevel.Pretty,
            LogLevel.Info => level == LogLevel.Info,
            LogLevel.Debug => level == LogLevel.Info || level == LogLevel.Debug,
            _ => false
        };
    }
}

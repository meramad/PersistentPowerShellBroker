using PersistentPowerShellBroker;
using PersistentPowerShellBroker.Native;
using PersistentPowerShellBroker.Util;

return await ProgramEntry.RunAsync(args);

internal static class ProgramEntry
{
    private const int ExitCodeOk = 0;
    private const int ExitCodeInvalidArgs = 2;
    private const int ExitCodeStartupFatal = 3;
    private const int ExitCodePipeFatal = 4;

    public static async Task<int> RunAsync(string[] args)
    {
        if (!TryParseArgs(args, out var options, out var error))
        {
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.Error.WriteLine(error);
            }

            return ExitCodeInvalidArgs;
        }

        using var stopSignal = new StopSignal();
        using var ctrlC = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            stopSignal.RequestStop();
            ctrlC.Cancel();
        };

        var logger = new ConsoleLogger(options.LogLevel);
        var context = new BrokerContext
        {
            PipeName = options.PipeName,
            StartedAtUtc = DateTimeOffset.UtcNow,
            ProcessId = Environment.ProcessId,
            RequestStop = stopSignal.RequestStop
        };

        var registry = new NativeRegistry(
        [
            new BrokerInfoCommand(),
            new BrokerStopCommand(),
            new BrokerHelpCommand(),
            new BrokerExcelGetWorkbookHandleCommand(),
            new BrokerExcelReleaseHandleCommand()
        ]);
        using var host = new BrokerHost(registry, context, options.InitScriptPath);
        try
        {
            await host.StartAsync(ctrlC.Token).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Startup error: {ex.Message}");
            return ExitCodeStartupFatal;
        }

        Console.WriteLine($"PersistentPowerShellBroker v{AppVersion.Value}");
        Console.WriteLine("\u00A9 2026 Mikl\u00F3s Ar\u00E1nyi");
        Console.WriteLine("Experimental local automation tool");
        Console.WriteLine($"PIPE=\\\\.\\pipe\\{options.PipeName}");
        TimeSpan? idleExit = options.IdleExitMinutes.HasValue ? TimeSpan.FromMinutes(options.IdleExitMinutes.Value) : null;
        var server = new PipeServer(options.PipeName, host, logger, stopSignal, idleExit);

        try
        {
            await server.RunAsync(ctrlC.Token).ConfigureAwait(false);
            return ExitCodeOk;
        }
        catch (Exception ex)
        {
            logger.Error($"Pipe server fatal error: {ex.Message}");
            return ExitCodePipeFatal;
        }
    }

    private static bool TryParseArgs(string[] args, out ProgramOptions options, out string? error)
    {
        error = null;
        options = new ProgramOptions();
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (!arg.StartsWith("--", StringComparison.Ordinal))
            {
                error = $"Unknown argument '{arg}'.";
                return false;
            }

            if (i + 1 >= args.Length)
            {
                error = $"Missing value for '{arg}'.";
                return false;
            }

            var value = args[++i];
            switch (arg)
            {
                case "--pipe":
                    options.PipeName = string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)
                        ? $"psbroker-{Guid.NewGuid():N}"
                        : value;
                    break;
                case "--log-option":
                    if (string.Equals(value, "silent", StringComparison.OrdinalIgnoreCase))
                    {
                        options.LogLevel = LogLevel.Silent;
                    }
                    else if (string.Equals(value, "info", StringComparison.OrdinalIgnoreCase))
                    {
                        options.LogLevel = LogLevel.Info;
                    }
                    else if (string.Equals(value, "debug", StringComparison.OrdinalIgnoreCase))
                    {
                        options.LogLevel = LogLevel.Debug;
                    }
                    else
                    {
                        error = "Invalid --log-option. Use silent, info, or debug.";
                        return false;
                    }

                    break;
                case "--init":
                    options.InitScriptPath = value;
                    break;
                case "--idle-exit-minutes":
                    if (!int.TryParse(value, out var minutes) || minutes < 1)
                    {
                        error = "Invalid --idle-exit-minutes. Use an integer >= 1.";
                        return false;
                    }

                    options.IdleExitMinutes = minutes;
                    break;
                default:
                    error = $"Unknown option '{arg}'.";
                    return false;
            }
        }

        if (string.IsNullOrWhiteSpace(options.PipeName))
        {
            error = "Missing required option --pipe auto|<name>.";
            return false;
        }

        return true;
    }

    private sealed class ProgramOptions
    {
        public string PipeName { get; set; } = string.Empty;
        public LogLevel LogLevel { get; set; } = LogLevel.Silent;
        public string? InitScriptPath { get; set; }
        public int? IdleExitMinutes { get; set; }
    }
}

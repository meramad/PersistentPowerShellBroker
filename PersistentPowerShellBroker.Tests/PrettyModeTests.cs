using System.Reflection;
using PersistentPowerShellBroker.Protocol;
using PersistentPowerShellBroker.Util;

namespace PersistentPowerShellBroker.Tests;

public sealed class PrettyModeTests
{
    [Fact]
    public void PrettyFormatter_PrintsPlainCommandAndStdout()
    {
        var request = new BrokerRequest
        {
            Id = "1",
            Kind = "powershell",
            Command = "Get-Date"
        };
        var response = new BrokerResponse
        {
            Id = "1",
            Success = true,
            Stdout = "hello world",
            Stderr = string.Empty,
            Error = null,
            DurationMs = 1
        };

        var lines = PrettyLogFormatter.Format(request, response);

        Assert.Contains("> Get-Date", lines);
        Assert.Contains("hello world", lines);
        Assert.DoesNotContain(lines, line => line.Contains("durationMs", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PrettyFormatter_PrettyPrintsJsonAndRedactsSensitiveFields()
    {
        var request = new BrokerRequest
        {
            Id = "2",
            Kind = "native",
            Command = "broker.help"
        };
        var response = new BrokerResponse
        {
            Id = "2",
            Success = true,
            Stdout = "{\"ok\":true,\"token\":\"abc\",\"nested\":{\"password\":\"x\"}}",
            Stderr = string.Empty,
            Error = null,
            DurationMs = 5
        };

        var lines = PrettyLogFormatter.Format(request, response);
        var text = string.Join(Environment.NewLine, lines);

        Assert.Contains("{", text);
        Assert.Contains("\"token\": \"***\"", text);
        Assert.Contains("\"password\": \"***\"", text);
        Assert.DoesNotContain("\"abc\"", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"x\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ProgramParse_DefaultLogOptionIsPretty_AndAcceptsExplicitPretty()
    {
        Assert.Equal(LogLevel.Pretty, ParseLogLevel(["--pipe", "psbroker-test"]));
        Assert.Equal(LogLevel.Pretty, ParseLogLevel(["--pipe", "psbroker-test", "--log-option", "pretty"]));
    }

    private static LogLevel ParseLogLevel(string[] args)
    {
        var tryParseArgs = typeof(ProgramEntry).GetMethod("TryParseArgs", BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(tryParseArgs);

        var parameters = new object?[] { args, null, null };
        var ok = (bool)tryParseArgs!.Invoke(null, parameters)!;
        Assert.True(ok);

        var options = parameters[1];
        Assert.NotNull(options);
        var logLevelProp = options!.GetType().GetProperty("LogLevel", BindingFlags.Public | BindingFlags.Instance);
        Assert.NotNull(logLevelProp);
        return (LogLevel)logLevelProp!.GetValue(options)!;
    }
}

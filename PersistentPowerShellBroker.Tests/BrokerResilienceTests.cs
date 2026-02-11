using System.Diagnostics;
using System.Management.Automation.Runspaces;
using PersistentPowerShellBroker.Native;
using PersistentPowerShellBroker.Protocol;

namespace PersistentPowerShellBroker.Tests;

public sealed class BrokerResilienceTests
{
    [Fact]
    public void TryInvokeWithTimeout_TimesOutBlockingAction()
    {
        var started = Stopwatch.StartNew();
        var completed = ExcelCommandSupport.TryInvokeWithTimeout(
            () =>
            {
                Thread.Sleep(Timeout.Infinite);
                return 1;
            },
            TimeSpan.FromMilliseconds(250),
            out int? result,
            out var error);
        started.Stop();

        Assert.False(completed);
        Assert.Null(error);
        Assert.Null(result);
        Assert.True(started.Elapsed < TimeSpan.FromSeconds(2));
    }

    [Fact]
    public void TryInvokeWithTimeout_CapturesException()
    {
        var completed = ExcelCommandSupport.TryInvokeWithTimeout(
            () =>
            {
                throw new InvalidOperationException("boom");
            },
            TimeSpan.FromSeconds(1),
            out object? result,
            out var error);

        Assert.True(completed);
        Assert.Null(result);
        Assert.NotNull(error);
        Assert.IsType<InvalidOperationException>(error);
    }

    [Fact]
    public async Task Broker_RemainsResponsive_AfterTimedNativeTimeout()
    {
        var context = new BrokerContext
        {
            PipeName = "test-pipe",
            StartedAtUtc = DateTimeOffset.UtcNow,
            ProcessId = Environment.ProcessId,
            RequestStop = static () => { }
        };
        var registry = new NativeRegistry(
        [
            new BrokerInfoCommand(),
            new TimeoutSimulationCommand()
        ]);

        using var host = new BrokerHost(registry, context, initScriptPath: null);
        await host.StartAsync(CancellationToken.None);

        var timeoutRequest = new BrokerRequest
        {
            Id = "r1",
            Kind = "native",
            Command = "broker.test.timeout"
        };

        var timeoutResponse = await host.ExecuteAsync(timeoutRequest, CancellationToken.None);
        Assert.False(timeoutResponse.Success);
        Assert.Equal("CommandTimeout", timeoutResponse.Error);

        var infoRequest = new BrokerRequest
        {
            Id = "r2",
            Kind = "native",
            Command = "broker.info"
        };

        var infoResponse = await host.ExecuteAsync(infoRequest, CancellationToken.None);
        Assert.True(infoResponse.Success);
        Assert.Contains("\"pipeName\":\"test-pipe\"", infoResponse.Stdout, StringComparison.Ordinal);
    }

    private sealed class TimeoutSimulationCommand : INativeCommand
    {
        public string Name => "broker.test.timeout";

        public Task<NativeResult> ExecuteAsync(
            System.Text.Json.JsonElement? args,
            BrokerContext context,
            Runspace runspace,
            CancellationToken cancellationToken)
        {
            var completed = ExcelCommandSupport.TryInvokeWithTimeout(
                () =>
                {
                    Thread.Sleep(Timeout.Infinite);
                    return 1;
                },
                TimeSpan.FromMilliseconds(200),
                out int? _,
                out Exception? _);

            return Task.FromResult(new NativeResult
            {
                Success = false,
                Stdout = "{\"ok\":false,\"status\":\"CommandTimeout\"}",
                Stderr = string.Empty,
                Error = completed ? "UnexpectedCompletion" : "CommandTimeout"
            });
        }
    }
}

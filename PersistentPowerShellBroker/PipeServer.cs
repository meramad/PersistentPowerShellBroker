using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using PersistentPowerShellBroker.Protocol;
using PersistentPowerShellBroker.Util;

namespace PersistentPowerShellBroker;

public sealed class PipeServer
{
    private readonly string _pipeName;
    private readonly BrokerHost _brokerHost;
    private readonly ConsoleLogger _logger;
    private readonly StopSignal _stopSignal;
    private readonly TimeSpan? _idleExit;
    private long _lastActivityTick;

    public PipeServer(
        string pipeName,
        BrokerHost brokerHost,
        ConsoleLogger logger,
        StopSignal stopSignal,
        TimeSpan? idleExit)
    {
        _pipeName = pipeName;
        _brokerHost = brokerHost;
        _logger = logger;
        _stopSignal = stopSignal;
        _idleExit = idleExit;
        TouchActivity();
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _stopSignal.Token);
        var token = linkedCts.Token;
        using var idleTimerCts = CancellationTokenSource.CreateLinkedTokenSource(token);
        var idleTimerTask = MonitorIdleAsync(idleTimerCts.Token);
        var handlers = new HashSet<Task>();

        try
        {
            while (!token.IsCancellationRequested)
            {
                var server = CreatePipeServerStream();
                try
                {
                    await server.WaitForConnectionAsync(token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    server.Dispose();
                    break;
                }

                var task = HandleConnectionAsync(server, token);
                lock (handlers)
                {
                    handlers.Add(task);
                }

                _ = task.ContinueWith(
                    completed =>
                    {
                        lock (handlers)
                        {
                            handlers.Remove(completed);
                        }
                    },
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
        }
        finally
        {
            idleTimerCts.Cancel();
            await AwaitQuietly(idleTimerTask).ConfigureAwait(false);

            Task[] outstanding;
            lock (handlers)
            {
                outstanding = handlers.ToArray();
            }

            await Task.WhenAll(outstanding).ConfigureAwait(false);
        }
    }

    private async Task HandleConnectionAsync(NamedPipeServerStream pipeStream, CancellationToken cancellationToken)
    {
        using var _ = pipeStream;
        TouchActivity();
        BrokerRequest? request = null;
        try
        {
            request = await JsonLineCodec.ReadLineAsync<BrokerRequest>(pipeStream, cancellationToken).ConfigureAwait(false);
            var validationError = ValidateRequest(request);
            if (validationError is not null)
            {
                await JsonLineCodec.WriteLineAsync(pipeStream, validationError, cancellationToken).ConfigureAwait(false);
                return;
            }

            if (request.TimeoutMs.HasValue)
            {
                _logger.Debug($"request={request.Id} timeoutMs ignored in v1 ({request.TimeoutMs.Value})");
            }

            _logger.Debug($"request={request.Id} kind={request.Kind} command={request.Command}");
            var response = await _brokerHost.ExecuteAsync(request, cancellationToken).ConfigureAwait(false);
            await JsonLineCodec.WriteLineAsync(pipeStream, response, cancellationToken).ConfigureAwait(false);

            var preview = request.Command.Length <= 80 ? request.Command : $"{request.Command[..80]}...";
            _logger.Info($"request={response.Id} kind={request.Kind} success={response.Success} durationMs={response.DurationMs} command=\"{preview}\"");
            if (!response.Success || !string.IsNullOrWhiteSpace(response.Error))
            {
                _logger.Debug($"request={response.Id} stderr={response.Stderr}");
                _logger.Debug($"request={response.Id} error={response.Error}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            var id = request?.Id ?? "unknown";
            _logger.Error($"request={id} failed: {ex.Message}");
            var response = new BrokerResponse
            {
                Id = id,
                Success = false,
                Stdout = string.Empty,
                Stderr = string.Empty,
                Error = ex.Message,
                DurationMs = 0
            };

            try
            {
                await JsonLineCodec.WriteLineAsync(pipeStream, response, cancellationToken).ConfigureAwait(false);
            }
            catch
            {
            }
        }
        finally
        {
            TouchActivity();
        }
    }

    private BrokerResponse? ValidateRequest(BrokerRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Id))
        {
            return Invalid("unknown", "Missing required field 'id'.");
        }

        if (string.IsNullOrWhiteSpace(request.Kind))
        {
            return Invalid(request.Id, "Missing required field 'kind'.");
        }

        if (string.IsNullOrWhiteSpace(request.Command))
        {
            return Invalid(request.Id, "Missing required field 'command'.");
        }

        if (!string.Equals(request.Kind, "powershell", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(request.Kind, "native", StringComparison.OrdinalIgnoreCase))
        {
            return Invalid(request.Id, "Field 'kind' must be 'powershell' or 'native'.");
        }

        return null;
    }

    private static BrokerResponse Invalid(string id, string error) => new()
    {
        Id = id,
        Success = false,
        Stdout = string.Empty,
        Stderr = string.Empty,
        Error = error,
        DurationMs = 0
    };

    private NamedPipeServerStream CreatePipeServerStream()
    {
        return new NamedPipeServerStream(
            _pipeName,
            PipeDirection.InOut,
            NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 4096,
            outBufferSize: 4096,
            CreatePipeSecurity());
    }

    private static PipeSecurity CreatePipeSecurity()
    {
        var currentUser = WindowsIdentity.GetCurrent().User
            ?? throw new InvalidOperationException("Could not resolve current Windows user.");

        var pipeSecurity = new PipeSecurity();
        pipeSecurity.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            currentUser,
            PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance,
            AccessControlType.Allow));
        pipeSecurity.SetOwner(currentUser);

        return pipeSecurity;
    }

    private async Task MonitorIdleAsync(CancellationToken cancellationToken)
    {
        if (_idleExit is null)
        {
            return;
        }

        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(5));
        while (await timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
        {
            var idleTicks = Environment.TickCount64 - Interlocked.Read(ref _lastActivityTick);
            if (idleTicks >= _idleExit.Value.TotalMilliseconds)
            {
                _logger.Info("idle timeout reached, stopping broker");
                _stopSignal.RequestStop();
                return;
            }
        }
    }

    private void TouchActivity()
    {
        Interlocked.Exchange(ref _lastActivityTick, Environment.TickCount64);
    }

    private static async Task AwaitQuietly(Task task)
    {
        try
        {
            await task.ConfigureAwait(false);
        }
        catch
        {
        }
    }
}

using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using PersistentPowerShellBroker.Native;
using PersistentPowerShellBroker.Protocol;

namespace PersistentPowerShellBroker;

public sealed class BrokerHost : IDisposable
{
    private readonly BlockingCollection<WorkItem> _queue = new();
    private readonly Thread _workerThread;
    private readonly TaskCompletionSource<bool> _startupTcs = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly NativeRegistry _nativeRegistry;
    private readonly BrokerContext _context;
    private readonly string? _initScriptPath;

    private Runspace? _runspace;
    private volatile bool _disposed;

    public BrokerHost(NativeRegistry nativeRegistry, BrokerContext context, string? initScriptPath)
    {
        _nativeRegistry = nativeRegistry;
        _context = context;
        _initScriptPath = initScriptPath;
        _workerThread = new Thread(WorkerLoop)
        {
            Name = "BrokerHost.STAWorker",
            IsBackground = true
        };
        _workerThread.SetApartmentState(ApartmentState.STA);
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _workerThread.Start();
        using var registration = cancellationToken.Register(() => _startupTcs.TrySetCanceled(cancellationToken));
        await _startupTcs.Task.ConfigureAwait(false);
    }

    public Task<BrokerResponse> ExecuteAsync(BrokerRequest request, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var tcs = new TaskCompletionSource<BrokerResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var item = new WorkItem(request, tcs, cancellationToken);
        _queue.Add(item, cancellationToken);
        return tcs.Task;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _queue.CompleteAdding();
        _workerThread.Join();
        _queue.Dispose();
    }

    private void WorkerLoop()
    {
        try
        {
            var initialState = InitialSessionState.CreateDefault();
            initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;
            _runspace = RunspaceFactory.CreateRunspace(initialState);
            _runspace.Open();

            if (!string.IsNullOrWhiteSpace(_initScriptPath))
            {
                var escapedPath = _initScriptPath.Replace("'", "''", StringComparison.Ordinal);
                var initRequest = new BrokerRequest
                {
                    Id = "init",
                    Kind = "powershell",
                    Command = $". '{escapedPath}'"
                };

                var initResponse = ExecuteInternal(initRequest, CancellationToken.None);
                if (!initResponse.Success)
                {
                    throw new InvalidOperationException($"Init script failed: {initResponse.Stderr} {initResponse.Error}".Trim());
                }
            }

            _startupTcs.TrySetResult(true);
        }
        catch (Exception ex)
        {
            _startupTcs.TrySetException(ex);
            return;
        }

        foreach (var item in _queue.GetConsumingEnumerable())
        {
            if (item.CancellationToken.IsCancellationRequested)
            {
                item.Completion.TrySetCanceled(item.CancellationToken);
                continue;
            }

            var started = Environment.TickCount64;
            try
            {
                var response = ExecuteInternal(item.Request, item.CancellationToken);
                var duration = (int)Math.Max(0, Environment.TickCount64 - started);
                item.Completion.TrySetResult(new BrokerResponse
                {
                    Id = response.Id,
                    Success = response.Success,
                    Stdout = response.Stdout,
                    Stderr = response.Stderr,
                    Error = response.Error,
                    DurationMs = duration
                });
            }
            catch (OperationCanceledException oce)
            {
                item.Completion.TrySetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                var duration = (int)Math.Max(0, Environment.TickCount64 - started);
                item.Completion.TrySetResult(new BrokerResponse
                {
                    Id = item.Request.Id,
                    Success = false,
                    Stdout = string.Empty,
                    Stderr = string.Empty,
                    Error = ex.Message,
                    DurationMs = duration
                });
            }
        }

        _runspace?.Dispose();
    }

    private BrokerResponse ExecuteInternal(BrokerRequest request, CancellationToken cancellationToken)
    {
        if (string.Equals(request.Kind, "powershell", StringComparison.OrdinalIgnoreCase))
        {
            return ExecutePowerShell(request);
        }

        if (string.Equals(request.Kind, "native", StringComparison.OrdinalIgnoreCase))
        {
            return ExecuteNative(request, cancellationToken);
        }

        return new BrokerResponse
        {
            Id = request.Id,
            Success = false,
            Stdout = string.Empty,
            Stderr = string.Empty,
            Error = $"Unsupported request kind '{request.Kind}'.",
            DurationMs = 0
        };
    }

    private BrokerResponse ExecuteNative(BrokerRequest request, CancellationToken cancellationToken)
    {
        if (_runspace is null)
        {
            throw new InvalidOperationException("Runspace is not initialized.");
        }

        if (!_nativeRegistry.TryGet(request.Command, out var command))
        {
            return new BrokerResponse
            {
                Id = request.Id,
                Success = false,
                Stdout = string.Empty,
                Stderr = string.Empty,
                Error = $"Unknown native command '{request.Command}'.",
                DurationMs = 0
            };
        }

        var result = command.ExecuteAsync(request.Args, _context, _runspace, cancellationToken).GetAwaiter().GetResult();
        return new BrokerResponse
        {
            Id = request.Id,
            Success = result.Success,
            Stdout = result.Stdout,
            Stderr = result.Stderr,
            Error = result.Error,
            DurationMs = 0
        };
    }

    private BrokerResponse ExecutePowerShell(BrokerRequest request)
    {
        if (_runspace is null)
        {
            throw new InvalidOperationException("Runspace is not initialized.");
        }

        using var ps = PowerShell.Create();
        ps.Runspace = _runspace;
        ps.AddScript(request.Command, useLocalScope: false);
        ps.AddCommand("Out-String").AddParameter("Width", 4096);

        Collection<PSObject> output = [];
        string exceptionText = string.Empty;
        try
        {
            output = ps.Invoke();
        }
        catch (RuntimeException ex)
        {
            exceptionText = ex.ToString();
        }

        var stdout = string.Concat(output.Select(value => value?.ToString()));
        var streamErrors = ps.Streams.Error.Select(error => error.ToString()).ToArray();
        var stderr = string.Join(Environment.NewLine, streamErrors);
        if (!string.IsNullOrWhiteSpace(exceptionText))
        {
            stderr = string.IsNullOrWhiteSpace(stderr) ? exceptionText : $"{stderr}{Environment.NewLine}{exceptionText}";
        }

        return new BrokerResponse
        {
            Id = request.Id,
            Success = string.IsNullOrWhiteSpace(stderr),
            Stdout = stdout,
            Stderr = stderr,
            Error = null,
            DurationMs = 0
        };
    }

    private sealed record WorkItem(BrokerRequest Request, TaskCompletionSource<BrokerResponse> Completion, CancellationToken CancellationToken);
}

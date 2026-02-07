namespace PersistentPowerShellBroker.Util;

public sealed class StopSignal : IDisposable
{
    private readonly CancellationTokenSource _cts = new();

    public CancellationToken Token => _cts.Token;

    public void RequestStop()
    {
        if (!_cts.IsCancellationRequested)
        {
            _cts.Cancel();
        }
    }

    public void Dispose()
    {
        _cts.Dispose();
    }
}

using System.Runtime.InteropServices;

namespace PersistentPowerShellBroker.Interop;

internal static class OleMessageFilter
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMsgWaitDefProcess = 2;

    public static IDisposable Register(int retryDelayMilliseconds = 100, int maxRetryWindowMilliseconds = 15000)
    {
        var newFilter = new MessageFilter(retryDelayMilliseconds, maxRetryWindowMilliseconds);
        var hr = CoRegisterMessageFilter(newFilter, out var oldFilter);
        Marshal.ThrowExceptionForHR(hr);
        return new Scope(oldFilter);
    }

    private sealed class Scope(IOleMessageFilter? oldFilter) : IDisposable
    {
        private IOleMessageFilter? _oldFilter = oldFilter;
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            var previous = _oldFilter;
            _oldFilter = null;
            var hr = CoRegisterMessageFilter(previous, out _);
            Marshal.ThrowExceptionForHR(hr);
        }
    }

    private sealed class MessageFilter(int retryDelayMilliseconds, int maxRetryWindowMilliseconds) : IOleMessageFilter
    {
        private readonly int _retryDelayMilliseconds = Math.Max(0, retryDelayMilliseconds);
        private readonly int _maxRetryWindowMilliseconds = Math.Max(0, maxRetryWindowMilliseconds);

        public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo)
        {
            return ServerCallIsHandled;
        }

        public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType)
        {
            if (rejectType != ServerCallRetryLater)
            {
                return -1;
            }

            if (tickCount >= _maxRetryWindowMilliseconds)
            {
                return -1;
            }

            return _retryDelayMilliseconds;
        }

        public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType)
        {
            return PendingMsgWaitDefProcess;
        }
    }

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

        [PreserveSig]
        int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
    }
}

using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PersistentPowerShellBroker.Native;

internal sealed class ExcelWorkbookOpener
{
    private const int RpcCallRejected = unchecked((int)0x80010001);

    public object? OpenWorkbookWithRetry(
        object application,
        string target,
        bool readOnly,
        string? openPassword,
        string? modifyPassword,
        TimeSpan timeout)
    {
        var stopwatch = Stopwatch.StartNew();
        var attempt = 0;
        while (true)
        {
            try
            {
                return OpenWorkbook(application, target, readOnly, openPassword, modifyPassword);
            }
            catch (COMException ex) when (ex.HResult == RpcCallRejected)
            {
                attempt++;
                var delayMs = Math.Min(1000, 50 * attempt);
                var delay = TimeSpan.FromMilliseconds(delayMs);
                if (stopwatch.Elapsed + delay >= timeout)
                {
                    throw new ComRetryTimeoutException("Excel remained busy and rejected COM calls until timeout.", ex);
                }

                Thread.Sleep(delay);
            }
        }
    }

    private static object? OpenWorkbook(
        object application,
        string target,
        bool readOnly,
        string? openPassword,
        string? modifyPassword)
    {
        dynamic app = application;
        dynamic workbooks = app.Workbooks;
        var missing = Type.Missing;
        Exception? fullSignatureError = null;

        try
        {
            if (readOnly || !string.IsNullOrWhiteSpace(openPassword) || !string.IsNullOrWhiteSpace(modifyPassword))
            {
                return workbooks.Open(
                    target,
                    missing,
                    readOnly,
                    missing,
                    openPassword ?? missing,
                    modifyPassword ?? missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing);
            }
        }
        catch (Exception ex)
        {
            fullSignatureError = ex;
        }

        try
        {
            return workbooks.Open(target);
        }
        catch (Exception minimalSignatureError)
        {
            if (fullSignatureError is null)
            {
                throw;
            }

            throw new InvalidOperationException(
                $"Excel Workbooks.Open failed with both signatures. full='{fullSignatureError.Message}', minimal='{minimalSignatureError.Message}'.",
                minimalSignatureError);
        }
    }

    internal sealed class ComRetryTimeoutException : TimeoutException
    {
        public ComRetryTimeoutException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}

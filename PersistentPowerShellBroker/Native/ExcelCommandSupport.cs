using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

internal static class ExcelCommandSupport
{
    private const string SuccessStatus = "Success";

    public static NativeResult BuildResult(object payload, bool success)
    {
        return new NativeResult
        {
            Success = success,
            Stdout = JsonSerializer.Serialize(payload),
            Stderr = string.Empty,
            Error = null
        };
    }

    public static NativeResult BuildErrorResult(string status, string? errorCode, string? errorMessage, object payload)
    {
        return new NativeResult
        {
            Success = string.Equals(status, SuccessStatus, StringComparison.Ordinal),
            Stdout = JsonSerializer.Serialize(payload),
            Stderr = string.Empty,
            Error = errorMessage ?? errorCode
        };
    }

    public static bool TryGetString(JsonElement? args, string name, out string? value)
    {
        value = null;
        if (!TryGetProperty(args, name, out var property))
        {
            return true;
        }

        if (property.ValueKind == JsonValueKind.Null)
        {
            value = null;
            return true;
        }

        if (property.ValueKind != JsonValueKind.String)
        {
            return false;
        }

        value = property.GetString();
        return true;
    }

    public static bool TryGetInt(JsonElement? args, string name, out int? value)
    {
        value = null;
        if (!TryGetProperty(args, name, out var property))
        {
            return true;
        }

        if (property.ValueKind == JsonValueKind.Null)
        {
            value = null;
            return true;
        }

        if (!property.TryGetInt32(out var parsed))
        {
            return false;
        }

        value = parsed;
        return true;
    }

    public static bool TryGetBool(JsonElement? args, string name, out bool? value)
    {
        value = null;
        if (!TryGetProperty(args, name, out var property))
        {
            return true;
        }

        if (property.ValueKind == JsonValueKind.Null)
        {
            value = null;
            return true;
        }

        if (property.ValueKind != JsonValueKind.True && property.ValueKind != JsonValueKind.False)
        {
            return false;
        }

        value = property.GetBoolean();
        return true;
    }

    public static bool TryGetVariable(Runspace runspace, string variableName, out object? value)
    {
        value = runspace.SessionStateProxy.GetVariable($"global:{variableName}");
        if (value is not null)
        {
            return true;
        }

        value = runspace.SessionStateProxy.GetVariable(variableName);
        return value is not null;
    }

    public static void RemoveGlobalVariable(Runspace runspace, string variableName)
    {
        using var ps = PowerShell.Create();
        ps.Runspace = runspace;
        ps.AddCommand("Remove-Variable")
            .AddParameter("Name", variableName)
            .AddParameter("Scope", "Global")
            .AddParameter("ErrorAction", "SilentlyContinue");
        ps.Invoke();
    }

    public static string NewHandleVariableName() => $"excelHandle_{Guid.NewGuid():N}";

    public static object BuildBundle(
        object application,
        object workbook,
        string requestedTarget,
        string workbookFullName,
        bool isReadOnly,
        bool attachedExisting,
        bool openedWorkbook,
        string instancePolicyUsed)
    {
        var metadata = new PSObject();
        metadata.Properties.Add(new PSNoteProperty("RequestedTarget", requestedTarget));
        metadata.Properties.Add(new PSNoteProperty("WorkbookFullName", workbookFullName));
        metadata.Properties.Add(new PSNoteProperty("IsReadOnly", isReadOnly));
        metadata.Properties.Add(new PSNoteProperty("AttachedExisting", attachedExisting));
        metadata.Properties.Add(new PSNoteProperty("OpenedWorkbook", openedWorkbook));
        metadata.Properties.Add(new PSNoteProperty("CreatedUtc", DateTime.UtcNow));
        metadata.Properties.Add(new PSNoteProperty("InstancePolicyUsed", instancePolicyUsed));

        var bundle = new PSObject();
        bundle.Properties.Add(new PSNoteProperty("Application", application));
        bundle.Properties.Add(new PSNoteProperty("Workbook", workbook));
        bundle.Properties.Add(new PSNoteProperty("Metadata", metadata));
        return bundle;
    }

    public static void SetGlobalVariable(Runspace runspace, string variableName, object value)
    {
        runspace.SessionStateProxy.SetVariable($"global:{variableName}", value);
    }

    public static List<object> EnumerateRunningExcelApplications()
    {
        var apps = new List<object>();
        var dedup = new HashSet<string>(StringComparer.Ordinal);

        if (GetRunningObjectTable(0, out var rotResult) != 0 || rotResult is null)
        {
            return apps;
        }

        rotResult.EnumRunning(out var enumMoniker);
        if (enumMoniker is null)
        {
            return apps;
        }

        CreateBindCtx(0, out var bindCtx);
        var monikers = new IMoniker[1];
        while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
        {
            var moniker = monikers[0];
            if (moniker is null)
            {
                continue;
            }

            try
            {
                rotResult.GetObject(moniker, out var candidate);
                if (candidate is null)
                {
                    continue;
                }

                var application = TryGetExcelApplicationFromComObject(candidate);
                if (application is null)
                {
                    continue;
                }

                var key = BuildComIdentityKey(application);
                if (dedup.Add(key))
                {
                    apps.Add(application);
                }
            }
            catch
            {
                // Ignore individual ROT enumeration failures and continue.
            }
            finally
            {
                Marshal.ReleaseComObject(moniker);
            }
        }

        if (bindCtx is not null && Marshal.IsComObject(bindCtx))
        {
            Marshal.ReleaseComObject(bindCtx);
        }

        if (Marshal.IsComObject(enumMoniker))
        {
            Marshal.ReleaseComObject(enumMoniker);
        }

        if (Marshal.IsComObject(rotResult))
        {
            Marshal.ReleaseComObject(rotResult);
        }

        return apps;
    }

    public static object? CreateExcelApplication()
    {
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
        if (type is null)
        {
            return null;
        }

        return Activator.CreateInstance(type);
    }

    public static bool TryInvokeWithTimeout<T>(Func<T> action, TimeSpan timeout, out T? result, out Exception? error)
    {
        var done = new ManualResetEventSlim(false);
        T? localResult = default;
        Exception? localError = null;

        var thread = new Thread(() =>
        {
            try
            {
                localResult = action();
            }
            catch (Exception ex)
            {
                localError = ex;
            }
            finally
            {
                done.Set();
            }
        })
        {
            IsBackground = true,
            Name = "Excel.TimeoutGuard"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        if (!done.Wait(timeout))
        {
            result = default;
            error = null;
            return false;
        }

        result = localResult;
        error = localError;
        return true;
    }

    public static object? GetProperty(object target, string propertyName)
    {
        return target.GetType().InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            binder: null,
            target,
            args: null,
            CultureInfo.InvariantCulture);
    }

    public static object? InvokeMethod(object target, string methodName, params object?[]? args)
    {
        return target.GetType().InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            binder: null,
            target,
            args ?? [],
            CultureInfo.InvariantCulture);
    }

    public static void SetProperty(object target, string propertyName, object? value)
    {
        target.GetType().InvokeMember(
            propertyName,
            BindingFlags.SetProperty,
            binder: null,
            target,
            [value],
            CultureInfo.InvariantCulture);
    }

    public static void SafeReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        try
        {
            Marshal.FinalReleaseComObject(value);
        }
        catch
        {
            // Ignore release failures in cleanup paths.
        }
    }

    public static string NormalizeTarget(string path, out bool isUrl)
    {
        if (Uri.TryCreate(path, UriKind.Absolute, out var uri)
            && (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
                || uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)))
        {
            isUrl = true;
            return path.Trim();
        }

        isUrl = false;
        return Path.GetFullPath(path.Trim());
    }

    public static bool TargetExists(string normalizedTarget, bool isUrl)
    {
        return isUrl || File.Exists(normalizedTarget);
    }

    public static string NormalizeWorkbookName(string fullName, bool targetIsUrl)
    {
        if (string.IsNullOrWhiteSpace(fullName))
        {
            return string.Empty;
        }

        if (targetIsUrl)
        {
            return fullName.Trim();
        }

        try
        {
            return Path.GetFullPath(fullName.Trim());
        }
        catch
        {
            return fullName.Trim();
        }
    }

    public static string FileNameFromTarget(string target, bool isUrl)
    {
        if (isUrl && Uri.TryCreate(target, UriKind.Absolute, out var uri))
        {
            return Path.GetFileName(uri.LocalPath);
        }

        return Path.GetFileName(target);
    }

    public static bool IsHandleStored(Runspace runspace, object application)
    {
        using var ps = PowerShell.Create();
        ps.Runspace = runspace;
        ps.AddScript(
            "$global:__psbrokerFound = $false;" +
            "Get-Variable -Scope Global | ForEach-Object {" +
            "  $value = $_.Value;" +
            "  if ($null -ne $value -and $value.PSObject.Properties['Application'] -and $value.Application -eq $args[0]) { $global:__psbrokerFound = $true }" +
            "};" +
            "$global:__psbrokerFound")
            .AddArgument(application);
        var result = ps.Invoke();
        return result.Count > 0 && result[0].BaseObject is bool b && b;
    }

    private static bool TryGetProperty(JsonElement? args, string name, out JsonElement value)
    {
        value = default;
        if (args is null || args.Value.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        return args.Value.TryGetProperty(name, out value);
    }

    private static object? TryGetExcelApplicationFromComObject(object candidate)
    {
        try
        {
            var workbooks = GetProperty(candidate, "Workbooks");
            if (workbooks is not null)
            {
                SafeReleaseComObject(workbooks);
                return candidate;
            }
        }
        catch
        {
            // Not an application object.
        }

        try
        {
            var app = GetProperty(candidate, "Application");
            if (app is not null)
            {
                return app;
            }
        }
        catch
        {
            // Not a workbook object.
        }

        return null;
    }

    private static string BuildComIdentityKey(object comObject)
    {
        var ptr = Marshal.GetIUnknownForObject(comObject);
        try
        {
            return ptr.ToString("X");
        }
        finally
        {
            Marshal.Release(ptr);
        }
    }

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);
}

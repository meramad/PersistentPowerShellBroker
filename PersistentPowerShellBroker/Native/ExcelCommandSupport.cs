using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.VisualBasic;
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
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var hr = GetRunningObjectTable(0, out var rot);
            if (hr == 0 && rot is not null)
            {
                rot.EnumRunning(out var enumMoniker);
                if (enumMoniker is not null)
                {
                    try
                    {
                        var monikers = new IMoniker[1];
                        while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                        {
                            var moniker = monikers[0];
                            if (moniker is null)
                            {
                                continue;
                            }

                            object? runningObject = null;
                            object? appCandidate = null;
                            var appAdded = false;

                            try
                            {
                                rot.GetObject(moniker, out runningObject);
                                if (runningObject is null)
                                {
                                    continue;
                                }

                                appCandidate = TryGetExcelApplicationFromComObject(runningObject);
                                if (appCandidate is null)
                                {
                                    continue;
                                }

                                var key = BuildComIdentityKey(appCandidate);
                                if (seen.Add(key))
                                {
                                    apps.Add(appCandidate);
                                    appAdded = true;
                                }
                            }
                            catch
                            {
                                // Ignore non-Excel and inaccessible ROT entries.
                            }
                            finally
                            {
                                if (!appAdded
                                    && appCandidate is not null
                                    && !ReferenceEquals(appCandidate, runningObject))
                                {
                                    SafeReleaseComObject(appCandidate);
                                }

                                if (runningObject is not null
                                    && (!appAdded || !ReferenceEquals(runningObject, appCandidate)))
                                {
                                    SafeReleaseComObject(runningObject);
                                }

                                SafeReleaseComObject(moniker);
                            }
                        }
                    }
                    finally
                    {
                        SafeReleaseComObject(enumMoniker);
                    }
                }

                SafeReleaseComObject(rot);
            }
        }
        catch
        {
            // Continue to active-object fallback below.
        }

        if (apps.Count == 0)
        {
            try
            {
                var active = GetActiveObjectFromProgId("Excel.Application");
                if (active is not null)
                {
                    var key = BuildComIdentityKey(active);
                    if (seen.Add(key))
                    {
                        apps.Add(active);
                    }
                    else
                    {
                        SafeReleaseComObject(active);
                    }
                }
            }
            catch (COMException)
            {
                // No active Excel application.
            }
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
        return Interaction.CallByName(target, propertyName, CallType.Get);
    }

    public static object? InvokeMethod(object target, string methodName, params object?[]? args)
    {
        return Interaction.CallByName(target, methodName, CallType.Method, args ?? []);
    }

    public static void SetProperty(object target, string propertyName, object? value)
    {
        Interaction.CallByName(target, propertyName, CallType.Set, value);
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
        if (IsExcelApplicationObject(candidate))
        {
            return candidate;
        }

        try
        {
            var app = GetProperty(candidate, "Application");
            if (app is not null && IsExcelApplicationObject(app))
            {
                return app;
            }

            SafeReleaseComObject(app);
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

    private static bool IsExcelApplicationObject(object candidate)
    {
        try
        {
            var workbooks = GetProperty(candidate, "Workbooks");
            var hwnd = GetProperty(candidate, "Hwnd");
            SafeReleaseComObject(workbooks);
            return workbooks is not null && hwnd is not null;
        }
        catch
        {
            return false;
        }
    }

    private static object? GetActiveObjectFromProgId(string progId)
    {
        var hr = CLSIDFromProgID(progId, out var clsid);
        if (hr != 0)
        {
            throw new COMException($"CLSIDFromProgID failed for '{progId}'.", hr);
        }

        hr = GetActiveObject(ref clsid, IntPtr.Zero, out var activeObject);
        if (hr != 0)
        {
            throw new COMException($"GetActiveObject failed for '{progId}'.", hr);
        }

        return activeObject;
    }

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable? pprot);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);
}

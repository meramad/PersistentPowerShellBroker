using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerExcelGetWorkbookHandleCommand : INativeCommand
{
    private const string ReuseIfRunning = "ReuseIfRunning";
    private const string AlwaysNew = "AlwaysNew";
    private const int DefaultTimeoutSeconds = 15;
    private const int RpcCallRejected = unchecked((int)0x80010001);

    public string Name => "broker.excel.get_workbook_handle";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, Runspace runspace, CancellationToken cancellationToken)
    {
        if (!ExcelCommandSupport.TryGetString(args, "path", out var path) || string.IsNullOrWhiteSpace(path))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "OpenFailed",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: path ?? string.Empty,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: ReuseIfRunning,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidArgs",
                errorMessage: "Missing required 'path'."));
        }

        if (!ExcelCommandSupport.TryGetBool(args, "readOnly", out var readOnlyArg)
            || !ExcelCommandSupport.TryGetString(args, "openPassword", out var openPassword)
            || !ExcelCommandSupport.TryGetString(args, "modifyPassword", out var modifyPassword)
            || !ExcelCommandSupport.TryGetInt(args, "timeoutSeconds", out var timeoutSecondsArg)
            || !ExcelCommandSupport.TryGetString(args, "instancePolicy", out var instancePolicyArg)
            || !ExcelCommandSupport.TryGetBool(args, "displayAlerts", out var displayAlertsArg)
            || !ExcelCommandSupport.TryGetBool(args, "forceVisible", out var forceVisibleArg))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "OpenFailed",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: path,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: ReuseIfRunning,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidArgs",
                errorMessage: "Invalid argument types."));
        }

        var readOnly = readOnlyArg ?? false;
        var timeoutSeconds = timeoutSecondsArg ?? DefaultTimeoutSeconds;
        if (timeoutSeconds < 1)
        {
            timeoutSeconds = DefaultTimeoutSeconds;
        }

        var instancePolicy = NormalizeInstancePolicy(instancePolicyArg);
        var displayAlerts = displayAlertsArg ?? false;
        var forceVisible = forceVisibleArg ?? true;

        WorkbookIdentityResolution identity;
        try
        {
            identity = LocalWorkbookIdentityResolver.Resolve(path);
        }
        catch (Exception ex)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "InvalidPath",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: path,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidPath",
                errorMessage: ex.Message));
        }

        if (!identity.InputIsUrl && !File.Exists(identity.NormalizedLocalPath))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "FileNotFound",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: identity.NormalizedLocalPath,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "FileNotFound",
                errorMessage: "Workbook target was not found."));
        }

        var requestedTarget = identity.InputIsUrl
            ? identity.NormalizedRemoteUrl ?? identity.RequestedInput
            : identity.NormalizedLocalPath;

        var runningApps = ExcelCommandSupport.EnumerateRunningExcelApplications();
        var locator = new ExcelWorkbookLocator();
        var match = locator.FindOpenWorkbook(
            runningApps,
            identity,
            allowFileNameFallback: instancePolicy == ReuseIfRunning);

        if (match.AmbiguousMatches.Count > 1)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "AmbiguousMatch",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: requestedTarget,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "AmbiguousMatch",
                errorMessage: $"Multiple workbook candidates: {string.Join(" | ", match.AmbiguousMatches)}"));
        }

        if (match.MatchedCandidate is not null)
        {
            var workbookFullName = Convert.ToString(ExcelCommandSupport.GetProperty(match.MatchedCandidate.Workbook, "FullName")) ?? requestedTarget;
            var isReadOnly = Convert.ToBoolean(ExcelCommandSupport.GetProperty(match.MatchedCandidate.Workbook, "ReadOnly") ?? false);
            var variableName = ExcelCommandSupport.NewHandleVariableName();
            var bundle = ExcelCommandSupport.BuildBundle(
                match.MatchedCandidate.Application,
                match.MatchedCandidate.Workbook,
                requestedTarget,
                workbookFullName,
                isReadOnly,
                attachedExisting: true,
                openedWorkbook: false,
                instancePolicy);
            ExcelCommandSupport.SetGlobalVariable(runspace, variableName, bundle);

            return Task.FromResult(BuildResult(
                ok: true,
                status: "Success",
                psVariableName: variableName,
                workbookFullName: workbookFullName,
                requestedTarget: requestedTarget,
                attachedExisting: true,
                openedWorkbook: false,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: isReadOnly,
                readOnlyReason: isReadOnly ? "FileLockedOrWriteDenied" : null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: null,
                errorMessage: null));
        }

        object? application = null;
        var createdApplication = false;
        var handleStored = false;
        try
        {
            // Deliberately create a fresh instance when workbook is not already open.
            application = ExcelCommandSupport.CreateExcelApplication();
            createdApplication = true;
            if (application is null)
            {
                return Task.FromResult(BuildResult(
                    ok: false,
                    status: "ExcelNotInstalledOrCOMUnavailable",
                    psVariableName: null,
                    workbookFullName: null,
                    requestedTarget: requestedTarget,
                    attachedExisting: false,
                    openedWorkbook: false,
                    excelInstancePolicyUsed: instancePolicy,
                    isReadOnly: null,
                    readOnlyReason: null,
                    blockedLikely: false,
                    blockingHint: null,
                    errorCode: "ExcelNotInstalledOrCOMUnavailable",
                    errorMessage: "Excel COM automation is unavailable."));
            }

            if (forceVisible)
            {
                try
                {
                    ExcelCommandSupport.SetProperty(application, "Visible", true);
                }
                catch
                {
                    // Best effort only.
                }
            }

            if (!displayAlerts)
            {
                try
                {
                    ExcelCommandSupport.SetProperty(application, "DisplayAlerts", false);
                }
                catch
                {
                    // Best effort only.
                }
            }

            var opener = new ExcelWorkbookOpener();
            object? workbook;
            try
            {
                workbook = opener.OpenWorkbookWithRetry(
                    application,
                    requestedTarget,
                    readOnly,
                    openPassword,
                    modifyPassword,
                    TimeSpan.FromSeconds(timeoutSeconds));
            }
            catch (Exception openError)
            {
                var mapped = MapOpenException(openError);
                return Task.FromResult(BuildResult(
                    ok: false,
                    status: mapped.Status,
                    psVariableName: null,
                    workbookFullName: null,
                    requestedTarget: requestedTarget,
                    attachedExisting: false,
                    openedWorkbook: false,
                    excelInstancePolicyUsed: instancePolicy,
                    isReadOnly: null,
                    readOnlyReason: null,
                    blockedLikely: mapped.BlockedLikely,
                    blockingHint: mapped.BlockingHint,
                    errorCode: mapped.ErrorCode,
                    errorMessage: mapped.ErrorMessage));
            }

            if (workbook is null)
            {
                return Task.FromResult(BuildResult(
                    ok: false,
                    status: "OpenFailed",
                    psVariableName: null,
                    workbookFullName: null,
                    requestedTarget: requestedTarget,
                    attachedExisting: false,
                    openedWorkbook: false,
                    excelInstancePolicyUsed: instancePolicy,
                    isReadOnly: null,
                    readOnlyReason: null,
                    blockedLikely: false,
                    blockingHint: null,
                    errorCode: "OpenFailed",
                    errorMessage: "Excel did not return a workbook object."));
            }

            var workbookName = Convert.ToString(ExcelCommandSupport.GetProperty(workbook, "FullName")) ?? requestedTarget;
            var workbookReadOnly = Convert.ToBoolean(ExcelCommandSupport.GetProperty(workbook, "ReadOnly") ?? false);
            var handleName = ExcelCommandSupport.NewHandleVariableName();
            var bundleValue = ExcelCommandSupport.BuildBundle(
                application,
                workbook,
                requestedTarget,
                workbookName,
                workbookReadOnly,
                attachedExisting: false,
                openedWorkbook: true,
                instancePolicy);
            ExcelCommandSupport.SetGlobalVariable(runspace, handleName, bundleValue);
            handleStored = true;

            return Task.FromResult(BuildResult(
                ok: true,
                status: "Success",
                psVariableName: handleName,
                workbookFullName: workbookName,
                requestedTarget: requestedTarget,
                attachedExisting: false,
                openedWorkbook: true,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: workbookReadOnly,
                readOnlyReason: workbookReadOnly ? "FileLockedOrWriteDenied" : null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: null,
                errorMessage: null));
        }
        catch (Exception ex)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "OpenFailed",
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: requestedTarget,
                attachedExisting: false,
                openedWorkbook: false,
                excelInstancePolicyUsed: instancePolicy,
                isReadOnly: null,
                readOnlyReason: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "OpenFailed",
                errorMessage: ex.Message));
        }
        finally
        {
            if (createdApplication && application is not null && !handleStored)
            {
                ExcelCommandSupport.SafeReleaseComObject(application);
            }
        }
    }

    private static (string Status, string ErrorCode, string ErrorMessage, bool BlockedLikely, string? BlockingHint) MapOpenException(Exception ex)
    {
        if (ex is ExcelWorkbookOpener.ComRetryTimeoutException)
        {
            return ("ComBusyRetryTimeout", "ComBusyRetryTimeout", ex.Message, true, "Excel was busy and kept rejecting calls.");
        }

        if (ex is COMException comEx && comEx.HResult == RpcCallRejected)
        {
            return ("ComCallRejected", "ComCallRejected", ex.Message, true, "Excel rejected the call while busy.");
        }

        if (ex is TimeoutException)
        {
            return ("CommandTimeout", "CommandTimeout", ex.Message, true, "Operation exceeded command timeout.");
        }

        var text = ex.ToString();
        if (text.Contains("password", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("modify", StringComparison.OrdinalIgnoreCase))
            {
                return ("ModifyPasswordRequired", "ModifyPasswordRequired", ex.Message, false, null);
            }

            if (text.Contains("required", StringComparison.OrdinalIgnoreCase))
            {
                return ("PasswordRequired", "PasswordRequired", ex.Message, false, null);
            }

            return ("InvalidPassword", "InvalidPassword", ex.Message, false, null);
        }

        return ("OpenFailed", "OpenFailed", ex.Message, false, null);
    }

    private static string NormalizeInstancePolicy(string? policy)
    {
        if (string.Equals(policy, AlwaysNew, StringComparison.OrdinalIgnoreCase))
        {
            return AlwaysNew;
        }

        return ReuseIfRunning;
    }

    private static NativeResult BuildResult(
        bool ok,
        string status,
        string? psVariableName,
        string? workbookFullName,
        string requestedTarget,
        bool attachedExisting,
        bool openedWorkbook,
        string excelInstancePolicyUsed,
        bool? isReadOnly,
        string? readOnlyReason,
        bool blockedLikely,
        string? blockingHint,
        string? errorCode,
        string? errorMessage)
    {
        var payload = new
        {
            ok,
            status,
            psVariableName,
            workbookFullName,
            requestedTarget,
            attachedExisting,
            openedWorkbook,
            excelInstancePolicyUsed,
            isReadOnly,
            readOnlyReason,
            blockedLikely,
            blockingHint,
            errorCode,
            errorMessage
        };

        return ExcelCommandSupport.BuildErrorResult(status, errorCode, errorMessage, payload);
    }
}

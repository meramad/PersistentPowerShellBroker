using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerExcelGetWorkbookHandleCommand : INativeCommand
{
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
                isReadOnly: null,
                readOnlyReason: null,
                createdApplicationByBroker: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidArgs",
                errorMessage: "Missing required 'path'."));
        }

        if (!ExcelCommandSupport.TryGetBool(args, "readOnly", out var readOnlyArg)
            || !ExcelCommandSupport.TryGetString(args, "openPassword", out var openPassword)
            || !ExcelCommandSupport.TryGetString(args, "modifyPassword", out var modifyPassword)
            || !ExcelCommandSupport.TryGetInt(args, "timeoutSeconds", out var timeoutSecondsArg)
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
                isReadOnly: null,
                readOnlyReason: null,
                createdApplicationByBroker: null,
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
                isReadOnly: null,
                readOnlyReason: null,
                createdApplicationByBroker: null,
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
                isReadOnly: null,
                readOnlyReason: null,
                createdApplicationByBroker: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "FileNotFound",
                errorMessage: "Workbook target was not found."));
        }

        AcquireWorkbookResult acquired;
        try
        {
            var service = new ExcelWorkbookHandleService();
            acquired = service.AcquireWorkbook(
                identity,
                new AcquireWorkbookOptions(
                    ReadOnly: readOnly,
                    OpenPassword: openPassword,
                    ModifyPassword: modifyPassword,
                    TimeoutSeconds: timeoutSeconds,
                    ForceVisible: forceVisible,
                    DisplayAlerts: displayAlerts));
        }
        catch (Exception ex)
        {
            var mapped = MapOpenException(ex);
            return Task.FromResult(BuildResult(
                ok: false,
                status: mapped.Status,
                psVariableName: null,
                workbookFullName: null,
                requestedTarget: identity.InputIsUrl ? identity.NormalizedRemoteUrl ?? identity.RequestedInput : identity.NormalizedLocalPath,
                attachedExisting: false,
                openedWorkbook: false,
                isReadOnly: null,
                readOnlyReason: null,
                createdApplicationByBroker: null,
                blockedLikely: mapped.BlockedLikely,
                blockingHint: mapped.BlockingHint,
                errorCode: mapped.ErrorCode,
                errorMessage: mapped.ErrorMessage));
        }

        if (!acquired.Ok || acquired.Session is null || acquired.Workbook is null)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: acquired.Status,
                psVariableName: null,
                workbookFullName: acquired.WorkbookFullName,
                requestedTarget: acquired.RequestedTarget,
                attachedExisting: acquired.AttachedExisting,
                openedWorkbook: acquired.OpenedWorkbook,
                isReadOnly: acquired.IsReadOnly,
                readOnlyReason: acquired.ReadOnlyReason,
                createdApplicationByBroker: null,
                blockedLikely: false,
                blockingHint: null,
                errorCode: acquired.ErrorCode,
                errorMessage: acquired.ErrorMessage));
        }

        var handleName = ExcelCommandSupport.NewHandleVariableName();
        var bundleValue = ExcelCommandSupport.BuildBundle(
            acquired.Session.Application,
            acquired.Workbook,
            acquired.RequestedTarget,
            acquired.WorkbookFullName ?? acquired.RequestedTarget,
            acquired.IsReadOnly ?? false,
            acquired.AttachedExisting,
            acquired.OpenedWorkbook,
            acquired.Session.CreatedByBroker);
        ExcelCommandSupport.SetGlobalVariable(runspace, handleName, bundleValue);

        ExcelHandleRegistry.Register(new ExcelHandleMetadata(
            VariableName: handleName,
            RequestedTarget: acquired.RequestedTarget,
            WorkbookFullName: acquired.WorkbookFullName ?? acquired.RequestedTarget,
            AttachedExisting: acquired.AttachedExisting,
            OpenedWorkbook: acquired.OpenedWorkbook,
            IsReadOnly: acquired.IsReadOnly ?? false,
            CreatedApplicationByBroker: acquired.Session.CreatedByBroker,
            CreatedUtc: DateTime.UtcNow));

        return Task.FromResult(BuildResult(
            ok: true,
            status: "Success",
            psVariableName: handleName,
            workbookFullName: acquired.WorkbookFullName,
            requestedTarget: acquired.RequestedTarget,
            attachedExisting: acquired.AttachedExisting,
            openedWorkbook: acquired.OpenedWorkbook,
            isReadOnly: acquired.IsReadOnly,
            readOnlyReason: acquired.ReadOnlyReason,
            createdApplicationByBroker: acquired.Session.CreatedByBroker,
            blockedLikely: false,
            blockingHint: null,
            errorCode: null,
            errorMessage: null));
    }

    private static (string Status, string ErrorCode, string ErrorMessage, bool BlockedLikely, string? BlockingHint) MapOpenException(Exception ex)
    {
        var root = ex is InvalidOperationException { InnerException: not null } ? ex.InnerException : ex;

        if (root is ExcelWorkbookOpener.ComRetryTimeoutException)
        {
            return ("ComBusyRetryTimeout", "ComBusyRetryTimeout", root.Message, true, "Excel was busy and kept rejecting calls.");
        }

        if (root is COMException comEx && comEx.HResult == RpcCallRejected)
        {
            return ("ComCallRejected", "ComCallRejected", comEx.Message, true, "Excel rejected the call while busy.");
        }

        if (root is TimeoutException)
        {
            return ("CommandTimeout", "CommandTimeout", root.Message, true, "Operation exceeded command timeout.");
        }

        var text = root.ToString();
        if (text.Contains("password", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("modify", StringComparison.OrdinalIgnoreCase))
            {
                return ("ModifyPasswordRequired", "ModifyPasswordRequired", root.Message, false, null);
            }

            if (text.Contains("required", StringComparison.OrdinalIgnoreCase))
            {
                return ("PasswordRequired", "PasswordRequired", root.Message, false, null);
            }

            return ("InvalidPassword", "InvalidPassword", root.Message, false, null);
        }

        return ("OpenFailed", "OpenFailed", $"{ex.Message} (root: {root.Message})", false, null);
    }

    private static NativeResult BuildResult(
        bool ok,
        string status,
        string? psVariableName,
        string? workbookFullName,
        string requestedTarget,
        bool attachedExisting,
        bool openedWorkbook,
        bool? isReadOnly,
        string? readOnlyReason,
        bool? createdApplicationByBroker,
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
            isReadOnly,
            readOnlyReason,
            createdApplicationByBroker,
            blockedLikely,
            blockingHint,
            errorCode,
            errorMessage
        };

        return ExcelCommandSupport.BuildErrorResult(status, errorCode, errorMessage, payload);
    }
}

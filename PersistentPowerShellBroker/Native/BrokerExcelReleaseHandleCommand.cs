using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerExcelReleaseHandleCommand : INativeCommand
{
    private const int DefaultTimeoutSeconds = 10;

    public string Name => "broker.excel.release_handle";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, Runspace runspace, CancellationToken cancellationToken)
    {
        if (!ExcelCommandSupport.TryGetString(args, "psVariableName", out var psVariableName)
            || string.IsNullOrWhiteSpace(psVariableName))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "InvalidHandle",
                psVariableName: psVariableName ?? string.Empty,
                workbookFullName: null,
                closedWorkbook: false,
                quitExcelAttempted: false,
                quitExcelSucceeded: false,
                quitSkipped: false,
                quitSkipReason: null,
                released: false,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidArgs",
                errorMessage: "Missing required 'psVariableName'."));
        }

        if (!ExcelCommandSupport.TryGetBool(args, "closeWorkbook", out var closeWorkbookArg)
            || !ExcelCommandSupport.TryGetBool(args, "saveChanges", out var saveChangesArg)
            || !ExcelCommandSupport.TryGetBool(args, "quitExcel", out var quitExcelArg)
            || !ExcelCommandSupport.TryGetBool(args, "onlyIfNoOtherWorkbooks", out var onlyIfNoOtherArg)
            || !ExcelCommandSupport.TryGetInt(args, "timeoutSeconds", out var timeoutSecondsArg)
            || !ExcelCommandSupport.TryGetBool(args, "displayAlerts", out var displayAlertsArg))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "InvalidHandle",
                psVariableName: psVariableName,
                workbookFullName: null,
                closedWorkbook: false,
                quitExcelAttempted: false,
                quitExcelSucceeded: false,
                quitSkipped: false,
                quitSkipReason: null,
                released: false,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidArgs",
                errorMessage: "Invalid argument types."));
        }

        var closeWorkbook = closeWorkbookArg ?? false;
        var quitExcel = quitExcelArg ?? false;
        var onlyIfNoOtherWorkbooks = onlyIfNoOtherArg ?? true;
        var timeoutSeconds = timeoutSecondsArg ?? DefaultTimeoutSeconds;
        if (timeoutSeconds < 1)
        {
            timeoutSeconds = DefaultTimeoutSeconds;
        }

        var displayAlerts = displayAlertsArg ?? false;
        var timeout = TimeSpan.FromSeconds(timeoutSeconds);

        if (!ExcelCommandSupport.TryGetVariable(runspace, psVariableName, out var variableValue))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "NotFound",
                psVariableName: psVariableName,
                workbookFullName: null,
                closedWorkbook: false,
                quitExcelAttempted: false,
                quitExcelSucceeded: false,
                quitSkipped: false,
                quitSkipReason: null,
                released: false,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "NotFound",
                errorMessage: "Handle variable was not found."));
        }

        var bundle = variableValue as PSObject ?? PSObject.AsPSObject(variableValue);
        var application = bundle.Properties["Application"]?.Value;
        var workbook = bundle.Properties["Workbook"]?.Value;
        if (application is null || workbook is null)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "InvalidHandle",
                psVariableName: psVariableName,
                workbookFullName: null,
                closedWorkbook: false,
                quitExcelAttempted: false,
                quitExcelSucceeded: false,
                quitSkipped: false,
                quitSkipReason: null,
                released: false,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "InvalidHandle",
                errorMessage: "Handle bundle is missing Application or Workbook."));
        }

        var workbookFullName = Convert.ToString(ExcelCommandSupport.GetProperty(workbook, "FullName"));
        var closedWorkbook = false;
        var quitAttempted = false;
        var quitSucceeded = false;
        var quitSkipped = false;
        string? quitSkipReason = null;

        try
        {
            if (closeWorkbook)
            {
                var closeWorked = ExcelCommandSupport.TryInvokeWithTimeout(
                    () =>
                    {
                        ExcelCommandSupport.SetProperty(application, "Visible", true);
                        if (!displayAlerts)
                        {
                            ExcelCommandSupport.SetProperty(application, "DisplayAlerts", false);
                        }

                        object? saveFlag = saveChangesArg.HasValue ? saveChangesArg.Value : Type.Missing;
                        ExcelCommandSupport.InvokeMethod(workbook, "Close", saveFlag);
                        return true;
                    },
                    timeout,
                    out _,
                    out var closeError);
                if (!closeWorked)
                {
                    return Task.FromResult(BuildResult(
                        ok: false,
                        status: "TimeoutLikelyModalDialog",
                        psVariableName: psVariableName,
                        workbookFullName: workbookFullName,
                        closedWorkbook: false,
                        quitExcelAttempted: false,
                        quitExcelSucceeded: false,
                        quitSkipped: false,
                        quitSkipReason: null,
                        released: false,
                        blockedLikely: true,
                        blockingHint: "Excel appears blocked by a modal dialog; user action required",
                        errorCode: "TimeoutLikelyModalDialog",
                        errorMessage: "Timed out while closing workbook."));
                }

                if (closeError is not null)
                {
                    return Task.FromResult(BuildResult(
                        ok: false,
                        status: "CloseFailed",
                        psVariableName: psVariableName,
                        workbookFullName: workbookFullName,
                        closedWorkbook: false,
                        quitExcelAttempted: false,
                        quitExcelSucceeded: false,
                        quitSkipped: false,
                        quitSkipReason: null,
                        released: false,
                        blockedLikely: false,
                        blockingHint: null,
                        errorCode: "CloseFailed",
                        errorMessage: closeError.Message));
                }

                closedWorkbook = true;
            }

            if (quitExcel)
            {
                quitAttempted = true;
                if (onlyIfNoOtherWorkbooks)
                {
                    var workbooks = ExcelCommandSupport.GetProperty(application, "Workbooks");
                    var openCount = Convert.ToInt32(ExcelCommandSupport.GetProperty(workbooks!, "Count") ?? 0);
                    ExcelCommandSupport.SafeReleaseComObject(workbooks);
                    var hasOther = closeWorkbook ? openCount > 0 : openCount > 1;
                    if (hasOther)
                    {
                        quitSkipped = true;
                        quitSkipReason = "OtherWorkbooksOpen";
                    }
                }

                if (!quitSkipped)
                {
                    var quitWorked = ExcelCommandSupport.TryInvokeWithTimeout(
                        () =>
                        {
                            if (!displayAlerts)
                            {
                                ExcelCommandSupport.SetProperty(application, "DisplayAlerts", false);
                            }

                            ExcelCommandSupport.InvokeMethod(application, "Quit");
                            return true;
                        },
                        timeout,
                        out _,
                        out var quitError);
                    if (!quitWorked)
                    {
                        return Task.FromResult(BuildResult(
                            ok: false,
                            status: "TimeoutLikelyModalDialog",
                            psVariableName: psVariableName,
                            workbookFullName: workbookFullName,
                            closedWorkbook: closedWorkbook,
                            quitExcelAttempted: true,
                            quitExcelSucceeded: false,
                            quitSkipped: false,
                            quitSkipReason: null,
                            released: false,
                            blockedLikely: true,
                            blockingHint: "Excel appears blocked by a modal dialog; user action required",
                            errorCode: "TimeoutLikelyModalDialog",
                            errorMessage: "Timed out while quitting Excel."));
                    }

                    if (quitError is not null)
                    {
                        return Task.FromResult(BuildResult(
                            ok: false,
                            status: "QuitFailed",
                            psVariableName: psVariableName,
                            workbookFullName: workbookFullName,
                            closedWorkbook: closedWorkbook,
                            quitExcelAttempted: true,
                            quitExcelSucceeded: false,
                            quitSkipped: false,
                            quitSkipReason: null,
                            released: false,
                            blockedLikely: false,
                            blockingHint: null,
                            errorCode: "QuitFailed",
                            errorMessage: quitError.Message));
                    }

                    quitSucceeded = true;
                }
            }

            ExcelCommandSupport.RemoveGlobalVariable(runspace, psVariableName);
            ExcelCommandSupport.SafeReleaseComObject(workbook);
            ExcelCommandSupport.SafeReleaseComObject(application);

            return Task.FromResult(BuildResult(
                ok: true,
                status: "Success",
                psVariableName: psVariableName,
                workbookFullName: workbookFullName,
                closedWorkbook: closedWorkbook,
                quitExcelAttempted: quitAttempted,
                quitExcelSucceeded: quitSucceeded,
                quitSkipped: quitSkipped,
                quitSkipReason: quitSkipReason,
                released: true,
                blockedLikely: false,
                blockingHint: null,
                errorCode: null,
                errorMessage: null));
        }
        catch (Exception ex)
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "ReleaseFailed",
                psVariableName: psVariableName,
                workbookFullName: workbookFullName,
                closedWorkbook: closedWorkbook,
                quitExcelAttempted: quitAttempted,
                quitExcelSucceeded: quitSucceeded,
                quitSkipped: quitSkipped,
                quitSkipReason: quitSkipReason,
                released: false,
                blockedLikely: false,
                blockingHint: null,
                errorCode: "ReleaseFailed",
                errorMessage: ex.Message));
        }
    }

    private static NativeResult BuildResult(
        bool ok,
        string status,
        string psVariableName,
        string? workbookFullName,
        bool closedWorkbook,
        bool quitExcelAttempted,
        bool quitExcelSucceeded,
        bool quitSkipped,
        string? quitSkipReason,
        bool released,
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
            closedWorkbook,
            quitExcelAttempted,
            quitExcelSucceeded,
            quitSkipped,
            quitSkipReason,
            released,
            blockedLikely,
            blockingHint,
            errorCode,
            errorMessage
        };

        return ExcelCommandSupport.BuildErrorResult(status, errorCode, errorMessage, payload);
    }
}

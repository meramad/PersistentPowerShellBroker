using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerExcelReleaseHandleCommand : INativeCommand
{
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
        var displayAlerts = displayAlertsArg ?? false;

        if (!ExcelCommandSupport.TryGetVariable(runspace, psVariableName, out var variableValue))
        {
            ExcelHandleRegistry.Remove(psVariableName);
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
        if (application is PSObject appPsObject)
        {
            application = appPsObject.BaseObject;
        }

        if (workbook is PSObject workbookPsObject)
        {
            workbook = workbookPsObject.BaseObject;
        }

        if (application is null || workbook is null)
        {
            ExcelHandleRegistry.Remove(psVariableName);
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

        ExcelHandleRegistry.TryGet(psVariableName, out var metadata);
        var session = new ExcelApplicationSession(application, metadata?.CreatedApplicationByBroker ?? false);

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
                try
                {
                    session.EnsureVisible(forceVisible: true, workbook);
                    if (!displayAlerts)
                    {
                        session.SetDisplayAlerts(enabled: false);
                    }

                    object? saveFlag = saveChangesArg.HasValue ? saveChangesArg.Value : Type.Missing;
                    ExcelCommandSupport.InvokeMethod(workbook, "Close", saveFlag);
                }
                catch (Exception closeError)
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

                if (!session.CreatedByBroker)
                {
                    quitSkipped = true;
                    quitSkipReason = "NotBrokerOwnedApplication";
                }

                if (!quitSkipped && onlyIfNoOtherWorkbooks)
                {
                    var openCount = session.GetOpenWorkbookCount();
                    var hasOther = closeWorkbook ? openCount > 0 : openCount > 1;
                    if (hasOther)
                    {
                        quitSkipped = true;
                        quitSkipReason = "OtherWorkbooksOpen";
                    }
                }

                if (!quitSkipped)
                {
                    try
                    {
                        if (!displayAlerts)
                        {
                            session.SetDisplayAlerts(enabled: false);
                        }

                        session.Quit();
                        quitSucceeded = true;
                    }
                    catch (Exception quitError)
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
                }
            }

            ExcelCommandSupport.RemoveGlobalVariable(runspace, psVariableName);
            ExcelHandleRegistry.Remove(psVariableName);
            ExcelCommandSupport.SafeReleaseComObject(workbook);
            session.Release();

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

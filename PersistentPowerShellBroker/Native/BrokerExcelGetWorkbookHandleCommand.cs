using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerExcelGetWorkbookHandleCommand : INativeCommand
{
    private const string ReuseIfRunning = "ReuseIfRunning";
    private const string AlwaysNew = "AlwaysNew";
    private const int DefaultTimeoutSeconds = 15;

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
        var requestedTarget = ExcelCommandSupport.NormalizeTarget(path, out var targetIsUrl);
        if (!ExcelCommandSupport.TargetExists(requestedTarget, targetIsUrl))
        {
            return Task.FromResult(BuildResult(
                ok: false,
                status: "FileNotFound",
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
                errorCode: "FileNotFound",
                errorMessage: "Workbook target was not found."));
        }

        var runningApps = ExcelCommandSupport.EnumerateRunningExcelApplications();
        var match = FindOpenWorkbook(runningApps, requestedTarget, targetIsUrl);
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

        if (match.MatchedWorkbook is not null && match.MatchedApplication is not null)
        {
            var workbookFullName = Convert.ToString(ExcelCommandSupport.GetProperty(match.MatchedWorkbook, "FullName")) ?? requestedTarget;
            var isReadOnly = Convert.ToBoolean(ExcelCommandSupport.GetProperty(match.MatchedWorkbook, "ReadOnly") ?? false);
            var variableName = ExcelCommandSupport.NewHandleVariableName();
            var bundle = ExcelCommandSupport.BuildBundle(
                match.MatchedApplication,
                match.MatchedWorkbook,
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
        try
        {
            if (instancePolicy == ReuseIfRunning && runningApps.Count > 0)
            {
                application = runningApps[0];
            }
            else
            {
                application = ExcelCommandSupport.CreateExcelApplication();
                createdApplication = true;
            }

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
                ExcelCommandSupport.SetProperty(application, "Visible", true);
            }

            if (!displayAlerts)
            {
                ExcelCommandSupport.SetProperty(application, "DisplayAlerts", false);
            }

            var timeout = TimeSpan.FromSeconds(timeoutSeconds);
            var openWorked = ExcelCommandSupport.TryInvokeWithTimeout(
                () => OpenWorkbook(application, requestedTarget, readOnly, openPassword, modifyPassword),
                timeout,
                out object? workbook,
                out var openError);
            if (!openWorked)
            {
                return Task.FromResult(BuildResult(
                    ok: false,
                    status: "TimeoutLikelyModalDialog",
                    psVariableName: null,
                    workbookFullName: null,
                    requestedTarget: requestedTarget,
                    attachedExisting: false,
                    openedWorkbook: false,
                    excelInstancePolicyUsed: instancePolicy,
                    isReadOnly: null,
                    readOnlyReason: null,
                    blockedLikely: true,
                    blockingHint: "Excel appears blocked by a modal dialog; user action required",
                    errorCode: "TimeoutLikelyModalDialog",
                    errorMessage: "Timed out while opening workbook."));
            }

            if (openError is not null)
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
                    blockedLikely: false,
                    blockingHint: null,
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
            if (createdApplication && application is not null && !ExcelCommandSupport.IsHandleStored(runspace, application))
            {
                ExcelCommandSupport.SafeReleaseComObject(application);
            }
        }
    }

    private static object? OpenWorkbook(object application, string target, bool readOnly, string? openPassword, string? modifyPassword)
    {
        var workbooks = ExcelCommandSupport.GetProperty(application, "Workbooks")
            ?? throw new InvalidOperationException("Excel.Workbooks is unavailable.");

        try
        {
            return ExcelCommandSupport.InvokeMethod(
                workbooks,
                "Open",
                target,
                Type.Missing,
                readOnly,
                Type.Missing,
                openPassword ?? Type.Missing,
                modifyPassword ?? Type.Missing);
        }
        finally
        {
            ExcelCommandSupport.SafeReleaseComObject(workbooks);
        }
    }

    private static (object? MatchedApplication, object? MatchedWorkbook, List<string> AmbiguousMatches) FindOpenWorkbook(
        IReadOnlyList<object> applications,
        string requestedTarget,
        bool targetIsUrl)
    {
        var exactMatches = new List<(object App, object Workbook, string FullName)>();
        var fileMatches = new List<(object App, object Workbook, string FullName)>();
        var requestedFileName = ExcelCommandSupport.FileNameFromTarget(requestedTarget, targetIsUrl);

        foreach (var app in applications)
        {
            var workbooks = ExcelCommandSupport.GetProperty(app, "Workbooks");
            if (workbooks is null)
            {
                continue;
            }

            try
            {
                var count = Convert.ToInt32(ExcelCommandSupport.GetProperty(workbooks, "Count") ?? 0);
                for (var i = 1; i <= count; i++)
                {
                    var workbook = ExcelCommandSupport.InvokeMethod(workbooks, "Item", i);
                    if (workbook is null)
                    {
                        continue;
                    }

                    var fullName = Convert.ToString(ExcelCommandSupport.GetProperty(workbook, "FullName")) ?? string.Empty;
                    var normalizedWorkbook = ExcelCommandSupport.NormalizeWorkbookName(fullName, targetIsUrl);
                    if (string.Equals(normalizedWorkbook, requestedTarget, targetIsUrl ? StringComparison.OrdinalIgnoreCase : StringComparison.OrdinalIgnoreCase))
                    {
                        exactMatches.Add((app, workbook, fullName));
                        continue;
                    }

                    var workbookName = ExcelCommandSupport.FileNameFromTarget(fullName, IsHttpUrl(fullName));
                    if (!string.IsNullOrWhiteSpace(workbookName)
                        && string.Equals(workbookName, requestedFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        fileMatches.Add((app, workbook, fullName));
                    }
                    else
                    {
                        ExcelCommandSupport.SafeReleaseComObject(workbook);
                    }
                }
            }
            finally
            {
                ExcelCommandSupport.SafeReleaseComObject(workbooks);
            }
        }

        if (exactMatches.Count == 1)
        {
            foreach (var candidate in fileMatches)
            {
                ExcelCommandSupport.SafeReleaseComObject(candidate.Workbook);
            }

            return (exactMatches[0].App, exactMatches[0].Workbook, []);
        }

        if (exactMatches.Count > 1)
        {
            var candidates = exactMatches.Select(match => match.FullName).ToList();
            foreach (var match in exactMatches)
            {
                ExcelCommandSupport.SafeReleaseComObject(match.Workbook);
            }

            foreach (var match in fileMatches)
            {
                ExcelCommandSupport.SafeReleaseComObject(match.Workbook);
            }

            return (null, null, candidates);
        }

        if (fileMatches.Count == 1)
        {
            return (fileMatches[0].App, fileMatches[0].Workbook, []);
        }

        var ambiguous = fileMatches.Select(match => match.FullName).ToList();
        foreach (var match in fileMatches)
        {
            ExcelCommandSupport.SafeReleaseComObject(match.Workbook);
        }

        return (null, null, ambiguous);
    }

    private static (string Status, string ErrorCode, string ErrorMessage) MapOpenException(Exception ex)
    {
        var text = ex.ToString();
        if (text.Contains("password", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("modify", StringComparison.OrdinalIgnoreCase))
            {
                return ("ModifyPasswordRequired", "ModifyPasswordRequired", ex.Message);
            }

            if (text.Contains("required", StringComparison.OrdinalIgnoreCase))
            {
                return ("PasswordRequired", "PasswordRequired", ex.Message);
            }

            return ("InvalidPassword", "InvalidPassword", ex.Message);
        }

        return ("OpenFailed", "OpenFailed", ex.Message);
    }

    private static bool IsHttpUrl(string value)
    {
        return Uri.TryCreate(value, UriKind.Absolute, out var uri)
            && (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
                || uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase));
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

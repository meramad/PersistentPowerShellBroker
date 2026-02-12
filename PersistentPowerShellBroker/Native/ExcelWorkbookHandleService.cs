namespace PersistentPowerShellBroker.Native;

internal sealed class ExcelWorkbookHandleService
{
    private readonly ExcelWorkbookLocator _locator = new();
    private readonly ExcelWorkbookOpener _opener = new();

    public AcquireWorkbookResult AcquireWorkbook(WorkbookIdentityResolution identity, AcquireWorkbookOptions options)
    {
        var requestedTarget = identity.InputIsUrl
            ? identity.NormalizedRemoteUrl ?? identity.RequestedInput
            : identity.NormalizedLocalPath;

        var existing = TryAttachFromRot(identity, options, requestedTarget);
        if (existing is not null)
        {
            return existing;
        }

        {
            var application = ExcelCommandSupport.CreateExcelApplication();
            if (application is null)
            {
                return AcquireWorkbookResult.Failure(
                    requestedTarget,
                    "ExcelNotInstalledOrCOMUnavailable",
                    "ExcelNotInstalledOrCOMUnavailable",
                    "Excel COM automation is unavailable.");
            }

            var createdSession = new ExcelApplicationSession(application, createdByBroker: true);
            var releaseCreatedSession = true;
            try
            {
                try
                {
                    createdSession.EnsureVisible(options.ForceVisible);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("CreateNew pre-open visibility enforcement failed.", ex);
                }

                if (!options.DisplayAlerts)
                {
                    try
                    {
                        createdSession.SetDisplayAlerts(enabled: false);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("CreateNew display alerts setup failed.", ex);
                    }
                }

                object? workbook;
                try
                {
                    workbook = _opener.OpenWorkbookWithRetry(
                        application,
                        requestedTarget,
                        options.ReadOnly,
                        options.OpenPassword,
                        options.ModifyPassword,
                        TimeSpan.FromSeconds(options.TimeoutSeconds));
                }
                catch (Exception ex)
                {
                    var attachedAfterFailure = TryAttachFromRot(identity, options, requestedTarget);
                    if (attachedAfterFailure is not null && attachedAfterFailure.Ok)
                    {
                        TryQuitCreatedApplication(createdSession, options.DisplayAlerts);
                        releaseCreatedSession = false;
                        return attachedAfterFailure;
                    }

                    throw new InvalidOperationException("WorkbookOpen failed.", ex);
                }

                if (workbook is null)
                {
                    return AcquireWorkbookResult.Failure(
                        requestedTarget,
                        "OpenFailed",
                        "OpenFailed",
                        "Excel did not return a workbook object.");
                }

                try
                {
                    createdSession.EnsureVisible(options.ForceVisible, workbook);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("CreateNew post-open visibility enforcement failed.", ex);
                }
                var workbookName = Convert.ToString(ExcelCommandSupport.GetProperty(workbook, "FullName")) ?? requestedTarget;
                var workbookReadOnly = Convert.ToBoolean(ExcelCommandSupport.GetProperty(workbook, "ReadOnly") ?? false);
                releaseCreatedSession = false;
                return AcquireWorkbookResult.Success(
                    requestedTarget,
                    workbookName,
                    createdSession,
                    workbook,
                    attachedExisting: false,
                    openedWorkbook: true,
                    workbookReadOnly,
                    workbookReadOnly ? "FileLockedOrWriteDenied" : null);
            }
            finally
            {
                if (releaseCreatedSession)
                {
                    createdSession.Release();
                }
            }
        }
    }

    private AcquireWorkbookResult? TryAttachFromRot(
        WorkbookIdentityResolution identity,
        AcquireWorkbookOptions options,
        string requestedTarget)
    {
        var runningApps = ExcelCommandSupport.EnumerateRunningExcelApplications();
        ExcelWorkbookLocator.Candidate? retainedCandidate = null;
        try
        {
            var match = _locator.FindOpenWorkbook(
                runningApps,
                identity,
                allowFileNameFallback: true);

            if (match.AmbiguousMatches.Count > 1)
            {
                return AcquireWorkbookResult.Failure(
                    requestedTarget,
                    "AmbiguousMatch",
                    "AmbiguousMatch",
                    $"Multiple workbook candidates: {string.Join(" | ", match.AmbiguousMatches)}");
            }

            if (match.MatchedCandidate is null)
            {
                return null;
            }

            retainedCandidate = match.MatchedCandidate;
            var session = new ExcelApplicationSession(match.MatchedCandidate.Application, createdByBroker: false);
            try
            {
                session.EnsureVisible(options.ForceVisible, match.MatchedCandidate.Workbook);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("AttachExisting visibility enforcement failed.", ex);
            }

            if (!options.DisplayAlerts)
            {
                try
                {
                    session.SetDisplayAlerts(enabled: false);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("AttachExisting display alerts setup failed.", ex);
                }
            }

            var workbookFullName = Convert.ToString(ExcelCommandSupport.GetProperty(match.MatchedCandidate.Workbook, "FullName")) ?? requestedTarget;
            var readOnly = Convert.ToBoolean(ExcelCommandSupport.GetProperty(match.MatchedCandidate.Workbook, "ReadOnly") ?? false);
            return AcquireWorkbookResult.Success(
                requestedTarget,
                workbookFullName,
                session,
                match.MatchedCandidate.Workbook,
                attachedExisting: true,
                openedWorkbook: false,
                readOnly,
                readOnly ? "FileLockedOrWriteDenied" : null);
        }
        finally
        {
            foreach (var app in runningApps)
            {
                if (retainedCandidate is not null && ReferenceEquals(app, retainedCandidate.Application))
                {
                    continue;
                }

                ExcelCommandSupport.SafeReleaseComObject(app);
            }
        }
    }

    private static void TryQuitCreatedApplication(ExcelApplicationSession createdSession, bool displayAlerts)
    {
        try
        {
            if (!displayAlerts)
            {
                createdSession.SetDisplayAlerts(enabled: false);
            }

            createdSession.Quit();
        }
        catch
        {
            // Best-effort cleanup for the temporary created application.
        }
        finally
        {
            createdSession.Release();
        }
    }
}

internal sealed record AcquireWorkbookOptions(
    bool ReadOnly,
    string? OpenPassword,
    string? ModifyPassword,
    int TimeoutSeconds,
    bool ForceVisible,
    bool DisplayAlerts);

internal sealed class AcquireWorkbookResult
{
    private AcquireWorkbookResult()
    {
    }

    public bool Ok { get; init; }
    public string Status { get; init; } = string.Empty;
    public string RequestedTarget { get; init; } = string.Empty;
    public string? WorkbookFullName { get; init; }
    public ExcelApplicationSession? Session { get; init; }
    public object? Workbook { get; init; }
    public bool AttachedExisting { get; init; }
    public bool OpenedWorkbook { get; init; }
    public bool? IsReadOnly { get; init; }
    public string? ReadOnlyReason { get; init; }
    public string? ErrorCode { get; init; }
    public string? ErrorMessage { get; init; }

    public static AcquireWorkbookResult Success(
        string requestedTarget,
        string workbookFullName,
        ExcelApplicationSession session,
        object workbook,
        bool attachedExisting,
        bool openedWorkbook,
        bool isReadOnly,
        string? readOnlyReason)
    {
        return new AcquireWorkbookResult
        {
            Ok = true,
            Status = "Success",
            RequestedTarget = requestedTarget,
            WorkbookFullName = workbookFullName,
            Session = session,
            Workbook = workbook,
            AttachedExisting = attachedExisting,
            OpenedWorkbook = openedWorkbook,
            IsReadOnly = isReadOnly,
            ReadOnlyReason = readOnlyReason,
            ErrorCode = null,
            ErrorMessage = null
        };
    }

    public static AcquireWorkbookResult Failure(string requestedTarget, string status, string errorCode, string errorMessage)
    {
        return new AcquireWorkbookResult
        {
            Ok = false,
            Status = status,
            RequestedTarget = requestedTarget,
            WorkbookFullName = null,
            Session = null,
            Workbook = null,
            AttachedExisting = false,
            OpenedWorkbook = false,
            IsReadOnly = null,
            ReadOnlyReason = null,
            ErrorCode = errorCode,
            ErrorMessage = errorMessage
        };
    }
}

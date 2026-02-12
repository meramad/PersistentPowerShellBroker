namespace PersistentPowerShellBroker.Native;

internal sealed class ExcelWorkbookLocator
{
    public LocatorResult FindOpenWorkbook(
        IReadOnlyList<object> applications,
        WorkbookIdentityResolution target,
        bool allowFileNameFallback)
    {
        var candidates = new List<Candidate>();
        var discoveryErrors = new List<string>();
        foreach (var app in applications)
        {
            dynamic? workbooks;
            try
            {
                workbooks = ((dynamic)app).Workbooks;
            }
            catch (Exception ex)
            {
                discoveryErrors.Add($"Failed to access Workbooks collection: {ex.Message}");
                continue;
            }

            if (workbooks is null)
            {
                continue;
            }

            try
            {
                int count;
                try
                {
                    count = Convert.ToInt32(workbooks.Count);
                }
                catch (Exception ex)
                {
                    discoveryErrors.Add($"Failed to read workbook count: {ex.Message}");
                    continue;
                }

                for (var i = 1; i <= count; i++)
                {
                    dynamic? workbook;
                    try
                    {
                        workbook = workbooks.Item(i);
                    }
                    catch (Exception ex)
                    {
                        discoveryErrors.Add($"Failed to access workbook index {i}: {ex.Message}");
                        continue;
                    }

                    if (workbook is null)
                    {
                        continue;
                    }

                    string fullName;
                    try
                    {
                        fullName = Convert.ToString(workbook.FullName) ?? string.Empty;
                    }
                    catch (Exception ex)
                    {
                        discoveryErrors.Add($"Failed to read workbook FullName at index {i}: {ex.Message}");
                        ExcelCommandSupport.SafeReleaseComObject(workbook);
                        continue;
                    }

                    candidates.Add(new Candidate(app, workbook, fullName));
                }
            }
            finally
            {
                ExcelCommandSupport.SafeReleaseComObject(workbooks);
            }
        }

        if (candidates.Count == 0 && discoveryErrors.Count > 0)
        {
            throw new InvalidOperationException("Excel workbook discovery failed: " + string.Join(" | ", discoveryErrors.Distinct()));
        }

        var match = Match(target, candidates, allowFileNameFallback);
        foreach (var candidate in candidates)
        {
            if (ReferenceEquals(candidate, match.MatchedCandidate))
            {
                continue;
            }

            ExcelCommandSupport.SafeReleaseComObject(candidate.Workbook);
        }

        return match;
    }

    internal static LocatorResult Match(
        WorkbookIdentityResolution target,
        IReadOnlyList<Candidate> candidates,
        bool allowFileNameFallback)
    {
        var exact = new List<Candidate>();
        var byName = new List<Candidate>();
        foreach (var candidate in candidates)
        {
            if (IsExactMatch(target, candidate))
            {
                exact.Add(candidate);
                continue;
            }

            if (string.Equals(candidate.FileName, target.FileName, StringComparison.OrdinalIgnoreCase))
            {
                byName.Add(candidate);
            }
        }

        if (exact.Count == 1)
        {
            return new LocatorResult(exact[0], []);
        }

        if (exact.Count > 1)
        {
            return new LocatorResult(null, exact.Select(static item => item.FullName).ToList());
        }

        if (!allowFileNameFallback)
        {
            return new LocatorResult(null, []);
        }

        if (byName.Count == 1)
        {
            return new LocatorResult(byName[0], []);
        }

        return new LocatorResult(null, byName.Select(static item => item.FullName).ToList());
    }

    private static bool IsExactMatch(WorkbookIdentityResolution target, Candidate candidate)
    {
        if (!string.IsNullOrWhiteSpace(target.NormalizedRemoteUrl)
            && !string.IsNullOrWhiteSpace(candidate.NormalizedRemoteUrl)
            && LocalWorkbookIdentityResolver.UrlsEqual(target.NormalizedRemoteUrl, candidate.NormalizedRemoteUrl))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(target.NormalizedLocalPath)
            && !string.IsNullOrWhiteSpace(candidate.NormalizedLocalPath)
            && string.Equals(target.NormalizedLocalPath, candidate.NormalizedLocalPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    internal sealed class Candidate
    {
        public Candidate(object application, object workbook, string fullName)
        {
            Application = application;
            Workbook = workbook;
            FullName = fullName;
            FileName = ExcelCommandSupport.FileNameFromTarget(fullName, IsHttpUrl(fullName));
            if (IsHttpUrl(fullName))
            {
                NormalizedRemoteUrl = LocalWorkbookIdentityResolver.NormalizeUrlForComparison(fullName);
                NormalizedLocalPath = string.Empty;
            }
            else
            {
                NormalizedLocalPath = LocalWorkbookIdentityResolver.NormalizeLocalPathForComparison(fullName);
                NormalizedRemoteUrl = null;
            }
        }

        public object Application { get; }
        public object Workbook { get; }
        public string FullName { get; }
        public string FileName { get; }
        public string NormalizedLocalPath { get; }
        public string? NormalizedRemoteUrl { get; }
    }

    internal sealed class LocatorResult
    {
        public LocatorResult(Candidate? matchedCandidate, List<string> ambiguousMatches)
        {
            MatchedCandidate = matchedCandidate;
            AmbiguousMatches = ambiguousMatches;
        }

        public Candidate? MatchedCandidate { get; }
        public List<string> AmbiguousMatches { get; }
    }

    private static bool IsHttpUrl(string value)
    {
        return Uri.TryCreate(value, UriKind.Absolute, out var uri)
            && (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
                || uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase));
    }
}

using PersistentPowerShellBroker.Native;

namespace PersistentPowerShellBroker.Tests;

public sealed class ExcelLocatorModuleTests
{
    [Fact]
    public void Resolver_NormalizesUrlInput()
    {
        var resolved = LocalWorkbookIdentityResolver.Resolve(
            "https://contoso.sharepoint.com/sites/A/Shared Documents/My File.xlsx");

        Assert.Equal(WorkbookIdentityResolutionStatus.UrlInput, resolved.Status);
        Assert.Equal("My File.xlsx", resolved.FileName);
        Assert.True(LocalWorkbookIdentityResolver.UrlsEqual(
            "https://contoso.sharepoint.com/sites/A/Shared%20Documents/My%20File.xlsx",
            resolved.NormalizedRemoteUrl));
    }

    [Fact]
    public void Resolver_NormalizesLocalPathInput()
    {
        var local = Path.Combine(Path.GetTempPath(), "psbroker-locator-test.xlsx");
        var resolved = LocalWorkbookIdentityResolver.Resolve(local);

        Assert.NotEqual(WorkbookIdentityResolutionStatus.UrlInput, resolved.Status);
        Assert.Equal(Path.GetFileName(local), resolved.FileName);
        Assert.Equal(LocalWorkbookIdentityResolver.NormalizeLocalPathForComparison(local), resolved.NormalizedLocalPath);
    }

    [Fact]
    public void Locator_MatchesByRemoteUrlBeforeFileName()
    {
        var target = new WorkbookIdentityResolution
        {
            RequestedInput = @"C:\sync\Signal Template.xlsx",
            NormalizedLocalPath = @"C:\sync\Signal Template.xlsx",
            NormalizedRemoteUrl = LocalWorkbookIdentityResolver.NormalizeUrlForComparison(
                "https://tkautomotive.sharepoint.com/teams/SystemDesignCore/Freigegebene Dokumente/System Models/Model Templates/Signal Template.xlsx"),
            FileName = "Signal Template.xlsx",
            Status = WorkbookIdentityResolutionStatus.MappedRemoteUrl
        };

        var candidates = new List<ExcelWorkbookLocator.Candidate>
        {
            new(new object(), new object(), "https://contoso.sharepoint.com/sites/A/Shared Documents/Signal Template.xlsx"),
            new(new object(), new object(), "https://tkautomotive.sharepoint.com/teams/SystemDesignCore/Freigegebene%20Dokumente/System%20Models/Model%20Templates/Signal%20Template.xlsx")
        };

        var match = ExcelWorkbookLocator.Match(target, candidates, allowFileNameFallback: true);

        Assert.NotNull(match.MatchedCandidate);
        Assert.Contains("tkautomotive.sharepoint.com", match.MatchedCandidate!.FullName, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(match.AmbiguousMatches);
    }

    [Fact]
    public void Locator_ReturnsAmbiguous_WhenOnlyFilenameMatchesMultipleCandidates()
    {
        var target = new WorkbookIdentityResolution
        {
            RequestedInput = @"C:\sync\Signal Template.xlsx",
            NormalizedLocalPath = @"C:\sync\Signal Template.xlsx",
            NormalizedRemoteUrl = null,
            FileName = "Signal Template.xlsx",
            Status = WorkbookIdentityResolutionStatus.LocalOnly
        };

        var candidates = new List<ExcelWorkbookLocator.Candidate>
        {
            new(new object(), new object(), "C:\\A\\Signal Template.xlsx"),
            new(new object(), new object(), "C:\\B\\Signal Template.xlsx")
        };

        var match = ExcelWorkbookLocator.Match(target, candidates, allowFileNameFallback: true);

        Assert.Null(match.MatchedCandidate);
        Assert.Equal(2, match.AmbiguousMatches.Count);
    }
}

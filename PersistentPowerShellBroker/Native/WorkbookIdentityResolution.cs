namespace PersistentPowerShellBroker.Native;

internal enum WorkbookIdentityResolutionStatus
{
    LocalOnly,
    MappedRemoteUrl,
    UrlInput
}

internal sealed class WorkbookIdentityResolution
{
    public required string RequestedInput { get; init; }
    public required string NormalizedLocalPath { get; init; }
    public string? NormalizedRemoteUrl { get; init; }
    public required string FileName { get; init; }
    public required WorkbookIdentityResolutionStatus Status { get; init; }
    public bool InputIsUrl => Status == WorkbookIdentityResolutionStatus.UrlInput;
}

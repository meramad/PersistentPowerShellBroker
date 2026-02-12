using System.Collections.Concurrent;

namespace PersistentPowerShellBroker.Native;

internal sealed record ExcelHandleMetadata(
    string VariableName,
    string RequestedTarget,
    string WorkbookFullName,
    bool AttachedExisting,
    bool OpenedWorkbook,
    bool IsReadOnly,
    bool CreatedApplicationByBroker,
    DateTime CreatedUtc);

internal static class ExcelHandleRegistry
{
    private static readonly ConcurrentDictionary<string, ExcelHandleMetadata> Handles = new(StringComparer.Ordinal);

    public static void Register(ExcelHandleMetadata metadata)
    {
        Handles[metadata.VariableName] = metadata;
    }

    public static bool TryGet(string variableName, out ExcelHandleMetadata metadata)
    {
        return Handles.TryGetValue(variableName, out metadata!);
    }

    public static void Remove(string variableName)
    {
        Handles.TryRemove(variableName, out _);
    }
}

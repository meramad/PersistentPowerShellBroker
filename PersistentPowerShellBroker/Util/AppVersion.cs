using System.Reflection;

namespace PersistentPowerShellBroker.Util;

public static class AppVersion
{
    public static string Value { get; } = Resolve();

    private static string Resolve()
    {
        var assembly = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informational))
        {
            var plusIndex = informational.IndexOf('+');
            return plusIndex >= 0 ? informational[..plusIndex] : informational;
        }

        return assembly.GetName().Version?.ToString() ?? "unknown";
    }
}

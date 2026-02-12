using Microsoft.Win32;

namespace PersistentPowerShellBroker.Native;

internal static class LocalWorkbookIdentityResolver
{
    private const string OneDriveProvidersRegistryPath = @"Software\SyncEngines\Providers\OneDrive";

    public static WorkbookIdentityResolution Resolve(string pathOrUrl)
    {
        if (Uri.TryCreate(pathOrUrl, UriKind.Absolute, out var uri)
            && (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
                || uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)))
        {
            var normalizedUrl = NormalizeUrlForComparison(pathOrUrl);
            return new WorkbookIdentityResolution
            {
                RequestedInput = pathOrUrl.Trim(),
                NormalizedLocalPath = string.Empty,
                NormalizedRemoteUrl = normalizedUrl,
                FileName = Path.GetFileName(uri.LocalPath),
                Status = WorkbookIdentityResolutionStatus.UrlInput
            };
        }

        var normalizedLocal = NormalizeLocalPathForComparison(pathOrUrl);
        var remote = TryResolveRemoteUrl(normalizedLocal);

        return new WorkbookIdentityResolution
        {
            RequestedInput = normalizedLocal,
            NormalizedLocalPath = normalizedLocal,
            NormalizedRemoteUrl = remote,
            FileName = Path.GetFileName(normalizedLocal),
            Status = remote is null ? WorkbookIdentityResolutionStatus.LocalOnly : WorkbookIdentityResolutionStatus.MappedRemoteUrl
        };
    }

    public static string NormalizeLocalPathForComparison(string path)
    {
        var normalized = Path.GetFullPath(path.Trim());
        if (normalized.Length <= 3)
        {
            return normalized;
        }

        return normalized.TrimEnd('\\', '/');
    }

    public static string NormalizeUrlForComparison(string url)
    {
        if (!Uri.TryCreate(url.Trim(), UriKind.Absolute, out var uri))
        {
            return url.Trim();
        }

        var rawSegments = uri.AbsolutePath
            .Split('/', StringSplitOptions.RemoveEmptyEntries);
        var normalizedSegments = rawSegments
            .Select(segment => Uri.EscapeDataString(Uri.UnescapeDataString(segment)))
            .ToArray();
        var normalizedPath = "/" + string.Join("/", normalizedSegments);
        if (normalizedPath.Length > 1 && normalizedPath.EndsWith("/", StringComparison.Ordinal))
        {
            normalizedPath = normalizedPath[..^1];
        }

        var authority = uri.IsDefaultPort
            ? $"{uri.Scheme.ToLowerInvariant()}://{uri.Host.ToLowerInvariant()}"
            : $"{uri.Scheme.ToLowerInvariant()}://{uri.Host.ToLowerInvariant()}:{uri.Port}";

        return $"{authority}{normalizedPath}";
    }

    public static bool UrlsEqual(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        return string.Equals(
            NormalizeUrlForComparison(left),
            NormalizeUrlForComparison(right),
            StringComparison.OrdinalIgnoreCase);
    }

    private static string? TryResolveRemoteUrl(string normalizedLocalPath)
    {
        using var providersKey = Registry.CurrentUser.OpenSubKey(OneDriveProvidersRegistryPath, writable: false);
        if (providersKey is null)
        {
            return null;
        }

        string? bestMount = null;
        string? bestUrlNamespace = null;
        foreach (var childName in providersKey.GetSubKeyNames())
        {
            using var child = providersKey.OpenSubKey(childName, writable: false);
            if (child is null)
            {
                continue;
            }

            var mountPoint = child.GetValue("MountPoint") as string;
            var urlNamespace = child.GetValue("UrlNamespace") as string;
            if (string.IsNullOrWhiteSpace(mountPoint) || string.IsNullOrWhiteSpace(urlNamespace))
            {
                continue;
            }

            var normalizedMount = NormalizeLocalPathForComparison(mountPoint);
            if (!IsPathUnderRoot(normalizedLocalPath, normalizedMount))
            {
                continue;
            }

            if (bestMount is null || normalizedMount.Length > bestMount.Length)
            {
                bestMount = normalizedMount;
                bestUrlNamespace = urlNamespace.Trim();
            }
        }

        if (bestMount is null || bestUrlNamespace is null)
        {
            return null;
        }

        var relative = normalizedLocalPath.Length == bestMount.Length
            ? string.Empty
            : normalizedLocalPath[(bestMount.Length + 1)..];
        var normalizedRelative = relative.Replace('\\', '/');
        var encodedRelative = string.Join(
            "/",
            normalizedRelative.Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(Uri.EscapeDataString));

        var baseUrl = bestUrlNamespace.TrimEnd('/');
        var candidate = string.IsNullOrWhiteSpace(encodedRelative)
            ? baseUrl
            : $"{baseUrl}/{encodedRelative}";
        return NormalizeUrlForComparison(candidate);
    }

    private static bool IsPathUnderRoot(string fullPath, string rootPath)
    {
        if (string.Equals(fullPath, rootPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return fullPath.Length > rootPath.Length
            && (fullPath[rootPath.Length] == '\\' || fullPath[rootPath.Length] == '/');
    }
}

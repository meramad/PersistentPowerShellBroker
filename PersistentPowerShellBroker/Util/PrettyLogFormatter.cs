using System.Text.Json;
using System.Text.Json.Nodes;
using PersistentPowerShellBroker.Protocol;

namespace PersistentPowerShellBroker.Util;

internal static class PrettyLogFormatter
{
    private static readonly HashSet<string> SensitiveKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "password",
        "openPassword",
        "modifyPassword",
        "token",
        "secret",
        "apiKey"
    };

    public static IReadOnlyList<string> Format(BrokerRequest request, BrokerResponse response)
    {
        var lines = new List<string>();
        AppendRequest(lines, request);
        AppendResponse(lines, response);
        lines.Add(string.Empty);
        return lines;
    }

    private static void AppendRequest(List<string> lines, BrokerRequest request)
    {
        var command = request.Command ?? string.Empty;
        var split = SplitLines(command);
        if (split.Length == 0)
        {
            lines.Add("> ");
            return;
        }

        lines.Add($"> {split[0]}");
        for (var i = 1; i < split.Length; i++)
        {
            lines.Add($"  {split[i]}");
        }
    }

    private static void AppendResponse(List<string> lines, BrokerResponse response)
    {
        if (!string.IsNullOrWhiteSpace(response.Stdout))
        {
            if (TryFormatJson(response.Stdout, out var prettyJson))
            {
                foreach (var line in SplitLines(prettyJson))
                {
                    lines.Add(line);
                }
            }
            else
            {
                foreach (var line in SplitLines(response.Stdout))
                {
                    lines.Add(line);
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(response.Stderr))
        {
            foreach (var line in SplitLines(response.Stderr))
            {
                lines.Add($"! {line}");
            }
        }

        if (!string.IsNullOrWhiteSpace(response.Error))
        {
            lines.Add($"ERROR: {response.Error}");
        }
    }

    private static bool TryFormatJson(string value, out string prettyJson)
    {
        prettyJson = string.Empty;
        try
        {
            var parsed = JsonNode.Parse(value);
            if (parsed is not JsonObject && parsed is not JsonArray)
            {
                return false;
            }

            var redacted = parsed!.DeepClone();
            RedactSensitive(redacted);
            prettyJson = redacted.ToJsonString(new JsonSerializerOptions
            {
                WriteIndented = true
            });
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void RedactSensitive(JsonNode? node)
    {
        switch (node)
        {
            case JsonObject obj:
                var keys = obj.Select(kvp => kvp.Key).ToList();
                foreach (var key in keys)
                {
                    if (SensitiveKeys.Contains(key))
                    {
                        obj[key] = "***";
                    }
                    else
                    {
                        RedactSensitive(obj[key]);
                    }
                }

                break;
            case JsonArray arr:
                foreach (var child in arr)
                {
                    RedactSensitive(child);
                }

                break;
        }
    }

    private static string[] SplitLines(string value)
    {
        return value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.None);
    }
}

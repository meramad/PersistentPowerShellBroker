using System.Management.Automation.Runspaces;
using System.Text;
using System.Text.Json;
using PersistentPowerShellBroker.Util;

namespace PersistentPowerShellBroker.Native;

public sealed class BrokerHelpCommand : INativeCommand
{
    public string Name => "broker.help";

    public Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext context, Runspace runspace, CancellationToken cancellationToken)
    {
        ExcelCommandSupport.TryGetString(args, "command", out var commandName);
        ExcelCommandSupport.TryGetString(args, "format", out var formatValue);
        var format = ParseFormat(formatValue);

        var broker = new
        {
            name = "PersistentPowerShellBroker",
            version = AppVersion.Value,
            build = (string?)null,
            pipeName = context.PipeName,
            startedUtc = context.StartedAtUtc.UtcDateTime.ToString("O")
        };

        var notes = new[]
        {
            "Use kind=\"native\" for broker commands and kind=\"powershell\" for free-form scripts.",
            "PowerShell state persists in one long-lived runspace across requests.",
            "One request is processed per connection and responses are single-line JSON.",
            "Use broker.excel.* handles to keep Excel COM objects alive safely in the broker runspace."
        };

        if (string.IsNullOrWhiteSpace(commandName))
        {
            var payload = new
            {
                ok = true,
                status = "Success",
                broker,
                nativeCommands = BrokerHelpCatalog.ListItems,
                commandHelp = (object?)null,
                notes,
                availableCommands = Array.Empty<string>()
            };

            return Task.FromResult(BuildResult(payload, format, success: true));
        }

        if (!BrokerHelpCatalog.TryGet(commandName, out var entry))
        {
            var payload = new
            {
                ok = false,
                status = "NotFound",
                broker,
                nativeCommands = Array.Empty<object>(),
                commandHelp = (object?)null,
                notes,
                availableCommands = BrokerHelpCatalog.AvailableCommands
            };

            return Task.FromResult(BuildResult(payload, format, success: false));
        }

        var commandPayload = new
        {
            name = entry.Name,
            summary = entry.Summary,
            @params = entry.Params.Select(parameter => new
            {
                name = parameter.Name,
                type = parameter.Type,
                required = parameter.Required,
                @default = parameter.DefaultValue,
                summary = parameter.Summary
            }).ToArray(),
            returns = entry.Returns.Select(item => new
            {
                name = item.Name,
                type = item.Type,
                summary = item.Summary
            }).ToArray(),
            examples = entry.Examples
        };

        var foundPayload = new
        {
            ok = true,
            status = "Success",
            broker,
            nativeCommands = Array.Empty<object>(),
            commandHelp = commandPayload,
            notes,
            availableCommands = Array.Empty<string>()
        };

        return Task.FromResult(BuildResult(foundPayload, format, success: true));
    }

    private static NativeResult BuildResult<T>(T payload, HelpFormat format, bool success)
    {
        var json = JsonSerializer.Serialize(payload);
        if (format == HelpFormat.Text)
        {
            return new NativeResult
            {
                Success = success,
                Stdout = BuildText(payload),
                Stderr = string.Empty,
                Error = null
            };
        }

        if (format == HelpFormat.Both)
        {
            return new NativeResult
            {
                Success = success,
                Stdout = $"{json}{Environment.NewLine}{Environment.NewLine}{BuildText(payload)}",
                Stderr = string.Empty,
                Error = null
            };
        }

        return new NativeResult
        {
            Success = success,
            Stdout = json,
            Stderr = string.Empty,
            Error = null
        };
    }

    private static string BuildText<T>(T payload)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(payload));
        var root = doc.RootElement;
        var sb = new StringBuilder();

        var broker = root.GetProperty("broker");
        sb.AppendLine($"Broker: {broker.GetProperty("name").GetString()} v{broker.GetProperty("version").GetString()}");
        sb.AppendLine($"Pipe: \\\\.\\pipe\\{broker.GetProperty("pipeName").GetString()}");
        sb.AppendLine($"Started: {broker.GetProperty("startedUtc").GetString()}");
        sb.AppendLine();

        if (root.TryGetProperty("commandHelp", out var commandHelp) && commandHelp.ValueKind == JsonValueKind.Object)
        {
            sb.AppendLine($"Command: {commandHelp.GetProperty("name").GetString()}");
            sb.AppendLine(commandHelp.GetProperty("summary").GetString());
        }
        else if (root.TryGetProperty("nativeCommands", out var nativeCommands) && nativeCommands.ValueKind == JsonValueKind.Array)
        {
            sb.AppendLine("Native commands:");
            foreach (var item in nativeCommands.EnumerateArray())
            {
                sb.AppendLine($"- {item.GetProperty("name").GetString()}: {item.GetProperty("summary").GetString()}");
            }
        }

        sb.AppendLine();
        sb.Append("Hint: free-form PowerShell is available via kind=\"powershell\" and state persists.");
        return sb.ToString();
    }

    private static HelpFormat ParseFormat(string? raw)
    {
        if (string.Equals(raw, "Text", StringComparison.OrdinalIgnoreCase))
        {
            return HelpFormat.Text;
        }

        if (string.Equals(raw, "Both", StringComparison.OrdinalIgnoreCase))
        {
            return HelpFormat.Both;
        }

        return HelpFormat.Json;
    }

    private enum HelpFormat
    {
        Json,
        Text,
        Both
    }
}

internal static class BrokerHelpCatalog
{
    private static readonly IReadOnlyList<HelpEntry> Entries =
    [
        new(
            "broker.help",
            "Returns native command index or detailed schema for one command.",
            [
                new("command", "string", false, null, "Exact native command name for detailed help."),
                new("format", "enum(Json|Text|Both)", false, "Json", "Output format preference.")
            ],
            [
                new("ok", "bool", "True on success, false when requested command is not found."),
                new("status", "enum(Success|NotFound)", "Result status code."),
                new("broker", "object", "Broker identity and runtime context."),
                new("nativeCommands", "array", "Native command list when command is omitted."),
                new("commandHelp", "object|null", "Detailed schema when command is provided."),
                new("notes", "array", "Operational usage tips."),
                new("availableCommands", "array", "Available commands when status is NotFound.")
            ],
            [
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.help'",
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.help' -Raw",
                "$req=@{id='1';kind='native';command='broker.help';args=@{command='broker.info'}}|ConvertTo-Json -Compress"
            ]),
        new(
            "broker.info",
            "Returns broker runtime metadata including version, pipe and process id.",
            [],
            [
                new("version", "string", "Broker version."),
                new("pipeName", "string", "Named pipe identifier."),
                new("startedAtUtc", "string", "Broker start time."),
                new("pid", "int", "Broker process id.")
            ],
            [
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.info'"
            ]),
        new(
            "broker.stop",
            "Requests graceful broker shutdown after replying to the caller.",
            [],
            [
                new("stdout", "string", "Returns 'stopping'.")
            ],
            [
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.stop'"
            ]),
        new(
            "broker.excel.get_workbook_handle",
            "Finds or opens an Excel workbook and stores a global handle bundle.",
            [
                new("path", "string", true, null, "Local workbook path or SharePoint/OneDrive URL."),
                new("readOnly", "bool", false, false, "Open preference for read-only mode."),
                new("openPassword", "string|null", false, null, "Open password if required."),
                new("modifyPassword", "string|null", false, null, "Password to modify if required."),
                new("timeoutSeconds", "int", false, 15, "Open timeout guard."),
                new("instancePolicy", "enum(ReuseIfRunning|AlwaysNew)", false, "ReuseIfRunning", "Excel instance reuse policy."),
                new("displayAlerts", "bool", false, false, "Set Excel DisplayAlerts before open."),
                new("forceVisible", "bool", false, true, "Set Excel Visible before return.")
            ],
            [
                new("ok", "bool", "Command success indicator."),
                new("status", "string", "Outcome status such as Success, FileNotFound, TimeoutLikelyModalDialog."),
                new("psVariableName", "string|null", "Global handle variable name when successful."),
                new("workbookFullName", "string|null", "Resolved workbook fullname."),
                new("requestedTarget", "string", "Normalized requested target."),
                new("attachedExisting", "bool", "True when attached to an already-open workbook."),
                new("openedWorkbook", "bool", "True when opened during this call.")
            ],
            [
                "$req=@{id='1';kind='native';command='broker.excel.get_workbook_handle';args=@{path='C:\\\\Temp\\\\Book1.xlsx'}}|ConvertTo-Json -Compress",
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.excel.get_workbook_handle' -Raw"
            ]),
        new(
            "broker.excel.release_handle",
            "Releases an Excel handle bundle and optionally closes workbook / quits Excel.",
            [
                new("psVariableName", "string", true, null, "Handle variable name returned earlier."),
                new("closeWorkbook", "bool", false, false, "Close workbook before releasing."),
                new("saveChanges", "bool|null", false, null, "Save behavior when closing workbook."),
                new("quitExcel", "bool", false, false, "Attempt Excel quit after close."),
                new("onlyIfNoOtherWorkbooks", "bool", false, true, "Skip quit when other workbooks are open."),
                new("timeoutSeconds", "int", false, 10, "Close and quit timeout guard."),
                new("displayAlerts", "bool", false, false, "Set DisplayAlerts during close/quit.")
            ],
            [
                new("ok", "bool", "Command success indicator."),
                new("status", "string", "Outcome status such as Success, NotFound, CloseFailed."),
                new("closedWorkbook", "bool", "Whether workbook close was completed."),
                new("quitExcelAttempted", "bool", "Whether quit flow was attempted."),
                new("quitExcelSucceeded", "bool", "Whether quit succeeded."),
                new("released", "bool", "Whether references were released and variable removed.")
            ],
            [
                "$req=@{id='2';kind='native';command='broker.excel.release_handle';args=@{psVariableName='excelHandle_abc'}}|ConvertTo-Json -Compress",
                "Invoke-PSBroker -Pipe $pipe -Kind native -Command 'broker.excel.release_handle' -Raw"
            ])
    ];

    public static IReadOnlyList<string> AvailableCommands { get; } = Entries.Select(entry => entry.Name).ToArray();

    public static IReadOnlyList<object> ListItems { get; } = Entries
        .Select(entry => new
        {
            name = entry.Name,
            summary = entry.Summary,
            @params = entry.Params.Select(parameter => new
            {
                name = parameter.Name,
                type = parameter.Type,
                required = parameter.Required,
                @default = parameter.DefaultValue,
                summary = parameter.Summary
            }).ToArray()
        })
        .Cast<object>()
        .ToArray();

    public static bool TryGet(string name, out HelpEntry entry)
    {
        var found = Entries.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        if (found is null)
        {
            entry = null!;
            return false;
        }

        entry = found;
        return true;
    }

    internal sealed record HelpEntry(
        string Name,
        string Summary,
        IReadOnlyList<HelpParam> Params,
        IReadOnlyList<HelpReturn> Returns,
        IReadOnlyList<string> Examples);

    internal sealed record HelpParam(string Name, string Type, bool Required, object? DefaultValue, string Summary);

    internal sealed record HelpReturn(string Name, string Type, string Summary);
}

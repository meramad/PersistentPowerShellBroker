# PersistentPowerShellBroker

PersistentPowerShellBroker is a local Windows broker that hosts one persistent PowerShell runspace behind a named pipe.
It allows multiple clients to execute commands in the same shared session state safely through a single execution queue.

Owner: Miklós Arányi

This project was developed with substantial assistance from generative AI tools. Final design, integration, and release decisions are owned by Miklós Arányi.

## Why this exists

PowerShell state usually resets between independent invocations. This broker keeps one runspace alive so state persists across requests:

- variables
- current directory
- loaded modules/functions

## Example usage scenario

You want an automation client to set state once and reuse it:

1. Client A sets `$global:projectRoot` and loads helper functions.
2. Client B later calls commands that rely on that same state.
3. Both requests are serialized, so shared runspace state stays consistent.

## Quick start

1. Build:

```powershell
dotnet build .\PersistentPowerShellBroker.sln -c Release
```

2. Run broker:

```powershell
dotnet run --project .\PersistentPowerShellBroker\PersistentPowerShellBroker.csproj -- --pipe auto --log-option silent
```

3. Read startup output and capture:

```text
PIPE=\\.\pipe\psbroker-...
```

4. Use the helper client:

```powershell
. .\client\Invoke-PSBroker.ps1
Invoke-PSBroker -Pipe "\\.\pipe\psbroker-..." -Command "Get-Date"
Invoke-PSBroker -Pipe "\\.\pipe\psbroker-..." -Command '$global:x=5'
Invoke-PSBroker -Pipe "\\.\pipe\psbroker-..." -Command '$global:x*10'
Invoke-PSBroker -Pipe "\\.\pipe\psbroker-..." -Kind native -Command "broker.stop"
```

## Release build (single EXE)

```powershell
dotnet publish .\PersistentPowerShellBroker\PersistentPowerShellBroker.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeAllContentForSelfExtract=true /p:PublishTrimmed=false /p:DebugType=None /p:DebugSymbols=false -o .\publish
```

Output:

- `publish\PersistentPowerShellBroker.exe`

## Download release

After a GitHub Release is created, users can download assets from:

- `https://github.com/<your-user-or-org>/PersistentPowerShellBroker/releases/latest`

## License

Licensed under the MIT License. See `LICENSE`.

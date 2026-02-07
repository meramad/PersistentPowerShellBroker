# PersistentPowerShellBroker – Codex Build Spec (v1)

## Goal
Build a standalone Windows **.NET 8 console EXE** that hosts a **single persistent PowerShell Runspace** and exposes it over a **local Named Pipe** so multiple clients (you + Codex) can execute commands in the **same shared session state**.

This is **not WinRM/PSRemoting**. No admin rights required.

---

## Non-goals (v1)
- No Excel/COM automation yet (but architecture must be compatible later).
- No interactive typing into the broker console (logging only).
- No multi-command per connection (exactly **1 request per connection**).
- No remote/network access (localhost only via named pipes).

---

## High-level design

### Components
1. **BrokerHost**
   - Owns the single shared **PowerShell Runspace**.
   - Owns an **execution queue** so all requests execute sequentially in the same runspace.
   - Provides an API: `ExecutePowerShell(command) -> (stdout, stderr, success)` and `ExecuteNative(name,args)`.

2. **PipeServer**
   - Listens on a Named Pipe (Windows) with **ACL restricted to current user**.
   - Connection lifecycle: client connects → sends 1 JSON line → receives 1 JSON line → disconnect.
   - Supports multiple concurrent connections by accepting connections concurrently, but pushes work into the single execution queue.

3. **Protocol**
   - JSON line protocol (newline-delimited UTF-8).
   - Request + Response objects.

4. **Native command framework**
   - Registry of native commands (`broker.info`, `broker.stop`) with placeholders for future commands.

5. **Console logging**
   - Broker prints startup banner and logs requests/results to console.
   - Logs must **never** be written to the client pipe stream.

---

## Startup & CLI

### Mandatory args
- `--pipe auto|<name>`
  - `auto` generates a random name: `psbroker-<guid>`
  - `<name>` uses provided name directly

### Optional args
- `--log-option silent|info|debug` (default: `silent`)
- `--init <path>` (optional) dot-sources a PowerShell script into the runspace on startup
- `--idle-exit-minutes <n>` (optional; default: disabled)

### Startup output
Broker prints a startup banner and must always print one machine-readable pipe line:
- `PIPE=\\.\pipe\<pipeName>` (always printed, including `silent`)

Exit codes:
- `0` normal stop
- `2` invalid args
- `3` fatal startup error
- `4` pipe fatal error

---

## Security

### Named pipe access control (mandatory)
- Pipe must be created with a security descriptor allowing **only the current Windows user** (and optionally local Administrators) to connect.
- No tokens required in v1.
- Pipe name randomness (auto mode) reduces collisions.

Implementation hint: use `System.IO.Pipes.NamedPipeServerStream` overload that accepts `PipeSecurity`.

---

## Protocol (JSON line)

### Request schema (single JSON line)
Fields:
- `id` *(string, required)*: correlation id (GUID recommended)
- `kind` *(string, required)*: `"powershell"` or `"native"`
- `command` *(string, required)*:
  - if `powershell`: script text
  - if `native`: native command name (e.g., `broker.info`)
- `args` *(object, optional)*: native args (reserved)
- `timeoutMs` *(int, optional)*: reserved for v2; in v1 ignore or log
- `clientName` *(string, optional)*: caller identity for broker logging
- `clientPid` *(int, optional)*: caller process id for broker logging

Example:
```json
{"id":"b1","kind":"powershell","command":"Get-Date"}
```

### Response schema (single JSON line)
Fields:
- `id` *(string)*: echoes request id
- `success` *(bool)*
- `stdout` *(string)*
- `stderr` *(string)*
- `error` *(string|null)*: broker-level error (not PowerShell errors)
- `durationMs` *(int)*

Example:
```json
{"id":"b1","success":true,"stdout":"...","stderr":"","error":null,"durationMs":12}
```

### Output capture rules
- PowerShell pipeline output (`Write-Output`, emitted objects) -> `stdout` (string, one joined)
- PowerShell error stream -> `stderr`
- Terminating errors set `success=false` and should be reflected in `stderr` (and optionally `error` only for broker faults)

---

## PowerShell Hosting Requirements

### Runspace creation
- Use `InitialSessionState.CreateDefault()`
- Create a single runspace and keep it alive for broker lifetime.

### Apartment state (future Excel compatibility)
- The execution worker thread must run in **STA** (Single-Threaded Apartment).
  - In v1, enforce that the runspace execution loop runs on an STA thread.

### State persistence
Because it is the same runspace:
- current directory persists
- variables persist
- modules/dot-sourced functions persist

---

## Execution model

### Queueing
- All commands run sequentially using a single worker.
- Multiple clients can connect simultaneously; their requests are queued.

### One request per connection
- Client connects, sends 1 request, waits for response, disconnects.
- Broker must close connection after responding.

### Timeouts & cancellation (v1)
- Not implemented; ignore `timeoutMs` for now.
- Structure code so that v2 can add timeouts/cancellation.

---

## Native Commands (v1 minimal)

### Interface
`INativeCommand`:
- `string Name { get; }`
- `Task<NativeResult> ExecuteAsync(JsonElement? args, BrokerContext ctx, CancellationToken ct)`

`NativeResult`:
- `bool Success`
- `string Stdout`
- `string Stderr`
- `string? Error`

### Commands
1. `broker.info`
   - Returns version, pipe name, start time, pid.
2. `broker.stop`
   - Responds success, then triggers graceful shutdown.

---

## Logging
- Startup:
  - Broker prints startup banner.
  - Broker always prints `PIPE=\\.\pipe\<pipeName>`.
- `silent`:
  - No per-request/per-response logs.
- `info`:
  - One line per request:
    - `client=<clientName or ?> pid=<clientPid or ?> request=<id> kind=<kind> success=<bool> durationMs=<n> command="<preview>"`
- `debug`:
  - Multi-line block per request.
  - Includes all `info` fields plus:
    - `stdoutPreview="<first N chars>"`
    - `stderrPreview="<first N chars>"`
  - Previews must be truncated (e.g., `N=1000`) and append `...(truncated)` when truncated.
  - Never print full unbounded stdout/stderr.

---

## Project Structure

Solution name: `PersistentPowerShellBroker`  
Project: `PersistentPowerShellBroker` (Console App, .NET 8)

Suggested folders/files:
```
src/PersistentPowerShellBroker/
  Program.cs
  BrokerHost.cs
  PipeServer.cs
  Protocol/
    Request.cs
    Response.cs
    JsonLineCodec.cs
  Native/
    INativeCommand.cs
    NativeRegistry.cs
    BrokerInfoCommand.cs
    BrokerStopCommand.cs
  Util/
    ConsoleLogger.cs
    StopSignal.cs
```

NuGet:
- `Microsoft.PowerShell.SDK`

---

## Client Helper (PowerShell)
Create `client/Invoke-PSBroker.ps1` with a function:
- Parameters:
  - `-Pipe "\\.\pipe\psbroker-..."` (or a `-PipeName`)
  - `-Command <string>`
  - `-ClientName <string>` (default: `powershell`)
  - `-ClientPid <int>` (default: current `$PID`)
  - `-Raw` (optional: return full response object)
- Default behavior prints only `stdout`.

---

## Test Plan (v1)

1) **Basic execution**
- `Get-Date` returns non-empty stdout

2) **State persistence**
- `Set-Location $env:TEMP` then `pwd` reflects temp
- `$global:x=5` then `$global:x*10` returns 50

3) **Multi-client queueing**
- Client A: `Start-Sleep 2; "A"`
- Client B: `"B"`
- B should return after A completes (queued)

4) **Native stop**
- Send request: `{kind:"native", command:"broker.stop"}`
- Broker exits cleanly with exit code 0

5) **ACL sanity**
- Verify pipe is not accessible from other users (best-effort; document manual check).

---

## Implementation Notes / Guardrails
- Never write logs into pipe stream.
- Ensure JSON is exactly one line per request/response.
- Ensure output conversion is stable: convert pipeline output objects to strings (e.g., `Out-String` equivalent).
- Keep the core host extensible for future native commands and COM objects (Excel).

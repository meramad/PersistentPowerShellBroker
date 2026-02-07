# AGENTS.md – Guidance for Codex

## Repo
PersistentPowerShellBroker (Windows .NET 8 console application)
Read specification in PERSISTENT_POWERSHELL_BROKER_SPEC.md

## Primary goal (v1)
Implement a local Named Pipe broker that hosts a **single persistent PowerShell Runspace** and executes commands in that runspace with session persistence.

## Constraints
- Must run without admin rights.
- Must use Named Pipes (no TCP).
- Must restrict pipe connections to the current user (PipeSecurity ACL).
- One request per connection:
  - client connects -> sends ONE JSON line -> receives ONE JSON line -> disconnect.
- Multiple clients supported via concurrent accept + single execution queue.
- Broker logs to console only (never into client response stream).
- JSON protocol exactly as described in PERSISTENT_POWERSHELL_BROKER_SPEC.md

## Build order (recommended)
1. Implement Protocol DTOs and JsonLineCodec (read/write one JSON line).
2. Implement Native command framework + broker.info and broker.stop.
3. Implement BrokerHost (runspace + sequential execution queue on STA thread).
4. Implement PipeServer (NamedPipeServerStream with PipeSecurity, 1 req/conn).
5. Implement Program.cs CLI parsing and startup printing:
   - Print: PIPE=\\.\pipe\<name>
6. Add minimal PowerShell client helper script in /client.

## Coding guidelines
- Use async/await for pipe accept and per-connection handling.
- Use a single worker loop to execute commands sequentially.
- Convert PowerShell output to text deterministically (Out-String behavior).
- Keep code simple and testable; prefer small classes with clear responsibilities.
- Add basic unit tests only if fast; otherwise provide a manual test script in /client.

## Definition of Done (v1)
- Running broker prints PIPE=... line.
- Client can send PowerShell commands and get stdout/stderr JSON response.
- Runspace state persists across calls.
- Multiple simultaneous clients do not corrupt state; requests are queued.
- broker.stop stops the process gracefully after responding.

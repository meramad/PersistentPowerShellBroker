function Invoke-PSBroker {
    [CmdletBinding(DefaultParameterSetName = "ByName")]
    param(
        [Parameter(ParameterSetName = "ByPath", Mandatory = $true)]
        [string]$Pipe,

        [Parameter(ParameterSetName = "ByName")]
        [string]$PipeName,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [switch]$AutoDiscoverPipe,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [string]$Command,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [AllowNull()]
        [object]$Args,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [string]$Script,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [ValidateSet("Native", "PowerShell", "native", "powershell")]
        [string]$Kind,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 30,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [switch]$PassThru,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [switch]$Raw,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [string]$ClientName = "powershell",

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [int]$ClientPid = $PID,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [string]$BrokerStdOutPath,

        [Parameter(ParameterSetName = "ByPath")]
        [Parameter(ParameterSetName = "ByName")]
        [string[]]$KnownPipeNames
    )

    function Get-PipeNameFromInput {
        if (-not [string]::IsNullOrWhiteSpace($PipeName)) {
            return $PipeName.Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($Pipe)) {
            if ($Pipe.StartsWith("\\.\pipe\")) {
                return $Pipe.Substring("\\.\pipe\".Length)
            }

            throw "Pipe must start with \\.\pipe\."
        }

        return $null
    }

    function Invoke-PipeRequest {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ResolvedPipeName,
            [Parameter(Mandatory = $true)]
            [string]$RequestJson
        )

        $timeoutMs = [Math]::Max(1000, $TimeoutSeconds * 1000)
        $client = [System.IO.Pipes.NamedPipeClientStream]::new(".", $ResolvedPipeName, [System.IO.Pipes.PipeDirection]::InOut)
        try {
            $client.Connect($timeoutMs)
            $writer = [System.IO.StreamWriter]::new($client, [System.Text.UTF8Encoding]::new($false), 1024, $true)
            $reader = [System.IO.StreamReader]::new($client, [System.Text.UTF8Encoding]::new($false), $false, 1024, $true)
            try {
                $writeTask = $writer.WriteLineAsync($RequestJson)
                if (-not $writeTask.Wait($timeoutMs)) {
                    throw "Broker write timed out after $TimeoutSeconds seconds."
                }

                $flushTask = $writer.FlushAsync()
                if (-not $flushTask.Wait($timeoutMs)) {
                    throw "Broker flush timed out after $TimeoutSeconds seconds."
                }

                $readTask = $reader.ReadLineAsync()
                if (-not $readTask.Wait($timeoutMs)) {
                    throw "Broker read timed out after $TimeoutSeconds seconds."
                }

                return [string]$readTask.Result
            }
            finally {
                $writer.Dispose()
                $reader.Dispose()
            }
        }
        finally {
            $client.Dispose()
        }
    }

    function Parse-BrokerResponse {
        param([string]$ResponseJson)

        if ([string]::IsNullOrWhiteSpace($ResponseJson)) {
            throw "Broker returned no response line."
        }

        $envelope = $ResponseJson | ConvertFrom-Json
        $payload = $null
        if (-not [string]::IsNullOrWhiteSpace($envelope.stdout)) {
            try {
                $payload = $envelope.stdout | ConvertFrom-Json
            }
            catch {
                $payload = $null
            }
        }

        $result = [pscustomobject]@{
            id = $envelope.id
            success = [bool]$envelope.success
            stdout = [string]$envelope.stdout
            stderr = [string]$envelope.stderr
            error = [string]$envelope.error
            durationMs = $envelope.durationMs
            payload = $payload
            raw = $ResponseJson
        }
        return $result
    }

    function Get-DiscoveredPipeName {
        $candidateNames = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrWhiteSpace($BrokerStdOutPath) -and (Test-Path -LiteralPath $BrokerStdOutPath)) {
            $lines = Get-Content -LiteralPath $BrokerStdOutPath -ErrorAction SilentlyContinue
            foreach ($line in ($lines | Select-Object -Last 200)) {
                if ($line -match "^PIPE=\\\\\.\\pipe\\(?<name>.+)$") {
                    $candidateNames.Add($Matches["name"])
                }
            }
        }

        if ($KnownPipeNames) {
            foreach ($known in $KnownPipeNames) {
                if (-not [string]::IsNullOrWhiteSpace($known)) {
                    $candidateNames.Add($known.Trim())
                }
            }
        }

        try {
            $pipeItems = Get-ChildItem -Path "\\.\pipe\" -ErrorAction SilentlyContinue
            foreach ($item in $pipeItems) {
                if ($item.Name -like "psbroker*" -or $item.Name -like "PersistentPowerShellBroker*") {
                    $candidateNames.Add($item.Name)
                }
            }
        }
        catch {
            # best effort
        }

        $unique = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
        $probeRequest = [ordered]@{
            id = [guid]::NewGuid().ToString("N")
            kind = "native"
            command = "broker.info"
            clientName = $ClientName
            clientPid = $ClientPid
        } | ConvertTo-Json -Compress -Depth 20

        foreach ($candidate in $candidateNames) {
            if ([string]::IsNullOrWhiteSpace($candidate)) {
                continue
            }

            if (-not $unique.Add($candidate)) {
                continue
            }

            try {
                $rawLine = Invoke-PipeRequest -ResolvedPipeName $candidate -RequestJson $probeRequest
                $probe = Parse-BrokerResponse -ResponseJson $rawLine
                if ($probe.success) {
                    return $candidate
                }
            }
            catch {
                continue
            }
        }

        throw "Unable to auto-discover broker pipe."
    }

    $hasCommand = -not [string]::IsNullOrWhiteSpace($Command)
    $hasScript = -not [string]::IsNullOrWhiteSpace($Script)
    if (($hasCommand -and $hasScript) -or (-not $hasCommand -and -not $hasScript)) {
        throw "Exactly one of -Command or -Script must be provided."
    }

    if ($Args -and -not $hasCommand) {
        throw "-Args can only be used with -Command."
    }

    $resolvedPipeName = Get-PipeNameFromInput
    if ([string]::IsNullOrWhiteSpace($resolvedPipeName)) {
        if (-not $AutoDiscoverPipe) {
            throw "Provide -PipeName/-Pipe or set -AutoDiscoverPipe."
        }

        $resolvedPipeName = Get-DiscoveredPipeName
    }

    $normalizedKind = $null
    if (-not [string]::IsNullOrWhiteSpace($Kind)) {
        $normalizedKind = if ($Kind.Equals("Native", [System.StringComparison]::OrdinalIgnoreCase)) { "native" } else { "powershell" }
    }
    elseif ($hasScript) {
        $normalizedKind = "powershell"
    }
    elseif ($Command.StartsWith("broker.", [System.StringComparison]::OrdinalIgnoreCase)) {
        $normalizedKind = "native"
    }
    else {
        $normalizedKind = "powershell"
    }

    $request = [ordered]@{
        id = [guid]::NewGuid().ToString("N")
        kind = $normalizedKind
        command = $(if ($hasScript) { $Script } else { $Command })
        clientName = $ClientName
        clientPid = $ClientPid
    }
    if ($hasCommand -and $null -ne $Args) {
        $request.args = $Args
    }

    $requestJson = $request | ConvertTo-Json -Compress -Depth 50

    try {
        $rawLine = Invoke-PipeRequest -ResolvedPipeName $resolvedPipeName -RequestJson $requestJson
        if ($Raw) {
            return $rawLine
        }

        $parsed = Parse-BrokerResponse -ResponseJson $rawLine
        if (-not $PassThru -and -not $parsed.success) {
            $failure = if (-not [string]::IsNullOrWhiteSpace($parsed.error)) { $parsed.error } else { "Broker request failed." }
            throw $failure
        }

        return $parsed
    }
    catch {
        if ($PassThru) {
            return [pscustomobject]@{
                id = $null
                success = $false
                stdout = ""
                stderr = ""
                error = $_.Exception.Message
                durationMs = $null
                payload = $null
                raw = $null
            }
        }

        throw
    }
}

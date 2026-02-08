param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [string]$BrokerExePath,

    [int]$TimeoutSeconds = 30,

    [string]$PipeName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:BrokerProcess = $null
$script:ClientProcess = $null
$script:BrokerStdOutPath = $null
$script:BrokerStdErrPath = $null
$script:ClientStdOutPath = $null
$script:ClientStdErrPath = $null
$script:ClientScriptPath = $null

function Write-Pass {
    param([string]$Name)
    Write-Output "PASS $Name"
}

function Write-Fail {
    param([string]$Name, [string]$Reason)
    [Console]::Error.WriteLine(("FAIL {0}: {1}" -f $Name, $Reason))
}

function Get-LogTail {
    param(
        [string]$Path,
        [int]$Lines = 50
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return "<no log>"
    }

    $content = Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $content) {
        return "<empty>"
    }

    return ($content | Select-Object -Last $Lines) -join [Environment]::NewLine
}

function Resolve-RunnerPwsh {
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwsh) {
        return $pwsh.Source
    }

    $powershell = Get-Command powershell -ErrorAction SilentlyContinue
    if ($powershell) {
        return $powershell.Source
    }

    throw "No PowerShell executable found (pwsh/powershell)."
}

function Wait-BrokerReady {
    param(
        [string]$HelperPath,
        [string]$Pipe,
        [int]$TimeoutSec
    )

    . $HelperPath
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if ($script:BrokerProcess.HasExited) {
            return $false
        }

        try {
            $response = Invoke-PSBroker -PipeName $Pipe -Kind native -Command "broker.info" -Raw
            if ($response -and $response.success) {
                return $true
            }
        }
        catch {
            Start-Sleep -Milliseconds 250
            continue
        }

        Start-Sleep -Milliseconds 250
    }

    return $false
}

function Cleanup-Processes {
    if ($script:ClientProcess -and -not $script:ClientProcess.HasExited) {
        Stop-Process -Id $script:ClientProcess.Id -Force -ErrorAction SilentlyContinue
    }

    if ($script:BrokerProcess -and -not $script:BrokerProcess.HasExited) {
        Stop-Process -Id $script:BrokerProcess.Id -Force -ErrorAction SilentlyContinue
    }
}

function Remove-TempFiles {
    $paths = @(
        $script:BrokerStdOutPath,
        $script:BrokerStdErrPath,
        $script:ClientStdOutPath,
        $script:ClientStdErrPath,
        $script:ClientScriptPath
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        }
    }
}

function Resolve-BrokerLaunch {
    param(
        [string]$RepoRoot,
        [string]$Config,
        [string]$PathOverride
    )

    if (-not [string]::IsNullOrWhiteSpace($PathOverride)) {
        $fullOverride = [System.IO.Path]::GetFullPath($PathOverride)
        if (-not (Test-Path -LiteralPath $fullOverride)) {
            throw "Broker executable override does not exist: $fullOverride"
        }

        if ($fullOverride.EndsWith(".dll", [System.StringComparison]::OrdinalIgnoreCase)) {
            return @{
                FilePath = "dotnet"
                Arguments = @($fullOverride)
                ResolvedPath = $fullOverride
            }
        }

        return @{
            FilePath = $fullOverride
            Arguments = @()
            ResolvedPath = $fullOverride
        }
    }

    & dotnet build (Join-Path $RepoRoot "PersistentPowerShellBroker.sln") -c $Config | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed with exit code $LASTEXITCODE."
    }

    $baseOut = Join-Path $RepoRoot "PersistentPowerShellBroker\bin\$Config\net8.0"
    $exe = Join-Path $baseOut "PersistentPowerShellBroker.exe"
    if (Test-Path -LiteralPath $exe) {
        return @{
            FilePath = $exe
            Arguments = @()
            ResolvedPath = $exe
        }
    }

    $dll = Join-Path $baseOut "PersistentPowerShellBroker.dll"
    if (Test-Path -LiteralPath $dll) {
        return @{
            FilePath = "dotnet"
            Arguments = @($dll)
            ResolvedPath = $dll
        }
    }

    throw "Could not locate broker executable in $baseOut."
}

$failed = $false
$failureName = ""
$failureReason = ""

try {
    if ($TimeoutSeconds -lt 5) {
        throw "TimeoutSeconds must be >= 5."
    }

    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $helperPath = Join-Path $repoRoot "client\Invoke-PSBroker.ps1"
    if (-not (Test-Path -LiteralPath $helperPath)) {
        throw "Client helper not found: $helperPath"
    }

    if ([string]::IsNullOrWhiteSpace($PipeName)) {
        $PipeName = "psbroker-int-" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
    }

    $launch = Resolve-BrokerLaunch -RepoRoot $repoRoot -Config $Configuration -PathOverride $BrokerExePath
    Write-Output "Broker launch target: $($launch.ResolvedPath)"

    $script:BrokerStdOutPath = Join-Path $env:TEMP ("psbroker_it_broker_out_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:BrokerStdErrPath = Join-Path $env:TEMP ("psbroker_it_broker_err_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientStdOutPath = Join-Path $env:TEMP ("psbroker_it_client_out_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientStdErrPath = Join-Path $env:TEMP ("psbroker_it_client_err_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientScriptPath = Join-Path $env:TEMP ("psbroker_it_client_" + [Guid]::NewGuid().ToString("N") + ".ps1")

    $brokerArgs = @()
    $brokerArgs += $launch.Arguments
    $brokerArgs += @("--pipe", $PipeName, "--log-option", "silent")

    $script:BrokerProcess = Start-Process -FilePath $launch.FilePath `
        -ArgumentList $brokerArgs `
        -PassThru `
        -RedirectStandardOutput $script:BrokerStdOutPath `
        -RedirectStandardError $script:BrokerStdErrPath

    if (-not (Wait-BrokerReady -HelperPath $helperPath -Pipe $PipeName -TimeoutSec $TimeoutSeconds)) {
        $failed = $true
        $failureName = "TC1 Start/Reachable"
        $failureReason = "Broker did not become reachable within timeout."
        throw $failureReason
    }

    if ($script:BrokerProcess.HasExited) {
        $failed = $true
        $failureName = "TC1 Start/Reachable"
        $failureReason = "Broker exited before tests began."
        throw $failureReason
    }

    Write-Pass "TC1 Start/Reachable"

    $clientScript = @"
param(
    [Parameter(Mandatory=`$true)]
    [string]`$PipeName,
    [Parameter(Mandatory=`$true)]
    [string]`$HelperPath
)

Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"
. `$HelperPath

function Pass([string]`$Name) {
    Write-Output "PASS `$Name"
}

function Fail([string]`$Name, [string]`$Reason) {
    Write-Error ("FAIL {0}: {1}" -f `$Name, `$Reason)
    exit 1
}

try {
    `$infoResp = Invoke-PSBroker -PipeName `$PipeName -Kind native -Command "broker.info" -Raw
    if (-not `$infoResp.success) {
        Fail "TC2 Native broker.info" "broker.info returned non-success."
    }

    `$info = `$infoResp.stdout | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace(`$info.version) -or [string]::IsNullOrWhiteSpace(`$info.pipeName)) {
        Fail "TC2 Native broker.info" "Missing required info fields."
    }
    Pass "TC2 Native broker.info"

    `$helpResp = Invoke-PSBroker -PipeName `$PipeName -Kind native -Command "broker.help" -Raw
    if (-not `$helpResp.success) {
        Fail "TC3 Native broker.help" "broker.help returned non-success."
    }

    `$help = `$helpResp.stdout | ConvertFrom-Json
    if (-not `$help.ok -or `$help.status -ne "Success") {
        Fail "TC3 Native broker.help" "Unexpected help status."
    }

    `$commandNames = @(`$help.nativeCommands | ForEach-Object { `$_.name })
    foreach (`$required in @("broker.info", "broker.help", "broker.stop")) {
        if (`$commandNames -notcontains `$required) {
            Fail "TC3 Native broker.help" "Missing required command `$required."
        }
    }
    Pass "TC3 Native broker.help"

    `$dateResp = Invoke-PSBroker -PipeName `$PipeName -Kind powershell -Command "Get-Date" -Raw
    if (-not `$dateResp.success) {
        Fail "TC4 Non-native Get-Date" "Get-Date returned non-success."
    }

    if ([string]::IsNullOrWhiteSpace(`$dateResp.stdout)) {
        Fail "TC4 Non-native Get-Date" "Get-Date stdout was empty."
    }

    `$parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParse(`$dateResp.stdout.Trim(), [ref]`$parsedDate)) {
        Fail "TC4 Non-native Get-Date" "Get-Date output was not parseable as date/time."
    }
    Pass "TC4 Non-native Get-Date"

    `$stopResp = Invoke-PSBroker -PipeName `$PipeName -Kind native -Command "broker.stop" -Raw
    if (-not `$stopResp.success) {
        Fail "TC5 Stop request" "broker.stop returned non-success."
    }
    Pass "TC5 Stop request"
}
catch {
    Write-Error "FAIL ClientRun: unexpected client error"
    exit 1
}
exit 0
"@

    Set-Content -LiteralPath $script:ClientScriptPath -Value $clientScript -Encoding UTF8
    $pwshPath = Resolve-RunnerPwsh
    $script:ClientProcess = Start-Process -FilePath $pwshPath `
        -ArgumentList @("-NoProfile", "-File", $script:ClientScriptPath, "-PipeName", $PipeName, "-HelperPath", $helperPath) `
        -PassThru `
        -Wait `
        -RedirectStandardOutput $script:ClientStdOutPath `
        -RedirectStandardError $script:ClientStdErrPath

    $clientOut = Get-Content -LiteralPath $script:ClientStdOutPath -ErrorAction SilentlyContinue
    if ($clientOut) {
        $clientOut | ForEach-Object { Write-Output $_ }
    }

    if ($script:ClientProcess.ExitCode -ne 0) {
        $failed = $true
        $failureName = "Client test flow"
        $failureReason = "Client process exited with code $($script:ClientProcess.ExitCode)."
        throw $failureReason
    }

    if ($script:BrokerProcess.HasExited) {
        Write-Pass "TC5 Broker stopped cleanly"
    }
    else {
        try {
            Wait-Process -Id $script:BrokerProcess.Id -Timeout $TimeoutSeconds -ErrorAction Stop
        }
        catch {
            $failed = $true
            $failureName = "TC5 Broker stopped cleanly"
            $failureReason = "Broker did not exit after broker.stop."
            throw $failureReason
        }

        if (-not $script:BrokerProcess.HasExited) {
            $failed = $true
            $failureName = "TC5 Broker stopped cleanly"
            $failureReason = "Broker process still running after timeout."
            throw $failureReason
        }

        Write-Pass "TC5 Broker stopped cleanly"
    }

    exit 0
}
catch {
    if (-not $failed) {
        $failed = $true
        $failureName = if ([string]::IsNullOrWhiteSpace($failureName)) { "Run" } else { $failureName }
        $failureReason = if ([string]::IsNullOrWhiteSpace($failureReason)) { $PSItem.Exception.Message } else { $failureReason }
    }

    Write-Fail $failureName $failureReason

    [Console]::Error.WriteLine("---- broker stdout tail ----")
    [Console]::Error.WriteLine((Get-LogTail -Path $script:BrokerStdOutPath))
    [Console]::Error.WriteLine("---- broker stderr tail ----")
    [Console]::Error.WriteLine((Get-LogTail -Path $script:BrokerStdErrPath))
    [Console]::Error.WriteLine("---- client stdout tail ----")
    [Console]::Error.WriteLine((Get-LogTail -Path $script:ClientStdOutPath))
    [Console]::Error.WriteLine("---- client stderr tail ----")
    [Console]::Error.WriteLine((Get-LogTail -Path $script:ClientStdErrPath))

    exit 1
}
finally {
    Cleanup-Processes
    Remove-TempFiles
}

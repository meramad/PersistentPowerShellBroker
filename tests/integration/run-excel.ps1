param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [string]$BrokerExePath,

    [int]$TimeoutSeconds = 60,

    [string]$TempDir = "tests\integration\temp"
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
$script:WorkbookPath = $null

function Write-Pass {
    param([string]$Name)
    Write-Output "PASS $Name"
}

function Write-Skip {
    param([string]$Name, [string]$Reason)
    Write-Output ("SKIP {0}: {1}" -f $Name, $Reason)
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
        [string]$PipeName,
        [int]$TimeoutSec
    )

    . $HelperPath
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if ($script:BrokerProcess.HasExited) {
            return $false
        }

        try {
            $response = Invoke-PSBroker -PipeName $PipeName -Command "broker.info" -PassThru -TimeoutSeconds 2
            if ($response.success) {
                return $true
            }
        }
        catch {
            Start-Sleep -Milliseconds 300
            continue
        }

        Start-Sleep -Milliseconds 300
    }

    return $false
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

function New-ExcelWorkbookFile {
    param([string]$Path)

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        return @{
            Success = $false
            Reason = "Excel COM automation is unavailable."
        }
    }

    $workbook = $null
    try {
        $excel.DisplayAlerts = $false
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $workbook.SaveAs($Path)
        $workbook.Close($false)
        $excel.Quit()
    }
    catch {
        if ($workbook -ne $null) {
            try { $workbook.Close($false) } catch { }
        }
        try { $excel.Quit() } catch { }
        return @{
            Success = $false
            Reason = $_.Exception.Message
        }
    }
    finally {
        if ($workbook -ne $null) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch { }
        }
        if ($excel -ne $null) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
        }
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Success = $false
            Reason = "Workbook file was not created."
        }
    }

    return @{
        Success = $true
        Reason = ""
    }
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

$failed = $false
$failureName = ""
$failureReason = ""
$skipped = $false

try {
    if ($TimeoutSeconds -lt 10) {
        throw "TimeoutSeconds must be >= 10 for Excel integration."
    }

    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $helperPath = Join-Path $repoRoot "client\Invoke-PSBroker.ps1"
    if (-not (Test-Path -LiteralPath $helperPath)) {
        throw "Client helper not found: $helperPath"
    }
    $tempRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $TempDir))
    New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null

    $script:WorkbookPath = Join-Path $tempRoot "psbroker_excel_test.xlsx"
    if (Test-Path -LiteralPath $script:WorkbookPath) {
        Remove-Item -LiteralPath $script:WorkbookPath -Force -ErrorAction SilentlyContinue
    }

    $createResult = New-ExcelWorkbookFile -Path $script:WorkbookPath
    if (-not $createResult.Success) {
        Write-Skip "Excel tests" $createResult.Reason
        $skipped = $true
        exit 0
    }
    Write-Pass "TC1 Create temp workbook"

    $pipeName = "psbroker-excel-" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
    $launch = Resolve-BrokerLaunch -RepoRoot $repoRoot -Config $Configuration -PathOverride $BrokerExePath
    Write-Output "Broker launch target: $($launch.ResolvedPath)"

    $script:BrokerStdOutPath = Join-Path $env:TEMP ("psbroker_excel_broker_out_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:BrokerStdErrPath = Join-Path $env:TEMP ("psbroker_excel_broker_err_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientStdOutPath = Join-Path $env:TEMP ("psbroker_excel_client_out_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientStdErrPath = Join-Path $env:TEMP ("psbroker_excel_client_err_" + [Guid]::NewGuid().ToString("N") + ".log")
    $script:ClientScriptPath = Join-Path $env:TEMP ("psbroker_excel_client_" + [Guid]::NewGuid().ToString("N") + ".ps1")

    $brokerArgs = @()
    $brokerArgs += $launch.Arguments
    $brokerArgs += @("--pipe", $pipeName, "--log-option", "silent")

    $script:BrokerProcess = Start-Process -FilePath $launch.FilePath `
        -ArgumentList $brokerArgs `
        -PassThru `
        -RedirectStandardOutput $script:BrokerStdOutPath `
        -RedirectStandardError $script:BrokerStdErrPath

    if (-not (Wait-BrokerReady -HelperPath $helperPath -PipeName $pipeName -TimeoutSec $TimeoutSeconds)) {
        $failed = $true
        $failureName = "TC2 Start broker"
        $failureReason = "Broker did not become reachable within timeout."
        throw $failureReason
    }
    Write-Pass "TC2 Start broker"

    $clientScript = @"
param(
    [Parameter(Mandatory=`$true)]
    [string]`$PipeName,
    [Parameter(Mandatory=`$true)]
    [string]`$HelperPath,
    [Parameter(Mandatory=`$true)]
    [string]`$WorkbookPath,
    [Parameter(Mandatory=`$true)]
    [int]`$TimeoutSeconds
)

Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"
. `$HelperPath

function Pass([string]`$Name) {
    Write-Output "PASS `$Name"
}

function Fail([string]`$Name, [string]`$Reason) {
    [Console]::Error.WriteLine(("FAIL {0}: {1}" -f `$Name, `$Reason))
    exit 1
}

function Parse-NativePayload {
    param([object]`$BrokerResponse)
    if (-not `$BrokerResponse.success) {
        throw ("Broker request failed. error={0} stderr={1}" -f `$BrokerResponse.error, `$BrokerResponse.stderr)
    }

    return (`$BrokerResponse.stdout | ConvertFrom-Json)
}

function Probe-HandleVariable {
    param([string]`$VariableName)

    `$escapedName = `$VariableName.Replace("'", "''")
    `$commandTemplate = '`$var = Get-Variable -Name ''__VAR_NAME__'' -Scope Global -ErrorAction SilentlyContinue; if (`$null -eq `$var) { @{ exists = `$false } | ConvertTo-Json -Compress } else { @{ exists = `$true; hasApplication = (`$null -ne `$var.Value.Application); hasWorkbook = (`$null -ne `$var.Value.Workbook); requestedTarget = [string]`$var.Value.Metadata.RequestedTarget; workbookFullName = [string]`$var.Value.Workbook.FullName } | ConvertTo-Json -Compress }'
    `$command = `$commandTemplate.Replace('__VAR_NAME__', `$escapedName)

    `$probeResp = Invoke-PSBroker -PipeName `$PipeName -Script `$command -PassThru -TimeoutSeconds `$TimeoutSeconds
    if (-not `$probeResp.success) {
        throw ("Probe command failed. error={0} stderr={1}" -f `$probeResp.error, `$probeResp.stderr)
    }

    return (`$probeResp.stdout.Trim() | ConvertFrom-Json)
}

try {
    `$absolutePath = [System.IO.Path]::GetFullPath(`$WorkbookPath)
    `$externalExcel = `$null
    `$externalWorkbook = `$null

    `$get1Resp = Invoke-PSBroker -PipeName `$PipeName -Command "broker.excel.get_workbook_handle" -Args @{
        path = `$absolutePath
        timeoutSeconds = `$TimeoutSeconds
        forceVisible = `$true
        displayAlerts = `$false
    } -PassThru -TimeoutSeconds `$TimeoutSeconds
    `$get1 = Parse-NativePayload -BrokerResponse `$get1Resp
    if (-not `$get1.ok -or `$get1.status -ne "Success") {
        Fail "TC3 get_workbook_handle not-open" "Command returned non-success payload."
    }
    if ([string]::IsNullOrWhiteSpace(`$get1.psVariableName)) {
        Fail "TC3 get_workbook_handle not-open" "psVariableName was empty."
    }
    if ((-not `$get1.openedWorkbook) -and (-not `$get1.attachedExisting)) {
        Fail "TC3 get_workbook_handle not-open" "Neither openedWorkbook nor attachedExisting was true."
    }
    if ([string]::IsNullOrWhiteSpace(`$get1.workbookFullName)) {
        Fail "TC3 get_workbook_handle not-open" "workbookFullName was empty."
    }
    Pass "TC3 get_workbook_handle not-open"

    `$probe1 = Probe-HandleVariable -VariableName `$get1.psVariableName
    if (-not `$probe1.exists -or -not `$probe1.hasApplication -or -not `$probe1.hasWorkbook) {
        Fail "TC4 probe variable bundle" "Variable bundle shape was invalid."
    }
    if ([string]::IsNullOrWhiteSpace(`$probe1.workbookFullName)) {
        Fail "TC4 probe variable bundle" "Probe workbookFullName was empty."
    }
    if (-not [string]::Equals([System.IO.Path]::GetFullPath(`$probe1.requestedTarget), `$absolutePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        Fail "TC4 probe variable bundle" "requestedTarget did not match requested path."
    }
    Pass "TC4 probe variable bundle"

    `$release1Resp = Invoke-PSBroker -PipeName `$PipeName -Command "broker.excel.release_handle" -Args @{
        psVariableName = `$get1.psVariableName
        closeWorkbook = `$true
        saveChanges = `$false
        quitExcel = `$true
        onlyIfNoOtherWorkbooks = `$true
        displayAlerts = `$false
    } -PassThru -TimeoutSeconds `$TimeoutSeconds
    `$release1 = Parse-NativePayload -BrokerResponse `$release1Resp
    if (-not `$release1.ok -or `$release1.status -ne "Success" -or -not `$release1.released) {
        Fail "TC4b release first handle" "First release failed."
    }
    Pass "TC4b release first handle"

    Start-Sleep -Milliseconds 500

    try {
        `$externalExcel = New-Object -ComObject Excel.Application
        `$externalExcel.DisplayAlerts = `$false
        `$externalExcel.Visible = `$true
        `$externalWorkbook = `$externalExcel.Workbooks.Open(`$absolutePath)
    }
    catch {
        Fail "TC5 Prepare already-open workbook" "Failed opening workbook outside broker."
    }
    Pass "TC5 Prepare already-open workbook"

    `$get2Resp = Invoke-PSBroker -PipeName `$PipeName -Command "broker.excel.get_workbook_handle" -Args @{
        path = `$absolutePath
        timeoutSeconds = `$TimeoutSeconds
        forceVisible = `$true
        displayAlerts = `$false
    } -PassThru -TimeoutSeconds `$TimeoutSeconds
    `$get2 = Parse-NativePayload -BrokerResponse `$get2Resp
    if (-not `$get2.ok -or `$get2.status -ne "Success") {
        Fail "TC5 get_workbook_handle already-open" "Command returned non-success payload."
    }
    if (-not `$get2.attachedExisting) {
        Fail "TC5 get_workbook_handle already-open" "Expected attachedExisting=true."
    }
    if (`$get2.openedWorkbook) {
        Fail "TC5 get_workbook_handle already-open" "Expected openedWorkbook=false when attaching."
    }
    if ([string]::IsNullOrWhiteSpace(`$get2.psVariableName)) {
        Fail "TC5 get_workbook_handle already-open" "psVariableName was empty."
    }
    Pass "TC5 get_workbook_handle already-open"

    `$release2Resp = Invoke-PSBroker -PipeName `$PipeName -Command "broker.excel.release_handle" -Args @{
        psVariableName = `$get2.psVariableName
        closeWorkbook = `$true
        saveChanges = `$false
        quitExcel = `$true
        onlyIfNoOtherWorkbooks = `$true
        displayAlerts = `$false
    } -PassThru -TimeoutSeconds `$TimeoutSeconds
    `$release2 = Parse-NativePayload -BrokerResponse `$release2Resp
    if (-not `$release2.ok -or `$release2.status -ne "Success" -or -not `$release2.released) {
        Fail "TC6 release_handle cleanup" "release_handle returned non-success."
    }
    if (`$release2.quitExcelAttempted -and (-not `$release2.quitExcelSucceeded) -and (-not `$release2.quitSkipped)) {
        Fail "TC6 release_handle cleanup" "quitExcelAttempted without success/skip."
    }
    Pass "TC6 release_handle cleanup"

    `$probe2 = Probe-HandleVariable -VariableName `$get2.psVariableName
    if (`$probe2.exists) {
        Fail "TC7 variable removed" "Handle variable still exists after release."
    }
    Pass "TC7 variable removed"

    `$stopResp = Invoke-PSBroker -PipeName `$PipeName -Command "broker.stop" -PassThru -TimeoutSeconds `$TimeoutSeconds
    if (-not `$stopResp.success) {
        Fail "TC8 broker.stop" "broker.stop returned non-success."
    }
    Pass "TC8 broker.stop"
}
catch {
    [Console]::Error.WriteLine(("FAIL ClientRun: {0}" -f `$PSItem.Exception.Message))
    exit 1
}
finally {
    if (`$externalWorkbook -ne `$null) {
        try { `$externalWorkbook.Close(`$false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(`$externalWorkbook) } catch { }
    }

    if (`$externalExcel -ne `$null) {
        try { `$externalExcel.DisplayAlerts = `$false } catch { }
        try { `$externalExcel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(`$externalExcel) } catch { }
    }
}

exit 0
"@

    Set-Content -LiteralPath $script:ClientScriptPath -Value $clientScript -Encoding UTF8
    $pwshPath = Resolve-RunnerPwsh
    $script:ClientProcess = Start-Process -FilePath $pwshPath `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $script:ClientScriptPath, "-PipeName", $pipeName, "-HelperPath", $helperPath, "-WorkbookPath", $script:WorkbookPath, "-TimeoutSeconds", $TimeoutSeconds) `
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
        $failureName = "Excel client test flow"
        $failureReason = "Client process exited with code $($script:ClientProcess.ExitCode)."
        throw $failureReason
    }

    if (-not $script:BrokerProcess.HasExited) {
        try {
            Wait-Process -Id $script:BrokerProcess.Id -Timeout $TimeoutSeconds -ErrorAction Stop
        }
        catch {
            $failed = $true
            $failureName = "TC9 Broker stopped cleanly"
            $failureReason = "Broker did not exit after broker.stop."
            throw $failureReason
        }
    }
    Write-Pass "TC9 Broker stopped cleanly"

    if (Test-Path -LiteralPath $script:WorkbookPath) {
        Remove-Item -LiteralPath $script:WorkbookPath -Force -ErrorAction SilentlyContinue
    }
    Write-Pass "TC10 Cleanup temp workbook"

    exit 0
}
catch {
    if (-not $failed -and -not $skipped) {
        $failed = $true
        $failureName = if ([string]::IsNullOrWhiteSpace($failureName)) { "RunExcel" } else { $failureName }
        $failureReason = if ([string]::IsNullOrWhiteSpace($failureReason)) { $PSItem.Exception.Message } else { $failureReason }
    }

    if (-not $skipped) {
        Write-Fail $failureName $failureReason
        [Console]::Error.WriteLine("---- broker stdout tail ----")
        [Console]::Error.WriteLine((Get-LogTail -Path $script:BrokerStdOutPath))
        [Console]::Error.WriteLine("---- broker stderr tail ----")
        [Console]::Error.WriteLine((Get-LogTail -Path $script:BrokerStdErrPath))
        [Console]::Error.WriteLine("---- client stdout tail ----")
        [Console]::Error.WriteLine((Get-LogTail -Path $script:ClientStdOutPath))
        [Console]::Error.WriteLine("---- client stderr tail ----")
        [Console]::Error.WriteLine((Get-LogTail -Path $script:ClientStdErrPath))
    }

    exit 1
}
finally {
    Cleanup-Processes
    Remove-TempFiles
}


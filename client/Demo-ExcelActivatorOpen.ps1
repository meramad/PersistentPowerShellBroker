param(
    [switch]$QuitAfterOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$type = [Type]::GetTypeFromProgID('Excel.Application', $false)
if ($null -eq $type) {
    throw 'Excel COM ProgID not found: Excel.Application'
}

$excel = [Activator]::CreateInstance($type)
if ($null -eq $excel) {
    throw 'Activator.CreateInstance returned null for Excel.Application'
}

try {
    $excel.Visible = $true
    $excel.UserControl = $true

    $hwnd = 0
    try { $hwnd = [int]$excel.Hwnd } catch { }

    Write-Host 'Excel COM instance created via Activator.CreateInstance.'
    Write-Host "Visible=$($excel.Visible) UserControl=$($excel.UserControl) Hwnd=$hwnd"

    if ($QuitAfterOpen) {
        $excel.Quit()
        Write-Host 'Excel instance closed because -QuitAfterOpen was specified.'
    }
}
finally {
    if ($null -ne $excel -and [System.Runtime.InteropServices.Marshal]::IsComObject($excel)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

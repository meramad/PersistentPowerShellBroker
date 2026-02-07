function Invoke-PSBroker {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName='Path', Mandatory=$true)]
        [string]$Pipe,

        [Parameter(ParameterSetName='Name', Mandatory=$true)]
        [string]$PipeName,

        [Parameter(Mandatory=$true)]
        [string]$Command,

        [ValidateSet('powershell','native')]
        [string]$Kind = 'powershell',

        [string]$ClientName = 'powershell',

        [int]$ClientPid = $PID,

        [switch]$Raw
    )

    if ($PSCmdlet.ParameterSetName -eq 'Name') {
        $Pipe = "\\.\pipe\$PipeName"
    }

    if (-not $Pipe.StartsWith('\\.\pipe\')) {
        throw "Pipe must start with \\.\pipe\"
    }

    $name = $Pipe.Substring('\\.\pipe\'.Length)
    $request = [ordered]@{
        id = [guid]::NewGuid().ToString('N')
        kind = $Kind
        command = $Command
        clientName = $ClientName
        clientPid = $ClientPid
    }

    $requestJson = ($request | ConvertTo-Json -Compress)

    $client = [System.IO.Pipes.NamedPipeClientStream]::new('.', $name, [System.IO.Pipes.PipeDirection]::InOut)
    try {
        $client.Connect(5000)
        $writer = [System.IO.StreamWriter]::new($client, [System.Text.UTF8Encoding]::new($false), 1024, $true)
        $reader = [System.IO.StreamReader]::new($client, [System.Text.UTF8Encoding]::new($false), $false, 1024, $true)

        $writer.WriteLine($requestJson)
        $writer.Flush()

        $line = $reader.ReadLine()
        if ([string]::IsNullOrWhiteSpace($line)) {
            throw 'Broker returned no response line.'
        }

        $response = $line | ConvertFrom-Json
        if ($Raw) {
            return $response
        }

        if (-not [string]::IsNullOrEmpty($response.stderr)) {
            Write-Error $response.stderr
        }

        if (-not [string]::IsNullOrEmpty($response.stdout)) {
            Write-Output $response.stdout
        }

        if (-not $response.success -and -not [string]::IsNullOrEmpty($response.error)) {
            Write-Error $response.error
        }
    }
    finally {
        $client.Dispose()
    }
}

param(
  [string]$Name = 'Broker'
)

"Hello from client script: $Name"
"Timestamp: $([DateTime]::UtcNow.ToString('o'))"

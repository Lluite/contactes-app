$port = 8765
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = "C:\Users\Maso\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$url = "http://127.0.0.1:$port/index.html?v=5"

$listenerBusy = $false
try {
    $client = New-Object System.Net.Sockets.TcpClient
    $client.Connect("127.0.0.1", $port)
    $client.Close()
    $listenerBusy = $true
} catch {
    $listenerBusy = $false
}

if (-not $listenerBusy) {
    Start-Process -FilePath $python -ArgumentList "-m", "http.server", "$port", "--bind", "127.0.0.1" -WorkingDirectory $root -WindowStyle Hidden
    Start-Sleep -Seconds 1
}

Start-Process $url

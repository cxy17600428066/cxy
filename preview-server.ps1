$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultFile = Join-Path $root "product-manual-preview.html"
$port = 8123

function Get-ContentType([string]$file) {
  switch ([System.IO.Path]::GetExtension($file).ToLowerInvariant()) {
    ".html" { return "text/html; charset=utf-8" }
    ".js" { return "application/javascript; charset=utf-8" }
    ".css" { return "text/css; charset=utf-8" }
    ".json" { return "application/json; charset=utf-8" }
    default { return "application/octet-stream" }
  }
}

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $port)
$listener.Start()

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII, $false, 1024, $true)
      $requestLine = $reader.ReadLine()
      if ([string]::IsNullOrWhiteSpace($requestLine)) {
        $client.Close()
        continue
      }

      while (($line = $reader.ReadLine()) -ne "") { }

      $parts = $requestLine.Split(" ")
      $path = if ($parts.Length -ge 2) { $parts[1] } else { "/" }
      $path = $path.Split("?")[0]
      $path = [System.Uri]::UnescapeDataString($path.TrimStart("/"))
      $file = if ([string]::IsNullOrWhiteSpace($path)) { $defaultFile } else { Join-Path $root ($path -replace "/", "\") }

      if (-not (Test-Path -LiteralPath $file)) {
        $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
        $header = "HTTP/1.1 404 Not Found`r`nContent-Type: text/plain; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n"
        $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
        $stream.Write($headerBytes, 0, $headerBytes.Length)
        $stream.Write($body, 0, $body.Length)
        $stream.Flush()
        $client.Close()
        continue
      }

      $body = [System.IO.File]::ReadAllBytes($file)
      $header = "HTTP/1.1 200 OK`r`nContent-Type: $(Get-ContentType $file)`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n"
      $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
      $stream.Write($headerBytes, 0, $headerBytes.Length)
      $stream.Write($body, 0, $body.Length)
      $stream.Flush()
    }
    finally {
      $client.Close()
    }
  }
}
finally {
  $listener.Stop()
}

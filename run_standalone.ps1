param(
  [int]$Port = 8091
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Prefix = "http://127.0.0.1:$Port/"

Add-Type -AssemblyName System.Web

$mimeMap = @{
  '.html' = 'text/html; charset=utf-8'
  '.js' = 'application/javascript; charset=utf-8'
  '.json' = 'application/json; charset=utf-8'
  '.css' = 'text/css; charset=utf-8'
  '.png' = 'image/png'
  '.jpg' = 'image/jpeg'
  '.jpeg' = 'image/jpeg'
  '.svg' = 'image/svg+xml'
  '.pdf' = 'application/pdf'
  '.txt' = 'text/plain; charset=utf-8'
}

function Get-MimeType([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  if ($mimeMap.ContainsKey($ext)) {
    return $mimeMap[$ext]
  }
  return 'application/octet-stream'
}

function Resolve-RequestPath([string]$urlPath) {
  $decoded = [System.Web.HttpUtility]::UrlDecode($urlPath)
  if ([string]::IsNullOrWhiteSpace($decoded) -or $decoded -eq '/') {
    return 'ocr_picking_ticket_standalone.html'
  }

  $trimmed = $decoded.TrimStart('/')
  if ($trimmed -eq '') {
    return 'ocr_picking_ticket_standalone.html'
  }

  return $trimmed -replace '/', [System.IO.Path]::DirectorySeparatorChar
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($Prefix)
$listener.Start()

Write-Host "Starting standalone OCR Picking Ticket server..."
Write-Host "Folder: $Root"
Write-Host "URL: $Prefix"
Write-Host "(Press Ctrl+C to stop)"

Start-Process $Prefix | Out-Null

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $requestPath = Resolve-RequestPath $context.Request.Url.AbsolutePath
    $safePath = [System.IO.Path]::GetFullPath((Join-Path $Root $requestPath))

    if (-not $safePath.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase)) {
      $context.Response.StatusCode = 403
      $context.Response.Close()
      continue
    }

    if (-not (Test-Path -LiteralPath $safePath -PathType Leaf)) {
      $context.Response.StatusCode = 404
      $bytes = [System.Text.Encoding]::UTF8.GetBytes('Not Found')
      $context.Response.ContentType = 'text/plain; charset=utf-8'
      $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
      $context.Response.Close()
      continue
    }

    $bytes = [System.IO.File]::ReadAllBytes($safePath)
    $context.Response.StatusCode = 200
    $context.Response.ContentType = Get-MimeType $safePath
    $context.Response.ContentLength64 = $bytes.Length
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $context.Response.Close()
  }
}
finally {
  if ($listener.IsListening) {
    $listener.Stop()
  }
  $listener.Close()
}

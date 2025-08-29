$ErrorActionPreference = 'Stop'
param(
  [int]$Port = 5173,
  [string]$Root = (Join-Path $PSScriptRoot 'dist')
)

Add-Type -AssemblyName System.Net.HttpListener
if (-not [System.Net.HttpListener]::IsSupported) {
  Write-Host 'HttpListener not supported on this system.' -ForegroundColor Red
  exit 1
}

$prefix = "http://localhost:$Port/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
try { $listener.Start() } catch { Write-Host $_; exit 1 }
Write-Host "Serving '$Root' at $prefix (Ctrl+C or close window to stop)" -ForegroundColor Green

function Get-ContentType($path) {
  switch ([IO.Path]::GetExtension($path).ToLower()) {
    '.html' { 'text/html; charset=utf-8' }
    '.js'   { 'text/javascript; charset=utf-8' }
    '.css'  { 'text/css; charset=utf-8' }
    '.map'  { 'application/json; charset=utf-8' }
    '.json' { 'application/json; charset=utf-8' }
    '.svg'  { 'image/svg+xml' }
    '.png'  { 'image/png' }
    '.jpg'  { 'image/jpeg' }
    '.jpeg' { 'image/jpeg' }
    '.gif'  { 'image/gif' }
    '.ico'  { 'image/x-icon' }
    '.woff' { 'font/woff' }
    '.woff2'{ 'font/woff2' }
    '.ttf'  { 'font/ttf' }
    default { 'application/octet-stream' }
  }
}

while ($listener.IsListening) {
  try {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response

    $rel = $req.Url.AbsolutePath.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }
    $path = Join-Path $Root $rel
    if (-not (Test-Path $path)) {
      # SPA fallback to index.html
      $path = Join-Path $Root 'index.html'
    }
    try {
      $bytes = [IO.File]::ReadAllBytes($path)
      $res.ContentType = Get-ContentType $path
      $res.ContentLength64 = $bytes.Length
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
      $res.StatusCode = 200
    } catch {
      $res.StatusCode = 404
    } finally {
      $res.OutputStream.Close()
    }
  } catch {
    break
  }
}

try { $listener.Stop() } catch {}

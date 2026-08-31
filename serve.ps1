# Simple static web server for testing the app locally
# Usage: powershell -ExecutionPolicy Bypass -File serve.ps1  → open http://localhost:8642
$root = $PSScriptRoot
$port = 8642
$mime = @{
  ".html"="text/html; charset=utf-8"; ".js"="text/javascript; charset=utf-8"
  ".css"="text/css; charset=utf-8"; ".png"="image/png"; ".jpg"="image/jpeg"
  ".webmanifest"="application/manifest+json; charset=utf-8"; ".json"="application/json"
  ".svg"="image/svg+xml"; ".ico"="image/x-icon"
}
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Serving $root at http://localhost:$port/"
while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $path = $ctx.Request.Url.AbsolutePath.TrimStart("/")
  if ([string]::IsNullOrEmpty($path)) { $path = "index.html" }
  $file = Join-Path $root ($path -replace "/", "\")
  if ((Test-Path $file -PathType Leaf) -and $file.StartsWith($root)) {
    $ext = [System.IO.Path]::GetExtension($file).ToLower()
    $ct = $mime[$ext]; if (-not $ct) { $ct = "application/octet-stream" }
    $bytes = [System.IO.File]::ReadAllBytes($file)
    $ctx.Response.ContentType = $ct
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } else {
    $ctx.Response.StatusCode = 404
    $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
    $ctx.Response.OutputStream.Write($msg, 0, $msg.Length)
  }
  $ctx.Response.Close()
}

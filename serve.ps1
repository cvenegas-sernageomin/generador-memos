# serve.ps1 - sirve pwa/ en http://localhost:8130 para probar/instalar la PWA
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$pwa = Join-Path $root "pwa"
Write-Host "Sirviendo $pwa en http://localhost:8130 (Ctrl+C para detener)" -ForegroundColor Green
python -m http.server 8130 --directory $pwa

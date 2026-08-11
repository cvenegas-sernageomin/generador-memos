# build.ps1 - genera dist/GeneradorMemos_vN.html (autocontenido, sin Node) + pwa/index.html
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

# Version incremental
$versionFile = "$root\VERSION"
$version = 0
if (Test-Path $versionFile) { $version = [int](Get-Content $versionFile -Raw).Trim() }
$version++
Set-Content -Path $versionFile -Value $version -NoNewline -Encoding ascii
$versionStr = "v$version"

# Plantilla base -> base64 (se embebe en la app para que funcione offline / doble clic)
$tplBytes = [IO.File]::ReadAllBytes("$root\assets\plantilla_base.xlsx")
$tplB64 = [Convert]::ToBase64String($tplBytes)

$react    = [IO.File]::ReadAllText("$root\vendor\react.production.min.js")
$reactDom = [IO.File]::ReadAllText("$root\vendor\react-dom.production.min.js")
$jszip    = [IO.File]::ReadAllText("$root\vendor\jszip.min.js")
$babel    = [IO.File]::ReadAllText("$root\vendor\babel.min.js")
$app      = [IO.File]::ReadAllText("$root\src\app.jsx")
$tpl      = [IO.File]::ReadAllText("$root\src\template.html")

# Inyecciones en la app
$app = $app.Replace("/*__TEMPLATE_B64__*/", $tplB64)
$app = $app.Replace("__APP_VERSION__", $versionStr)
# Defensa: evitar cierre prematuro del <script>
$app = $app.Replace("</script", "<\/script")

$out = $tpl.Replace("/*__REACT__*/",    $react)
$out = $out.Replace("/*__REACTDOM__*/", $reactDom)
$out = $out.Replace("/*__JSZIP__*/",    $jszip)
$out = $out.Replace("/*__BABEL__*/",    $babel)
$out = $out.Replace("/*__APP__*/",      $app)

# UTF-8 CON BOM (acentos)
$enc = New-Object System.Text.UTF8Encoding($true)

# 1) HTML autocontenido (file://) para compartir por correo/USB
$distDir = "$root\dist"
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
$dest = "$distDir\GeneradorMemos_$versionStr.html"
[IO.File]::WriteAllText($dest, $out, $enc)
$kb = [Math]::Round((Get-Item $dest).Length / 1KB)
Write-Host "OK -> $dest ($kb KB) [$versionStr]" -ForegroundColor Green

# 2) index.html de la PWA (mismo contenido; se sirve por http junto a sw/manifest/icons)
$pwaDir = "$root\pwa"
if (-not (Test-Path $pwaDir)) { New-Item -ItemType Directory -Path $pwaDir | Out-Null }
$pwaDest = "$pwaDir\index.html"
[IO.File]::WriteAllText($pwaDest, $out, $enc)
Write-Host "OK -> $pwaDest" -ForegroundColor Green

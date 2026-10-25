# Crea los accesos directos de Alta de productos (escritorio y menú Inicio).
# Lo llama instalar.bat. -Destinos permite probar la instalación en otra carpeta.
param([string[]]$Destinos)

$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not (Test-Path (Join-Path $dir "node\node.exe")) -or -not (Test-Path (Join-Path $dir "app\server.js"))) {
  Write-Host "Faltan archivos (node\node.exe o app\server.js). Descomprime el .zip completo y vuelve a intentar." -ForegroundColor Red
  exit 1
}

# Quita el bloqueo de Windows a los archivos que vienen de un .zip descargado
Get-ChildItem -Path $dir -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

if (-not $Destinos) {
  $Destinos = @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))
}
$sh = New-Object -ComObject WScript.Shell
foreach ($d in $Destinos) {
  New-Item -ItemType Directory -Force $d | Out-Null
  $lnk = $sh.CreateShortcut((Join-Path $d "Alta de productos.lnk"))
  $lnk.TargetPath = Join-Path $env:WINDIR "System32\wscript.exe"
  $lnk.Arguments = '"' + (Join-Path $dir "iniciar.vbs") + '"'
  $lnk.WorkingDirectory = $dir
  $lnk.IconLocation = (Join-Path $env:WINDIR "System32\shell32.dll") + ",165"
  $lnk.Description = "Alta de productos (Siigo y BarTender)"
  $lnk.Save()
  Write-Host "Acceso directo creado en: $d"
}
Write-Host ""
Write-Host "Listo. Abre 'Alta de productos' desde el escritorio." -ForegroundColor Green

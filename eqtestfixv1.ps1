Param(
    [Parameter(Mandatory=$false)]
    [string]$playbackDeviceName
)

Add-Type -AssemblyName System.Windows.Forms

function exitWithErrorMsg ($msg){
    [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "EnableLoudnessEQ",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

# Auto elevate (Falcosc logic)
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent()
)

if (-not $currentPrincipal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Start-Process powershell `
        -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""

    return
}

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "Enable Loudness EQ (ReleaseTime = 2)" -ForegroundColor Cyan

# Ruta correcta estilo Falcosc
$renderBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$devices = Get-ChildItem "$renderBase\*\Properties"

if ($devices.Count -eq 0) {
    exitWithErrorMsg "Cannot access audio devices."
}

# Obtener dispositivos activos
$renderer = @()

foreach($device in $devices) {

    $friendly = $device.GetValue("{a45c254e-df1c-4efd-8020-67d146a850e0},2")

    if ($friendly) {

        $parent = Get-ItemProperty $device.PSParentPath

        if ($parent.DeviceState -eq 1) {

            $renderer += [PSCustomObject]@{
                Name = $friendly
                FxPath = $device.PSParentPath.Replace(
                    "Microsoft.PowerShell.Core\Registry::",""
                ) + "\FxProperties"
            }
        }
    }
}

if ($renderer.Count -eq 0) {
    exitWithErrorMsg "No active playback devices found."
}

# Mostrar lista si no se especifica
if (-not $playbackDeviceName) {

    Write-Host ""
    Write-Host "Available devices:" -ForegroundColor Yellow

    foreach ($r in $renderer) {
        Write-Host " - $($r.Name)"
    }

    Write-Host ""
    $playbackDeviceName = Read-Host "Enter device name"
}

# Buscar coincidencia
$activeRenderer = $renderer | Where-Object {
    $_.Name -like "*$playbackDeviceName*"
}

if ($activeRenderer.Count -eq 0) {
    exitWithErrorMsg "Device not found: $playbackDeviceName"
}

# Crear archivo REG temporal
$regFile = "$env:temp\EnableLoudnessEQ.reg"

"Windows Registry Editor Version 5.00" | Out-File $regFile -Encoding ascii

foreach ($dev in $activeRenderer) {

@"
[$($dev.FxPath)]

"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}"

"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:0b,00,00,00,01,00,00,00,ff,ff,00,00

"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3"=hex:03,00,00,60,02,00,00,00,07,00,00,00

"@ | Out-File $regFile -Append -Encoding ascii
}

Write-Host ""
Write-Host "Applying Loudness EQ..." -ForegroundColor Green

Start-Process "$Env:SystemRoot\REGEDIT.exe" `
    -ArgumentList "/s `"$regFile`"" `
    -Wait

# Reiniciar audio (CRÍTICO)
Restart-Service audiosrv -Force

# Limpiar archivo temporal
Remove-Item $regFile -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Loudness EQ applied successfully." -ForegroundColor Green

Read-Host "Press ENTER to exit"
<#
    PRUEBA DE CONCEPTO: SUBIDA A MEGA USANDO MEGA CMD
    Este script automatiza la instalación de MEGA CMD y la subida de archivos.
#>

# --- FUNCIONES AUXILIARES ---
Function Import-EnvFile {
    Param ([string]$Path = (Join-Path $PSScriptRoot ".env"))
    $envVars = @{}
    If (Test-Path $Path) {
        Get-Content $Path | ForEach-Object {
            If ($_ -match "^\s*([^#=\s]+)\s*=\s*(.*)\s*$") {
                $key = $Matches[1]
                $value = $Matches[2].Trim("'").Trim('"').Trim()
                
                # Expandir variables de entorno (como %VAR% o $env:VAR)
                $expandedValue = [System.Environment]::ExpandEnvironmentVariables($value)
                if ($key -notmatch "PASS|SECRET|TOKEN") {
                    Try {
                        $expandedValue = $ExecutionContext.InvokeCommand.ExpandString($expandedValue)
                    } Catch { }
                }
                $envVars[$key] = $expandedValue
            }
        }
    }
    Return $envVars
}

# 1. Cargar Configuración
$envConfig = Import-EnvFile
$MegaUser   = $envConfig["MEGA_USER"]
$MegaPass   = $envConfig["MEGA_PASS"]
$RemoteDest = $envConfig["MEGA_REMOTE_DEST"]
$LocalPath  = $envConfig["PST_SOURCE_DIR"]

# Resolver archivo PST si es una carpeta (elige el más pequeño para pruebas)
if (Test-Path $LocalPath) {
    if (Test-Path $LocalPath -PathType Container) {
        $SmallestPST = Get-ChildItem -Path $LocalPath -Filter "*.pst" | Sort-Object Length | Select-Object -First 1
        if ($SmallestPST) {
            $LocalPath = $SmallestPST.FullName
            Write-Host "[+] Archivo seleccionado: $($SmallestPST.Name) ($([math]::Round($SmallestPST.Length / 1MB, 2)) MB)" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "[ERROR] La ruta de origen no existe: $LocalPath" -ForegroundColor Red
    Exit 1
}

# 2. Instalación/Verificación de MEGA CMD
if (-not (Get-Command "mega-put" -ErrorAction SilentlyContinue)) {
    Write-Host "[+] Instalando MEGA CMD..." -ForegroundColor Cyan
    winget install --id MEGA.MEGACMD --silent --accept-source-agreements --accept-package-agreements
    $env:Path += ";C:\AppData\Local\MEGAcmd;C:\Program Files\MEGA CMD"
}

# 3. Iniciar Servidor si es necesario
if (-not (Get-Process "Mega-cmd-server" -ErrorAction SilentlyContinue)) {
    $MegaServerPath = (Get-Command "Mega-cmd-server.exe" -ErrorAction SilentlyContinue).Source
    if ($MegaServerPath) {
        Start-Process $MegaServerPath -WindowStyle Hidden
        Start-Sleep -Seconds 5
    }
}

# 4. Login y Subida
Write-Host "[!] Verificando sesión..." -ForegroundColor Yellow
$LoginCheck = mega-whoami 2>&1
if ($LoginCheck -match "Not logged in") {
    mega-login $MegaUser $MegaPass
}

Write-Host "[!] Preparando archivo con nombre de equipo y fecha..." -ForegroundColor Yellow
$Timestamp = Get-Date -Format "yyyy-MM-dd-HH"
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($LocalPath)
$BackupFileName = "$($env:COMPUTERNAME)_$($BaseName)_$($Timestamp).pst"
$TempPath = Join-Path $env:TEMP $BackupFileName

# Copiar a temporal con el nuevo nombre para la subida
Copy-Item "$LocalPath" "$TempPath" -Force

Write-Host "[!] Subiendo a MEGA: $RemoteDest/$BackupFileName" -ForegroundColor Yellow
mega-put "$TempPath" "$RemoteDest"

# Limpiar temporal tras la subida
Remove-Item "$TempPath" -ErrorAction SilentlyContinue

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[SUCCESS] Subida completada." -ForegroundColor Green
} else {
    Write-Host "`n[ERROR] Falló la subida." -ForegroundColor Red
}

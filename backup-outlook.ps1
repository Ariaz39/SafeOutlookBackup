# PowerShell Script for Outlook PST Backup to OneDrive

<#
.SYNOPSIS
    This script performs a critical backup of Outlook PST data files to OneDrive, ensuring absolute integrity and traceability.
    It includes graceful Outlook shutdown, local snapshot, SHA256 integrity validation, and detailed email notifications.

.DESCRIPTION
    The script automates the backup process for multiple large Outlook PST files (20GB+).
    It dynamically defines paths from a .env file, ensures Outlook is closed safely before copying,
    validates data integrity for each file, moves validated backups to OneDrive, and sends comprehensive email notifications.
    Error handling, logging, and secure credential management for SMTP are central to its design.

.NOTES
    Author: AI Automation Engineer
    Version: 2.1
    Date: 2026-05-04

    Prerequisites:
    - PowerShell 7.6.1 or later.
    - Outlook application must be installed.
    - OneDrive sync client must be configured and running.
    - SMTP credentials must be securely stored using Export-CliXml (e.g., $creds = Get-Credential | Export-CliXml -Path "C:\Ruta\smtp_creds.xml").
    - Staging folder must exist and ideally be on a fast drive (SSD).
    - Exclusions for the staging folder in antivirus software are recommended for performance.
    - A '.env' file must be present in the script's directory with the required environment variables.

.PARAMETER PST_Source_Dir
    The path to the directory containing the original Outlook PST files.
.PARAMETER Staging_Folder
    The path to the temporary staging folder where the PSTs will be copied before validation.
.PARAMETER OneDrive_Folder
    The path to the destination folder in OneDrive where the validated backups will be stored.
.PARAMETER SmtpCredentialPath
    The full path to the XML file containing securely stored SMTP credentials.
.PARAMETER SmtpServer
    The SMTP server address for sending notifications.
.PARAMETER SmtpPort
    The port for the SMTP server.
.PARAMETER SmtpFrom
    The sender email address for notifications.
.PARAMETER SmtpTo
    The recipient email address for notifications.
#>

#region Funciones Auxiliares
Function Write-Log {
    Param (
        [string]$Message,
        [string]$Level = "INFO" # INFO, WARN, ERROR, DEBUG
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    If ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $LogEntry
    }
    Write-Host $LogEntry
}

Function Load-EnvFile {
    Param (
        [string]$Path = (Join-Path $PSScriptRoot ".env")
    )
    $envVars = @{}
    If (Test-Path $Path) {
        Get-Content $Path | ForEach-Object {
            If ($_ -match "^\s*([^#=\s]+)\s*=\s*(.*)\s*$") {
                $key = $Matches[1]
                $value = $Matches[2].Trim("'").Trim('"').Trim() # Clean quotes and spaces
                # Expand environment variables in the value (supports %VAR% and $env:VAR)
                $expandedValue = [System.Environment]::ExpandEnvironmentVariables($value)
                Try {
                    $expandedValue = $ExecutionContext.InvokeCommand.ExpandString($expandedValue)
                }
                Catch {
                    # If expansion fails, keep the value as is
                }
                $envVars[$key] = $expandedValue
            }
        }
    }
    Else {
        Write-Log ("ERROR: Archivo .env no encontrado en '" + $Path + "'.") -Level "ERROR"
        Throw "Archivo .env no encontrado."
    }
    Return $envVars
}

Function Find-CorporateOneDriveRoot {
    [CmdletBinding()]
    Param ()

    $OneDriveRoot = $null
    $RegPath = "HKCU:\Software\Microsoft\OneDrive\Accounts"

    Try {
        $BusinessAccounts = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like "Business*" }

        If ($BusinessAccounts) {
            ForEach ($Account in $BusinessAccounts) {
                $UserFolder = Get-ItemProperty -Path "$($Account.PSPath)" -Name "UserFolder" -ErrorAction SilentlyContinue
                If ($UserFolder -and (Test-Path $UserFolder.UserFolder)) {
                    $OneDriveRoot = $UserFolder.UserFolder
                    Write-Log ("Raiz de OneDrive corporativo encontrada: '" + $OneDriveRoot + "'.")
                    Break # Found one, take the first one
                }
            }
        }
        Else {
            Write-Log "No se encontraron cuentas de OneDrive corporativas (Business*) en el registro." -Level "WARN"
        }
    }
    Catch {
        Write-Log ("ERROR al buscar la raiz de OneDrive corporativo en el registro: " + $_.Exception.Message) -Level "ERROR"
    }

    Return $OneDriveRoot
}

Function Show-UserMessage {
    Param([string]$Title, [string]$Message)
    Try {
        Add-Type -AssemblyName System.Windows.Forms
        $notification = New-Object System.Windows.Forms.NotifyIcon
        
        # Usar el icono de información del sistema
        $notification.Icon = [System.Drawing.SystemIcons]::Information
        $notification.BalloonTipIcon = "Info"
        $notification.BalloonTipTitle = $Title
        $notification.BalloonTipText = $Message
        $notification.Visible = $true
        
        # Mostrar la notificación por 5 segundos
        $notification.ShowBalloonTip(5000)
        
        # Limpiar el icono después de un momento
        Start-Sleep -Seconds 6
        $notification.Dispose()
    }
    Catch {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, "OK", "Information") | Out-Null
    }
}
#endregion

# Global error handling for critical failures
Trap {
    Write-Log "ERROR CRITICO: Una excepcion no manejada ha ocurrido." -Level "ERROR"
    Write-Log ("Mensaje: " + $_.Exception.Message) -Level "ERROR"
    Write-Log ("StackTrace: " + $_.ScriptStackTrace) -Level "ERROR"

    $EmailSubject = "Backup Automatico (Error)"
    $CurrentFileContext = $(If ($PSTFile) { "`nArchivo siendo procesado: $($PSTFile.FullName)" } Else { "" })
    $EmailBody = @"
Ha ocurrido un error crítico en el script de backup de PST.
Por favor, revisa el log para mas detalles: $script:LogFile
$CurrentFileContext

Mensaje de Error: $($_.Exception.Message)
StackTrace: $($_.ScriptStackTrace)
"@

    # Attempt to send notification even on critical failure
    Try {
        If ($script:SmtpServer -and $script:SmtpUsername -and $script:SmtpPassword) {
            $SecureSmtpCreds = New-Object System.Management.Automation.PSCredential($script:SmtpUsername, ($script:SmtpPassword | ConvertTo-SecureString -AsPlainText -Force))
            Send-MailMessage -To $script:SmtpTo -From $script:SmtpFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $script:SmtpServer -Port $script:SmtpPort -UseSSL -Credential $SecureSmtpCreds -Attachments $script:LogFile -WarningAction SilentlyContinue
            Write-Log "Notificación de error crítico enviada por correo."
        }
    }
    Catch {
        Write-Log ("Fallo al enviar notificacion de error crítico por correo: " + $_.Exception.Message) -Level "ERROR"
    }

    Exit 1 # Terminate script with an error code
}

#region 1. Configuración de Variables
Write-Host "Iniciando script de backup de PST..." -ForegroundColor Green

# Cargar variables de entorno
$envConfig = Load-EnvFile
$script:PST_Source_Dir = $envConfig["PST_SOURCE_DIR"]
$script:SmtpServer = $envConfig["SMTP_SERVER"]
$script:SmtpPort = $(If ($envConfig["SMTP_PORT"]) { [int]$envConfig["SMTP_PORT"] } Else { 587 })
$script:SmtpFrom = $envConfig["SMTP_FROM"]
$script:SmtpTo = $envConfig["SMTP_TO"]
$script:SmtpUsername = $envConfig["SMTP_USERNAME"]
$script:SmtpPassword = $envConfig["SMTP_PASSWORD"]

# 1.1 Determinar y crear la Carpeta de Transito (Staging Folder)
$script:Staging_Folder = Join-Path $env:USERPROFILE "Documents\OutlookBackupStaging"
If (-not (Test-Path $script:Staging_Folder)) {
    Write-Log ("Creando carpeta de transito: '" + $script:Staging_Folder + "'.")
    New-Item -Path $script:Staging_Folder -ItemType Directory | Out-Null
}
Else {
    Write-Log ("Carpeta de transito existente: '" + $script:Staging_Folder + "'.")
}

# 1.2 Determinar y crear la Carpeta de OneDrive
$CorporateOneDriveRoot = Find-CorporateOneDriveRoot
If (-not $CorporateOneDriveRoot) {
    Write-Log "ERROR: No se pudo detectar la raiz de OneDrive corporativo. Asegúrate de que OneDrive esta configurado y sincronizado." -Level "ERROR"
    Exit 1
}
$script:OneDrive_Folder = Join-Path $CorporateOneDriveRoot "backup_correos"
If (-not (Test-Path $script:OneDrive_Folder)) {
    Write-Log ("Creando carpeta de OneDrive: '" + $script:OneDrive_Folder + "'.")
    New-Item -Path $script:OneDrive_Folder -ItemType Directory | Out-Null
}
Else {
    Write-Log ("Carpeta de OneDrive existente: '" + $script:OneDrive_Folder + "'.")
}

# Crear carpeta de logs si no existe
$script:LogFolder = Join-Path $script:Staging_Folder "Logs"
If (-not (Test-Path $script:LogFolder)) {
    New-Item -Path $script:LogFolder -ItemType Directory | Out-Null
}
$script:LogFile = Join-Path $script:LogFolder ("OutlookPSTBackup_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

Write-Log "Variables de configuracion cargadas del archivo .env y rutas dinamicas establecidas."
Write-Log ("PST Origen Directorio: " + $script:PST_Source_Dir)
Write-Log ("Carpeta de Transito: " + $script:Staging_Folder)
Write-Log ("Carpeta de OneDrive: " + $script:OneDrive_Folder)
Write-Log ("Archivo de Log: " + $script:LogFile)

$ScriptStartTime = Get-Date

#endregion

#region 2. Protocolo de Cierre Seguro (Anti-Corrupción)
Write-Log "Verificando si Outlook está corriendo..."

# Informar al usuario
Show-UserMessage -Title "Backup de Outlook" -Message "Iniciando backup crítico. Por favor, no abra Outlook ni interrumpa el proceso."

$OutlookProcess = Get-Process -Name "outlook" -ErrorAction SilentlyContinue

If ($OutlookProcess) {
    Write-Log "Outlook.exe detectado. Intentando cierre elegante..."
    Try {
        $OutlookProcess.CloseMainWindow() | Out-Null
        Write-Log "Senal de cierre enviada a Outlook. Esperando hasta 90 segundos para que termine."

        $OutlookClosed = $false
        For ($i = 0; $i -lt 90; $i++) {
            If (-not (Get-Process -Name "outlook" -ErrorAction SilentlyContinue)) {
                $OutlookClosed = $true
                Break
            }
            Start-Sleep -Seconds 1
        }

        If (-not $OutlookClosed) {
            Write-Log "ERROR: Outlook no se cerro en el tiempo esperado (90 segundos). Abortando operacion de backup." -Level "ERROR"
            Exit 1
        }
        Write-Log "Outlook se ha cerrado éxitosamente."
        Write-Log "Esperando 20 segundos de 'tiempo de asentamiento' para que el sistema libere los archivos naturalmente..."
        Start-Sleep -Seconds 20
    }
    Catch {
        Write-Log ("ERROR al intentar cerrar Outlook elegantemente: " + $_.Exception.Message + ". Abortando.") -Level "ERROR"
        Exit 1
    }
}
Else {
    Write-Log "Outlook.exe no esta corriendo, o ya se ha cerrado. Procediendo."
}
#endregion

#region 3. Procesamiento de Múltiples PSTs
Write-Log ("Buscando archivos PST en el directorio de origen: '" + $script:PST_Source_Dir + "'...")
If (-not $script:PST_Source_Dir -or -not (Test-Path $script:PST_Source_Dir)) {
    Write-Log ("ERROR: El directorio de origen PST no existe o no esta configurado: '" + $script:PST_Source_Dir + "'") -Level "ERROR"
    Exit 1
}
$PSTFiles = Get-ChildItem -Path $script:PST_Source_Dir -Filter "*.pst" -File -Recurse

If (-not $PSTFiles) {
    Write-Log ("ADVERTENCIA: No se encontraron archivos PST en '" + $script:PST_Source_Dir + "'. Saliendo del script.") -Level "WARN"
    Exit 0
}

$BackupResults = @() # Para almacenar resultados de cada backup

ForEach ($PSTFile in $PSTFiles) {
    $PST_Source = $PSTFile.FullName
    $BaseFileName = $PSTFile.BaseName
    Write-Log ("Procesando archivo PST: '" + $PST_Source + "'")

    # --- Lógica de Nombres de Archivo con Fecha y Timestamp ---
    $Timestamp = Get-Date -Format "yyyy-MM-dd-HH"
    $BackupFileName = "$($BaseFileName)_$($Timestamp).pst"
    $StagingPSTPath = Join-Path $script:Staging_Folder $BackupFileName
    $OneDrivePSTPath = Join-Path $script:OneDrive_Folder $BackupFileName

    Write-Log ("Nombre del archivo de backup generado: " + $BackupFileName)
    Write-Log ("Ruta en Transito: " + $StagingPSTPath)
    Write-Log ("Ruta en OneDrive: " + $OneDrivePSTPath)

    #region 3.1 Fase de 'Snapshot' (Velocidad de Operacion)
    Write-Log "Iniciando fase de snapshot: copiando PST a la carpeta de transito local."
    $MaxRetries = 12
    $RetryCount = 0
    $SnapshotSuccess = $false

    While ($RetryCount -lt $MaxRetries) {
        $RetryCount++
        Try {
            Copy-Item -Path $PST_Source -Destination $StagingPSTPath -Force -ErrorAction Stop
            $SnapshotSuccess = $true
            Write-Log ("Copia local del PST completada en '" + $StagingPSTPath + "' (Intento $RetryCount).")
            Break # Exito, salir del bucle While
        }
        Catch {
            If ($RetryCount -lt $MaxRetries) {
                Write-Log ("ADVERTENCIA: Intento $RetryCount fallido para '$BaseFileName'. El archivo sigue ocupado. Reintentando en 15 segundos...") -Level "WARN"
                Start-Sleep -Seconds 15
            }
            Else {
                Write-Log ("ERROR persistente durante la fase de snapshot para '$BaseFileName' tras $MaxRetries intentos: " + $_.Exception.Message + ". Saltando este archivo.") -Level "ERROR"
                $BackupResults += [PSCustomObject]@{
                    FileName    = $BaseFileName
                    Status      = "FALLO_SNAPSHOT"
                    Details     = "Bloqueo persistente tras 3 minutos de espera: " + $_.Exception.Message
                    Hash        = "N/A"
                    SizeMB      = "N/A"
                    TimeSeconds = "N/A"
                }
            }
        }
    }

    If (-not $SnapshotSuccess) {
        Continue # Salta al siguiente archivo PST en el ForEach loop
    }
    #endregion

    #region 4. Validacion de Integridad Bit-a-Bit
    Write-Log ("Iniciando validacion de integridad (SHA256) para el archivo PST: '" + $BaseFileName + "'.")
    $OriginalHash = "N/A"
    $StagingHash = "N/A"
    Try {
        Write-Log ("Calculando SHA256 para el archivo PST original: '" + $PST_Source + "'...")
        $OriginalHash = (Get-FileHash -Path $PST_Source -Algorithm SHA256).Hash
        Write-Log ("Hash original: " + $OriginalHash)

        Write-Log ("Calculando SHA256 para la copia en transito: '" + $StagingPSTPath + "'...")
        $StagingHash = (Get-FileHash -Path $StagingPSTPath -Algorithm SHA256).Hash
        Write-Log ("Hash en transito: " + $StagingHash)

        If ($OriginalHash -ne $StagingHash) {
            Write-Log ("ERROR DE INTEGRIDAD: Los hashes SHA256 no coinciden para '" + $BaseFileName + "'. Borrando copia corrupta en transito.") -Level "ERROR"
            Remove-Item -Path $StagingPSTPath -Force -ErrorAction SilentlyContinue
            $BackupResults += [PSCustomObject]@{
                FileName    = $BaseFileName
                Status      = "FALLO_INTEGRIDAD"
                Details     = "Hashes no coinciden."
                Hash        = "N/A"
                SizeMB      = "N/A"
                TimeSeconds = "N/A"
            }
            Continue # Salta al siguiente archivo PST
        }
        Write-Log ("VALIDACION EXITOSA: Los hashes SHA256 coinciden para '" + $BaseFileName + "'.")
    }
    Catch {
        Write-Log ("ERROR durante la validacion de integridad para '" + $BaseFileName + "': " + $_.Exception.Message + ". Saltando este archivo.") -Level "ERROR"
        Remove-Item -Path $StagingPSTPath -Force -ErrorAction SilentlyContinue # Clean up
        $BackupResults += [PSCustomObject]@{
            FileName    = $BaseFileName
            Status      = "FALLO_HASH_CALC"
            Details     = $_.Exception.Message
            Hash        = "N/A"
            SizeMB      = "N/A"
            TimeSeconds = "N/A"
        }
        Continue # Salta al siguiente archivo PST
    }
    #endregion

    #region 5. Gestión de Archivos y OneDrive
    Write-Log ("Moviendo archivo validado a la carpeta de OneDrive para '" + $BaseFileName + "': '" + $script:OneDrive_Folder + "'...")
    Try {
        If (-not (Test-Path $script:OneDrive_Folder)) {
            New-Item -Path $script:OneDrive_Folder -ItemType Directory -Force | Out-Null
        }

        Move-Item -Path $StagingPSTPath -Destination $OneDrivePSTPath -Force -ErrorAction Stop
        Write-Log ("Archivo PST '" + $BaseFileName + "' movido a OneDrive: '" + $OneDrivePSTPath + "'.")

        # --- Gestión de Rotacion (Mantener solo las 2 copias mas recientes) ---
        $OldBackups = Get-ChildItem -Path $script:OneDrive_Folder -Filter "$($BaseFileName)_*.pst" | 
        Sort-Object LastWriteTime -Descending
        If ($OldBackups.Count -gt 2) {
            $FilesToDelete = $OldBackups | Select-Object -Skip 2
            ForEach ($FileToDelete in $FilesToDelete) {
                Write-Log ("Rotacion: Eliminando backup antiguo '" + $FileToDelete.Name + "'.")
                Remove-Item -Path $FileToDelete.FullName -Force
            }
        }

        $FinalFileSizeMB = ([math]::Round((Get-Item $OneDrivePSTPath).Length / 1MB, 2))
        $BackupResults += [PSCustomObject]@{
            FileName    = $BaseFileName
            Status      = "EXITO"
            Details     = "Backup completado."
            Hash        = $StagingHash
            SizeMB      = $FinalFileSizeMB
            TimeSeconds = ([math]::Round((Get-Date).Subtract($ScriptStartTime).TotalSeconds, 2))
        }
    }
    Catch {
        Write-Log ("ERROR durante la gestion de archivos y OneDrive para '" + $BaseFileName + "': " + $_.Exception.Message + ". Saltando este archivo.") -Level "ERROR"
        Remove-Item -Path $StagingPSTPath -Force -ErrorAction SilentlyContinue # Clean up
        $BackupResults += [PSCustomObject]@{
            FileName    = $BaseFileName
            Status      = "FALLO_ONEDRIVE"
            Details     = $_.Exception.Message
            Hash        = "N/A"
            SizeMB      = "N/A"
            TimeSeconds = "N/A"
        }
        Continue # Salta al siguiente archivo PST
    }
    #endregion
}

Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host ">>> IMPORTANTE: Outlook ya puede ser reabierto. <<<" -ForegroundColor Yellow
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Log "Notificación al usuario: Outlook puede ser reabierto."
#endregion

#region 6. Trazabilidad y Notificación SMTP (Final)
$ScriptEndTime = Get-Date
$ElapsedTime = ([math]::Round(($ScriptEndTime - $ScriptStartTime).TotalSeconds, 2))

Write-Log "Procesamiento de todos los archivos PST completado."
Write-Log ("Tiempo total de ejecución del script: " + $ElapsedTime + " segundos.")

$EmailSubject = "Resumen de Backup PST Outlook"
$EmailBody = "El script de backup de PST ha finalizado. Aqui esta el resumen de cada archivo:<br><br>"
$EmailBody += "<table border='1' cellpadding='5' cellspacing='0'>"
$EmailBody += "<tr><th>Archivo</th><th>Estado</th><th>Detalles</th><th>Hash SHA256</th><th>Tamano (MB)</th></tr>"

$OverallStatus = "EXITO"
ForEach ($Result in $BackupResults) {
    $EmailBody += "<tr>"
    $EmailBody += "<td>$($Result.FileName)</td>"
    $EmailBody += "<td>$($Result.Status)</td>"
    $EmailBody += "<td>$($Result.Details)</td>"
    $EmailBody += "<td>$($Result.Hash)</td>"
    $EmailBody += "<td>$($Result.SizeMB)</td>"
    $EmailBody += "</tr>"
    If ($Result.Status -ne "EXITO") {
        $OverallStatus = "FALLO_PARCIAL"
    }
}
$EmailBody += "</table><br>"

If ($OverallStatus -eq "FALLO_PARCIAL") {
    $EmailSubject = "Backup Automatico (Error)"
    $EmailBody = "ADVERTENCIA: Se detectaron uno o mas fallos durante el proceso de backup.<br><br>" + $EmailBody
}
ElseIf ($BackupResults.Count -eq 0) {
    $OverallStatus = "ADVERTENCIA"
    $EmailSubject = "Backup Automatico (Error)"
    $EmailBody = "ADVERTENCIA: No se procesaron archivos PST. Verifique el directorio de origen.<br><br>"
}
else {
    $EmailSubject = "Backup Automatico (Exito)"
}

$EmailBody += ("Tiempo total de ejecucion del script: " + $ElapsedTime + " segundos.")
$EmailBody += "<br><br>Este es un mensaje automatico. Por favor, no responder a este correo."

Try {
    If ($script:SmtpServer -and $script:SmtpUsername -and $script:SmtpPassword) {
        $SecureSmtpCreds = New-Object System.Management.Automation.PSCredential($script:SmtpUsername, ($script:SmtpPassword | ConvertTo-SecureString -AsPlainText -Force))
        Send-MailMessage -To $script:SmtpTo -From $script:SmtpFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $script:SmtpServer -Port $script:SmtpPort -UseSSL -Credential $SecureSmtpCreds -BodyAsHtml -Attachments $script:LogFile -WarningAction SilentlyContinue
        Write-Log "Notificación de resumen de backup enviada por correo."
    }
}
Catch {
    Write-Log ("ERROR: Fallo al enviar notificacion de resumen por correo: " + $_.Exception.Message) -Level "ERROR"
}

# Notificación final en pantalla
$FinalMsg = $(If ($OverallStatus -eq "EXITO") { "El backup de Outlook ha finalizado con éxito. Ya puede continuar con sus actividades normalmente." } Else { "El backup ha finalizado, pero se detectaron algunos detalles. Por favor, informele al area de TI." })
$FinalTitle = $(If ($OverallStatus -eq "EXITO") { "Backup Completado" } Else { "Backup con Advertencias" })

# Limpieza final: Eliminar carpeta de staging y logs
Write-Log "Limpiando archivos temporales..."
Remove-Item -Path $script:Staging_Folder -Recurse -Force -ErrorAction SilentlyContinue

Show-UserMessage -Title $FinalTitle -Message $FinalMsg

#endregion

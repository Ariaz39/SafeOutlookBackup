# Script de Backup de Archivos PST de Outlook a OneDrive

Este script de PowerShell automatiza el proceso de copia de seguridad de uno o varios archivos de datos de Outlook (.pst) a una carpeta de OneDrive. Está diseñado para garantizar la integridad absoluta de los datos y una trazabilidad completa mediante un sistema de registro detallado y notificaciones por correo electrónico.

## Características Principales

*   **Variables de Entorno (`.env`):** Las configuraciones se definen en un archivo `.env` externo para facilitar la portabilidad.
*   **Gestión de Versiones (Rotación):** Mantiene automáticamente las **2 copias más recientes** de cada archivo PST en OneDrive, eliminando versiones antiguas para optimizar el espacio.
*   **Nomenclatura con Timestamp:** Los archivos se guardan con el formato `NombreOriginal_AAAA-MM-DD-HH.pst` para una identificación precisa.
*   **Cierre Seguro de Outlook:** Verifica si Outlook está en ejecución e intenta un cierre elegante antes de iniciar la copia.
*   **Validación SHA256:** Calcula y compara hashes para asegurar que la copia es idéntica al original bit a bit.
*   **Notificaciones Interactivas:** Muestra alertas visuales (BalloonTips) en la barra de tareas al iniciar y finalizar el proceso.
*   **Reportes por Correo:** Envía un resumen detallado del backup y **adjunta el archivo de registro (.log)** automáticamente.
*   **Limpieza Automática:** Elimina la carpeta de tránsito (`staging`) al finalizar para no dejar archivos temporales en el equipo.

## Requisitos

*   **PowerShell 7.0 o superior:** (Recomendado 7.6.1 o superior).
*   **Microsoft Outlook:** Debe estar instalado.
*   **OneDrive Corporativo:** Configurado y sincronizando.

## Configuración

### 1. Archivo `.env`

Crea un archivo `.env` en la misma carpeta del script con los siguientes parámetros:

```ini
# Directorio de origen de los PST
PST_SOURCE_DIR="$env:USERPROFILE\Documents\Archivos de Outlook"

# Configuración SMTP
SMTP_SERVER="mail.tuempresa.com"
SMTP_PORT=587
SMTP_FROM="info@tuempresa.com"
SMTP_TO="sistemas@tuempresa.com"
SMTP_USERNAME="info@tuempresa.com"
SMTP_PASSWORD="tu_password_segura"
```

> [!IMPORTANT]
> Se recomienda usar el puerto **587** para máxima compatibilidad con PowerShell.

### 2. Carpetas de Trabajo

*   **Staging:** El script usa `%USERPROFILE%\Documents\OutlookBackupStaging` temporalmente y la elimina al terminar.
*   **OneDrive:** Los backups se guardan en la raíz detectada de OneDrive bajo la carpeta `backup_correos`.

## Ejecución

Puedes ejecutarlo manualmente o programarlo como una Tarea de Windows:

**Argumentos para la Tarea Programada:**
`powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Ruta\backup-outlook.ps1"`

## Notificaciones

El script envía un correo con el asunto:
*   `Backup Automatico (Exito)`: Si todo se procesó correctamente.
*   `Backup Automatico (Error)`: Si hubo fallos o errores críticos.

El archivo de registro detallado se envía como **adjunto** en estos correos.

---
**Nota de Codificación:** El script debe guardarse siempre como **UTF-8 con BOM** para asegurar que las tildes y caracteres especiales se muestren correctamente en las notificaciones.

# Script de Backup de Archivos PST de Outlook a OneDrive

Este script de PowerShell automatiza el proceso de copia de seguridad de uno o varios archivos de datos de Outlook (.pst) a una carpeta de OneDrive. Está diseñado para garantizar la integridad absoluta de los datos y una trazabilidad completa mediante un sistema de registro detallado y notificaciones por correo electrónico.

## Características Principales

*   **MEGA CMD Portable:** Incluye una copia local de MEGA CMD en la carpeta `bin/`, eliminando la necesidad de instalar software adicional en el equipo.
*   **Consola en Colores:** Sistema de logging visual con códigos de color (Verde para éxito, Amarillo para procesos, Rojo para errores) para un monitoreo intuitivo.
*   **Gestión de Versiones (Rotación):** Mantiene automáticamente las **2 copias más recientes** de cada archivo PST en OneDrive.
*   **Nomenclatura Multi-Equipo:** Archivos con formato `NombreEquipo_NombreOriginal_AAAA-MM-DD-HH.pst` para una identificación precisa en entornos con varios PCs.
*   **Cierre Seguro de Outlook:** Intenta un cierre elegante de Outlook y espera a que los archivos se liberen antes de copiar.
*   **Validación SHA256:** Garantiza la integridad bit a bit comparando hashes antes de la distribución.
*   **Notificaciones Visuales:** Alertas tipo BalloonTip en la barra de tareas al iniciar y finalizar.
*   **Respaldo Dual (OneDrive + MEGA):** Máxima redundancia enviando copias a ambas nubes simultáneamente.

## Requisitos

*   **PowerShell 7.0 o superior:** (Recomendado 7.6.1 o superior).
*   **Microsoft Outlook:** Debe estar instalado.
*   **OneDrive Corporativo:** Configurado y sincronizando.

## Configuración

### 1. Archivo `.env`

Crea un archivo `.env` en la raíz del proyecto. El script ignorará este archivo en Git por seguridad. Ejemplo:

```ini
# --- RUTAS ---
PST_SOURCE_DIR="C:\Users\TuUsuario\Documents\Archivos de Outlook"

# --- CONFIGURACIÓN SMTP ---
SMTP_SERVER="smtp.tuempresa.com"
SMTP_PORT=587
SMTP_FROM="backup@tuempresa.com"
SMTP_TO="admin@tuempresa.com"
SMTP_USERNAME="backup@tuempresa.com"
SMTP_PASSWORD="tu_password_smtp"

# --- CONFIGURACIÓN MEGA ---
MEGA_USER="tu_usuario_mega@email.com"
MEGA_PASS="tu_password_mega"
MEGA_REMOTE_DEST="backup_correos_v2"
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

## Tiempos Estimados de Subida

Tiempos aproximados basados en la **Velocidad de Subida (Upload)** de tu conexión (Horas : Minutos):

| Tamaño PST | 10 Mbps (Subida) | 20 Mbps (Subida) | 50 Mbps (Subida) | 100 Mbps (Subida) |
| :--- | :--- | :--- | :--- | :--- |
| **5 GB** | 1h 08m | 34 min | 14 min | 7 min |
| **10 GB** | 2h 16m | 1h 08m | 27 min | 14 min |
| **20 GB** | 4h 33m | 2h 16m | 54 min | 27 min |
| **50 GB** | 11h 22m | 5h 41m | 2h 16m | 1h 08m |

---
**Nota de Codificación:** El script debe guardarse siempre como **UTF-8 con BOM** para asegurar que las tildes y caracteres especiales se muestren correctamente en las notificaciones.

## Créditos y Licencias

*   **MEGA CMD:** Este script utiliza los binarios de [MEGA CMD](https://github.com/meganz/MEGAcmd), los cuales se distribuyen bajo la licencia **GPLv3**. MEGA es una marca registrada de MEGA Ltd.
*   **Seguridad:** Este script se proporciona "tal cual", sin garantías. El usuario es responsable de la custodia de sus credenciales en el archivo `.env`.

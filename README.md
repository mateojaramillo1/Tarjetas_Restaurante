# Control de Comedor - Lector de Tarjetas

Sistema para registrar asistencias con lector de tarjetas HID Omnikey.

---

## Cómo iniciar el programa

1. **Abre PowerShell** en la carpeta del proyecto

2. **Ejecuta este comando:**

```powershell
& ".\.venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8001
```

3. **Abre tu navegador** en:
   - Inicio de la pagina: `http://localhost:8001`
   - Pantalla de lectura: `http://localhost:8001/control`
   - Panel administrador: `http://localhost:8001/admin/login`

---

## Clave de administrador

**Clave por defecto:** `12345678`

---

## Cómo detener el programa

En la terminal donde está corriendo, presiona: **Ctrl + C**

---

## Notas importantes

- Si reinicias Windows, podrás usar el puerto `8000` en lugar de `8001`
- El lector de tarjetas debe estar conectado antes de iniciar el programa
- Los datos se guardan automáticamente en `data/card_reads.db`

---

## Reportes automaticos por correo

El sistema puede enviar reportes en Excel automaticamente:

- Cada 3 dias: reporte de control de asistencias, ordenado por fecha (por dias).
- Quincenal: reporte completo de asistencias por quincena.
   - El dia 16 envia del 1 al 15 del mes actual.
   - El dia 1 envia del 16 al fin de mes del mes anterior.

### 1) Crear configuracion de correo

1. Copia el archivo `data/report_config.example.json` a `data/report_config.json`.
2. Edita `data/report_config.json` con tu correo SMTP:

```json
{
   "enabled": true,
   "recipient_email": "mateo.jaramillo@vinus.com.co",
   "sender_email": "teojaramillosuarez@gmail.com",
   "sender_password": "fyvqygxfjmshpwov",
   "smtp_host": "smtp.gmail.com",
   "smtp_port": 587,
   "use_tls": true,
   "send_every_days": 3,
   "check_interval_minutes": 15
}
```

`send_every_days` controla cada cuantos dias se envia el reporte de control.

### 2) Iniciar la app normalmente

```powershell
& ".\.venv\Scripts\python.exe" -m uvicorn app.main:app --host 0.0.0.0 --port 8001
```

### 3) Archivos usados por el scheduler

- Configuracion: `data/report_config.json`
- Estado interno (fechas ya enviadas): `data/report_state.json`

Si cambias `enabled` a `false`, se desactiva el envio automatico.

---

## Instalacion recomendada en otro equipo (paso a paso)

La forma mas estable para evitar errores de rutas es instalar siempre en esta ruta fija:

- `C:\Apps\Lectortarjetas`

Evita instalar dentro de OneDrive para el equipo de produccion.

### 1) Ruta y copia del proyecto

1. Crea la carpeta `C:\Apps\Lectortarjetas`.
2. Copia ahi todo el proyecto.
3. Verifica que existan las carpetas `app`, `data` y el archivo `requirements.txt`.

### 2) Instalar Python y dependencias

Abre PowerShell en `C:\Apps\Lectortarjetas` y ejecuta:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 3) Probar manualmente antes del Programador

```powershell
.\.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8001
```

Prueba en navegador:

- `http://127.0.0.1:8001`
- `http://127.0.0.1:8001/api/health`

Si esto no funciona, no continúes con el Programador hasta corregirlo.

### 4) Crear tarea programada de inicio automatico

Abre PowerShell como Administrador y ejecuta este bloque completo:

```powershell
$ProjectDir = "C:\Apps\Lectortarjetas"
$TaskName = "LectorTarjetas-Autostart"
$Launcher = Join-Path $ProjectDir "run-lectortarjetas.ps1"
$LogsDir = Join-Path $ProjectDir "logs"

New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

$script = @'
Set-Location "C:\Apps\Lectortarjetas"

while ($true) {
   try {
      "$(Get-Date -Format s) START uvicorn" | Out-File -FilePath ".\\logs\\launcher.log" -Append -Encoding utf8
      & ".\\.venv\\Scripts\\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8001 >> ".\\logs\\app.log" 2>&1
      "$(Get-Date -Format s) EXIT CODE=$LASTEXITCODE" | Out-File -FilePath ".\\logs\\launcher.log" -Append -Encoding utf8
   } catch {
      "$(Get-Date -Format s) ERROR $($_.Exception.Message)" | Out-File -FilePath ".\\logs\\launcher.log" -Append -Encoding utf8
   }
   Start-Sleep -Seconds 5
}
'@

Set-Content -Path $Launcher -Value $script -Encoding UTF8

schtasks /Delete /TN "$TaskName" /F 2>$null | Out-Null

$taskCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""$Launcher"""
schtasks /Create /TN "$TaskName" /SC ONLOGON /RU "$env:USERNAME" /RL HIGHEST /TR "$taskCmd" /F

schtasks /Run /TN "$TaskName"
Start-Sleep -Seconds 5
schtasks /Query /TN "$TaskName" /V /FO LIST
```

### 5) Verificar que quedo activa

```powershell
Invoke-WebRequest -Uri "http://127.0.0.1:8001/api/health" -UseBasicParsing
Get-Content "C:\Apps\Lectortarjetas\logs\launcher.log" -Tail 20
```

Si en Programador de tareas aparece `0x41301`, significa que la tarea esta corriendo (estado esperado para este launcher en bucle).

### 6) Orden recomendado para instalar en un PC nuevo

1. Instalar Python 3.11.
2. Copiar proyecto en `C:\Apps\Lectortarjetas`.
3. Crear `.venv` e instalar `requirements.txt`.
4. Probar inicio manual con uvicorn.
5. Crear tarea programada.
6. Verificar `api/health` y logs.

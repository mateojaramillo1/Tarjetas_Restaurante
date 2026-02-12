# Control Comedor (FastAPI + SQLite)

Aplicativo web local para registrar personas por tarjeta y llevar control de almuerzos con lector HID Omnikey via PC/SC.

## Requisitos

- Windows con driver del lector Omnikey instalado.
- Python 3.10+.

## Instalacion

```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
```

## Ejecutar

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Si usas el entorno conda configurado en VS Code, puedes iniciar asi:

```bash
$env:ADMIN_PASSWORD="12345678"; $env:SESSION_SECRET="cambia-esta-clave"; C:/Users/teoja/miniforge3/Scripts/conda.exe run -p C:\Users\teoja\miniforge3 --no-capture-output python c:\Users\teoja\.vscode\extensions\ms-python.python-2026.0.0-win32-x64\python_files\get_output_via_markers.py -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```


Abre http://localhost:8000 para ver el menu principal.

## Reiniciar servidor

1) En la terminal donde esta `uvicorn`, presiona `Ctrl + C`.
2) Ejecuta de nuevo el comando de la seccion "Ejecutar".

## Apagar servidor sin terminal

En PowerShell, ejecuta:

```powershell
netstat -ano | findstr :8000
```

Luego toma el PID de la ultima columna que diga LISTENING y ejecuta:

```powershell
taskkill /PID <PID> /F
```

Ejemplo: si el PID es 12664:

```powershell
taskkill /PID 12664 /F
```

## Login administrador

- Clave por defecto: `12345678` (cambiar con variable `ADMIN_PASSWORD`).
- Para sesiones, define `SESSION_SECRET` con un valor seguro.

## Pantallas

- `/control`: pantalla publica para lectura de tarjetas (muestra datos y se limpia).
- `/admin/login`: acceso administrador.
- `/admin`: registro de personas y exportaciones.

## Notas

- La lectura de nombre del titular depende del tipo de tarjeta y la aplicacion en la tarjeta. En este proyecto se registra el UID y el ATR. Puedes implementar la lectura del nombre en `CardReaderService._read_holder_name` cuando tengas el estandar/APDUs.
- La base de datos queda en `data/card_reads.db`.
- El control diario se almacena en la tabla `attendance` y se puede exportar a Excel.

## Endpoints

- `GET /api/health`
- `GET /api/latest`
- `POST /api/people` (admin)
- `GET /api/people?limit=100` (admin)
- `GET /api/attendance?limit=100` (admin)
- `GET /api/export.xlsx` (admin)
- `GET /api/export-people.xlsx` (admin)
- `GET /api/people/search` (admin)
- `GET /api/export-people-filtered.xlsx` (admin)

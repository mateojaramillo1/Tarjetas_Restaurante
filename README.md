# Control Comedor (FastAPI + SQLite)

Aplicativo web local para registrar personas por tarjeta y llevar control de almuerzos con lector HID Omnikey via PC/SC.

## Requisitos

- Windows con driver del lector Omnikey instalado.
- Python 3.10+.

## Instalacion y primer inicio (paso a paso)

### Opcion A (recomendada): entorno virtual `venv`

1. Abre PowerShell en la carpeta del proyecto.
2. Crea el entorno virtual:

```powershell
python -m venv .venv
```

3. Activa el entorno:

```powershell
.\.venv\Scripts\Activate.ps1
```

4. Instala dependencias:

```powershell
pip install -r requirements.txt
```

5. Define variables de entorno (puedes cambiar los valores):

```powershell
$env:ADMIN_PASSWORD="12345678"
$env:SESSION_SECRET="cambia-esta-clave-segura"
```

6. Inicia el aplicativo:

```powershell
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

7. Abre en el navegador:

```text
http://localhost:8000
```

### Opcion B: entorno Conda (si ya lo usas en VS Code)

1. Activa tu entorno Conda (o usa `conda run`).
2. Define las variables de entorno:

```powershell
$env:ADMIN_PASSWORD="12345678"
$env:SESSION_SECRET="cambia-esta-clave-segura"
```

3. Ejecuta:

```powershell
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## Inicio rapido (uso diario)

Cada vez que vayas a usar el sistema:

1. Abre PowerShell en la carpeta del proyecto.
2. Activa el entorno (`.\.venv\Scripts\Activate.ps1` o tu entorno Conda).
3. Ejecuta el servidor:

```powershell
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

4. Abre `http://localhost:8000`.

## Procesos operativos (dia a dia)

### 1) Iniciar aplicativo

1. Abre PowerShell en la carpeta del proyecto.
2. Activa el entorno (`.\\.venv\\Scripts\\Activate.ps1` o Conda).
3. Define variables de entorno:

```powershell
$env:ADMIN_PASSWORD="12345678"
$env:SESSION_SECRET="cambia-esta-clave-segura"
```

4. Ejecuta:

```powershell
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

5. Abre:

```text
http://localhost:8000
```

### 2) Acceder al panel administrador

1. Entra a `http://localhost:8000/admin/login`.
2. Usa la clave definida en `ADMIN_PASSWORD`.
3. Desde `/admin` puedes registrar personas y exportar datos.

### 3) Reiniciar servidor

1) En la terminal donde esta `uvicorn`, presiona `Ctrl + C`.
2) Ejecuta de nuevo el comando de la seccion "Iniciar aplicativo".

### 4) Apagar servidor sin terminal

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

### 5) Acceder a la base de datos (SQLite)

La base se encuentra en:

```text
data/card_reads.db
```

Opciones para revisarla:

- Con herramienta grafica (recomendado): abre `data/card_reads.db` en DB Browser for SQLite.
- Con Python (sin instalar nada extra):

```powershell
python -c "import sqlite3; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); print('Tablas:', cur.execute(\"SELECT name FROM sqlite_master WHERE type='table' ORDER BY name\").fetchall()); print('People:', cur.execute('SELECT COUNT(*) FROM people').fetchone()[0]); print('Attendance:', cur.execute('SELECT COUNT(*) FROM attendance').fetchone()[0]); conn.close()"
```

### 6) Borrar registros de forma segura

Recomendado: primero haz respaldo y deten el servidor para evitar bloqueos.

1. Deten el servidor (`Ctrl + C`).
2. Crea respaldo de la BD:

```powershell
Copy-Item data\card_reads.db data\card_reads.backup.db -Force
```

3. Ejecuta el borrado que necesites.

#### Borrar asistencias de una fecha o rango

```powershell
python -c "import sqlite3; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); cur.execute(\"DELETE FROM attendance WHERE read_at >= ? AND read_at < ?\", ('2026-02-01T00:00:00', '2026-03-01T00:00:00')); conn.commit(); print('Filas borradas:', cur.rowcount); conn.close()"
```

#### Borrar una persona por UID (y sus asistencias)

```powershell
python -c "import sqlite3; uid='A1B2C3D4'; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); cur.execute('DELETE FROM attendance WHERE uid = ?', (uid,)); print('Asistencias borradas:', cur.rowcount); cur.execute('DELETE FROM people WHERE uid = ?', (uid,)); print('Personas borradas:', cur.rowcount); conn.commit(); conn.close()"
```

#### Vaciar toda la tabla de asistencias

```powershell
python -c "import sqlite3; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); cur.execute('DELETE FROM attendance'); conn.commit(); print('Asistencias borradas:', cur.rowcount); conn.close()"
```

4. Inicia nuevamente el aplicativo.

### 7) Consultas utiles de verificacion

Ver ultimos 10 registros de asistencia:

```powershell
python -c "import sqlite3; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); rows=cur.execute('SELECT uid, read_at FROM attendance ORDER BY id DESC LIMIT 10').fetchall(); [print(r) for r in rows]; conn.close()"
```

Ver personas registradas:

```powershell
python -c "import sqlite3; conn=sqlite3.connect('data/card_reads.db'); cur=conn.cursor(); rows=cur.execute('SELECT uid, first_name, last_name FROM people ORDER BY id DESC LIMIT 20').fetchall(); [print(r) for r in rows]; conn.close()"
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
- `GET /api/attendance?limit=200&from_dt=&to_dt=&name=&id_number=&area=&uid=` (admin)
- `GET /api/export.xlsx` (admin)
- `GET /api/export-people.xlsx` (admin)
- `GET /api/people/search` (admin)
- `GET /api/export-people-filtered.xlsx` (admin)

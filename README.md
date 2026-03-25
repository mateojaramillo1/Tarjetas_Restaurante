# Control de Comedor - Lector de Tarjetas

Sistema para registrar asistencias con lector de tarjetas HID Omnikey.

---

## Cómo iniciar el programa

1. **Abre PowerShell** en la carpeta del proyecto

2. **Ejecuta este comando:**

```powershell
 ".\.venv\Scripts\python.exe" -m uvicorn app.main:app --host 0.0.0.0 --port 8001
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

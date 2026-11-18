# Cobranzas

A Windows desktop app (it runs on Windows 7) that reads the receivables report ("Resumen de Cartera"), finds each client's email in the clients file, and sends a payment reminder to every client with invoices **more than 30 days** overdue.

Python 3.8 + Tkinter, packaged as a single `.exe` with PyInstaller.

## How it works

1. Greets the user and asks for the receivables file (`.xls` or `.xlsx`).
2. Reads the report rows (`CLI…`) and keeps the invoices with `Dias Mora` > 30. The amount owed is the sum of `Valor Deuda` across those invoices.
3. Matches each client against the clients file by **normalized name** (case, accents, periods/commas and extra spaces are ignored). Cut-off emails such as `ventas@empresa` are ignored.
4. Preview screen: one row per client with editable emails and total. Rows are highlighted when the client wasn't found or has no valid email.
5. **Previsualizar correo** opens the email for editing (Guardar / Cancelar). **Enviar** sends it, and **Enviar a todos** sends every ready row.
6. Every send is logged to `envios.csv`, next to the `.exe`.

## What to change and where

| What | Where |
|---|---|
| **Path to the clients file** | `config.py`, `RUTA_CLIENTES` (line 15) |
| Clients file column names | `config.py`, `COLUMNA_NOMBRE_CLIENTE`, `COLUMNA_CORREO` |
| Minimum days overdue | `config.py`, `DIAS_MORA_MINIMO` |
| Test mode (on by default) | `config.py`, `MODO_PRUEBA` / `CORREO_PRUEBA` |
| Colors, fonts, window size | `config.py`, `COLORES`, `FUENTE_*`, `TAMANO_VENTANA` |
| Logo | `assets/logo_full.png` (PNG/GIF), bundled into the `.exe`. Path and box size are in `config.py`, `LOGO_*` |
| Email subject and text | `plantilla.py`, `ASUNTO`, `CUERPO` (placeholders are listed in the file) |
| Invoice table columns | `plantilla.py`, `TABLA_COLUMNAS` |
| SMTP account / password | `credenciales.ini` next to the `.exe` (**not** in the code) |

Changes to `config.py` or `plantilla.py` require rebuilding the `.exe` (`build.bat`). Replacing the logo also requires a rebuild. `credenciales.ini` is read at runtime, so you can change it without rebuilding.

## Test mode

`MODO_PRUEBA = True` is the default. Every email goes to `CORREO_PRUEBA` (or, if that's empty, to the sender account) with `[PRUEBA]` in the subject and the real recipients at the top. Test it, then set `MODO_PRUEBA = False` and rebuild.

## Building the .exe (on Windows)

1. Install **Python 3.8.10** (the last version that supports Windows 7): https://www.python.org/downloads/release/python-3810/
   - If unsure whether the Windows 7 PC is 32 or 64-bit, use the **32-bit** installer (it works on both).
   - Windows 7 needs SP1 and update KB2533623. If the `.exe` complains about `api-ms-win-crt-*.dll`, install KB2999226.
2. Copy this folder to that PC and double-click `build.bat`. It creates a venv, installs dependencies, runs the tests and produces `dist\Cobranzas.exe`.
3. On the Windows 7 PC, put these in one folder: `Cobranzas.exe` and a filled-in `credenciales.ini`. Make sure the clients file is at `RUTA_CLIENTES`.

The simplest option is to build on the Windows 7 PC itself. A `.exe` built on Windows 10 with Python 3.8 also works.

## Development

```bash
python -m venv .venv            # Python 3.8
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
.venv/bin/python -m unittest discover -s tests
```

On Linux with a uv-installed Python, Tk needs `TCL_LIBRARY`/`TK_LIBRARY` pointed at the base install (`~/.local/share/uv/python/cpython-3.8.*/lib/tcl8.6` and `tk8.6`).

## Files

- `app.py`: GUI (welcome, preview table, editing modal, sending in a background thread).
- `lectura.py`: reads the Excel files, normalizes names, validates emails.
- `correo.py`: builds the text/table, reads credentials, sends over SMTP (TLS with certifi's CA bundle, since Windows 7's is outdated).
- `config.py`, `plantilla.py`: everything you're meant to edit.
- `tests/`: tests with synthetic data (real client data is never committed).

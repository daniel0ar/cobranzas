"""
App configuration. Edit the values here and rebuild the .exe (build.bat).

Passwords and the SMTP account are NOT here: they live in credenciales.ini,
next to the .exe (see credenciales.example.ini).
"""

# --------------------------------------------------------------------------
# Files
# --------------------------------------------------------------------------

# >>> PATH TO THE CLIENTS FILE (names + emails). Change it here. <<<
# A relative path (e.g. "clients_info.xlsx") is resolved against the folder
# where the .exe (or app.py) lives. Supports .xlsx and .xls.
RUTA_CLIENTES = r"C:\Cobranzas\clients_info.xlsx"

# Column headers in the clients file (compared ignoring case, accents,
# punctuation and extra spaces).
COLUMNA_NOMBRE_CLIENTE = "Nombre Cliente"
COLUMNA_CORREO = "E-Mail"

# Logo shown in the header. PNG or GIF, ideally ~180x60 px.
# Relative paths are resolved against the app folder, so you can just drop
# a logo.png next to the .exe without rebuilding.
LOGO_ARCHIVO = "logo.png"
LOGO_ANCHO = 180
LOGO_ALTO = 60

# Log of sent emails (CSV, opens in Excel). Relative to the app folder.
ARCHIVO_REGISTRO = "envios.csv"

# --------------------------------------------------------------------------
# Business rules
# --------------------------------------------------------------------------

# Only invoices with MORE than this many days overdue are included.
DIAS_MORA_MINIMO = 30

NOMBRE_EMPRESA = "Flexoprint"

# Formats used in the email and the preview screen.
FORMATO_FECHA = "%d/%m/%Y"
SIMBOLO_MONEDA = "$"
SEPARADOR_MILES = ","
SEPARADOR_DECIMAL = "."

# --------------------------------------------------------------------------
# Test mode
# --------------------------------------------------------------------------
# While True, NO email goes to clients: every email is sent to CORREO_PRUEBA
# (or, if empty, to the sender account in credenciales.ini) with "[PRUEBA]"
# in the subject. Set to False when you are ready to send for real.
MODO_PRUEBA = True
CORREO_PRUEBA = ""

# --------------------------------------------------------------------------
# Styles (simple light theme)
# --------------------------------------------------------------------------
TEMA_TTK = "clam"

COLORES = {
    "fondo": "#F5F6F8",          # window background
    "panel": "#FFFFFF",          # table / cards background
    "borde": "#D8DCE2",
    "texto": "#1F2933",
    "texto_suave": "#6B7280",
    "primario": "#2563EB",       # main buttons
    "primario_hover": "#1D4ED8",
    "encabezado_tabla": "#EEF1F5",
    "fila_alterna": "#FAFBFC",
    "sin_coincidencia": "#FFF1C2",  # client not found in clients file
    "sin_correo": "#FFE1DC",        # found, but no valid email
    "exito": "#15803D",
    "error": "#B91C1C",
    "aviso_prueba": "#FEF3C7",
}

FUENTE_BASE = ("Segoe UI", 10)
FUENTE_NEGRITA = ("Segoe UI", 10, "bold")
FUENTE_TITULO = ("Segoe UI", 16, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 11)
FUENTE_MONO = ("Consolas", 10)   # used in the email preview

TAMANO_VENTANA = "1200x700"

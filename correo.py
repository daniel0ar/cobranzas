"""Building the email text and sending it over SMTP."""

import configparser
import datetime
import os
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from typing import List, Optional, Tuple

import config
import plantilla
from lectura import ClienteDeuda

ARCHIVO_CREDENCIALES = "credenciales.ini"

CREDENCIALES_EJEMPLO = """\
; SMTP account used to send the emails.
; This file stays OUTSIDE the .exe and the repository: keep it next to Cobranzas.exe.
;
; Gmail / Google Workspace: servidor = smtp.gmail.com, puerto = 587, seguridad = starttls
;   contrasena = an "App Password" (Google account > Security > App passwords)
; Office 365 / Outlook.com: servidor = smtp.office365.com, puerto = 587, seguridad = starttls
; seguridad: starttls | ssl | ninguna   (ssl usually goes with port 465)

[smtp]
servidor = smtp.gmail.com
puerto = 587
seguridad = starttls
usuario = cobranzas@tuempresa.com
contrasena =
remitente_nombre = Cobranzas Flexoprint
remitente_correo = cobranzas@tuempresa.com
; Optional: hidden copy of every email (e.g. to keep a record). Leave empty if not needed.
copia_oculta =
"""


class ErrorCorreo(Exception):
    """Error with a user-facing message (in Spanish)."""


# --------------------------------------------------------------------------
# Formatting
# --------------------------------------------------------------------------

def formato_moneda(valor) -> str:
    texto = "{:,.2f}".format(valor)  # 1,234.56
    texto = texto.replace(",", "\0").replace(".", config.SEPARADOR_DECIMAL).replace("\0", config.SEPARADOR_MILES)
    return config.SIMBOLO_MONEDA + texto


def formato_fecha(fecha: Optional[datetime.date]) -> str:
    return fecha.strftime(config.FORMATO_FECHA) if fecha else ""


def tabla_facturas(cliente: ClienteDeuda, total) -> str:
    """Plain-text table, aligned for monospace fonts and separated with '|'
    so it stays readable in mail clients that use proportional fonts."""
    filas = [
        ("{} {}".format(f.tipo, f.numero).strip(), formato_fecha(f.emision), formato_fecha(f.vencimiento),
         str(f.dias_mora), formato_moneda(f.valor_deuda))
        for f in cliente.facturas
    ]
    encabezado = plantilla.TABLA_COLUMNAS
    anchos = [max(len(encabezado[i]), *(len(fila[i]) for fila in filas)) if filas else len(encabezado[i])
              for i in range(len(encabezado))]
    derecha = (False, False, False, True, True)  # right-align numbers

    def linea(valores):
        celdas = [v.rjust(anchos[i]) if derecha[i] else v.ljust(anchos[i]) for i, v in enumerate(valores)]
        return " | ".join(celdas).rstrip()

    separador = "-+-".join("-" * a for a in anchos)
    lineas = [linea(encabezado), separador] + [linea(f) for f in filas] + [separador]
    etiqueta = plantilla.TABLA_ETIQUETA_TOTAL + ":"
    monto = formato_moneda(total)
    ancho_total = len(separador)
    lineas.append(etiqueta + monto.rjust(max(ancho_total - len(etiqueta), len(monto) + 1)))
    return "\n".join(lineas)


def _rellenar(texto: str, valores: dict) -> str:
    # Plain replace (not str.format) so stray braces in the template don't break it.
    for clave, valor in valores.items():
        texto = texto.replace("{" + clave + "}", valor)
    return texto


def generar_correo(cliente: ClienteDeuda, total) -> Tuple[str, str]:
    """Returns (subject, body) from the default template."""
    valores = {
        "cliente": cliente.nombre,
        "empresa": config.NOMBRE_EMPRESA,
        "total": formato_moneda(total),
        "num_facturas": str(len(cliente.facturas)),
        "dias_minimo": str(config.DIAS_MORA_MINIMO),
        "tabla_facturas": tabla_facturas(cliente, total),
        "fecha": formato_fecha(datetime.date.today()),
    }
    return _rellenar(plantilla.ASUNTO, valores), _rellenar(plantilla.CUERPO, valores)


# --------------------------------------------------------------------------
# Credentials and sending
# --------------------------------------------------------------------------

def cargar_credenciales(carpeta: str) -> dict:
    """Read credenciales.ini from the app folder. If it doesn't exist, create
    it from the example so the user only has to fill it in."""
    ruta = os.path.join(carpeta, ARCHIVO_CREDENCIALES)
    if not os.path.isfile(ruta):
        try:
            with open(ruta, "w", encoding="utf-8") as archivo:
                archivo.write(CREDENCIALES_EJEMPLO)
            creado = "Se creó un archivo de ejemplo; complételo con los datos de la cuenta de correo."
        except OSError:
            creado = "Cree el archivo a partir de credenciales.example.ini."
        raise ErrorCorreo("No se encontró el archivo de credenciales:\n{}\n\n{}".format(ruta, creado))

    lector = configparser.ConfigParser(interpolation=None)
    try:
        lector.read(ruta, encoding="utf-8")
        seccion = lector["smtp"]
        datos = {
            "servidor": seccion.get("servidor", "").strip(),
            "puerto": seccion.getint("puerto", 587),
            "seguridad": seccion.get("seguridad", "starttls").strip().lower(),
            "usuario": seccion.get("usuario", "").strip(),
            "contrasena": seccion.get("contrasena", ""),
            "remitente_nombre": seccion.get("remitente_nombre", "").strip(),
            "remitente_correo": seccion.get("remitente_correo", "").strip(),
            "copia_oculta": seccion.get("copia_oculta", "").strip(),
        }
    except (configparser.Error, KeyError, ValueError) as exc:
        raise ErrorCorreo("El archivo de credenciales tiene un error:\n{}\n\nDetalle: {}".format(ruta, exc))

    faltantes = [c for c in ("servidor", "usuario", "contrasena", "remitente_correo") if not datos[c]]
    if faltantes:
        raise ErrorCorreo("Faltan datos en el archivo de credenciales ({}):\n{}".format(", ".join(faltantes), ruta))
    if datos["seguridad"] not in ("starttls", "ssl", "ninguna"):
        raise ErrorCorreo("En credenciales.ini, 'seguridad' debe ser starttls, ssl o ninguna.")
    return datos


def _contexto_ssl():
    # Use certifi's CA bundle: Windows 7's certificate store is often outdated
    # and causes "CERTIFICATE_VERIFY_FAILED".
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


class ServidorCorreo:
    """SMTP connection reused for several sends:
        with ServidorCorreo(credenciales) as servidor:
            servidor.enviar([...], subject, body)
    """

    def __init__(self, credenciales: dict):
        self.cred = credenciales
        self.smtp = None

    def __enter__(self):
        c = self.cred
        try:
            if c["seguridad"] == "ssl":
                self.smtp = smtplib.SMTP_SSL(c["servidor"], c["puerto"], timeout=30, context=_contexto_ssl())
            else:
                self.smtp = smtplib.SMTP(c["servidor"], c["puerto"], timeout=30)
                if c["seguridad"] == "starttls":
                    self.smtp.starttls(context=_contexto_ssl())
            self.smtp.login(c["usuario"], c["contrasena"])
        except smtplib.SMTPAuthenticationError:
            self._cerrar()
            raise ErrorCorreo("Usuario o contraseña de correo incorrectos (revise credenciales.ini).\n"
                              "En Gmail se necesita una \"contraseña de aplicación\".")
        except Exception as exc:
            self._cerrar()
            raise ErrorCorreo("No se pudo conectar al servidor de correo {}:{}.\n\nDetalle: {}".format(
                c["servidor"], c["puerto"], exc))
        return self

    def __exit__(self, *args):
        self._cerrar()

    def _cerrar(self):
        if self.smtp is not None:
            try:
                self.smtp.quit()
            except Exception:
                pass
            self.smtp = None

    def enviar(self, destinatarios: List[str], asunto: str, cuerpo: str) -> None:
        c = self.cred
        mensaje = EmailMessage()
        mensaje["From"] = formataddr((c["remitente_nombre"], c["remitente_correo"]))
        mensaje["To"] = ", ".join(destinatarios)
        mensaje["Subject"] = asunto
        mensaje["Date"] = formatdate(localtime=True)
        mensaje["Message-ID"] = make_msgid(domain=c["remitente_correo"].split("@")[-1])
        mensaje.set_content(cuerpo)
        todos = list(destinatarios)
        if c["copia_oculta"]:
            todos.append(c["copia_oculta"])  # BCC: delivered, not in headers
        try:
            self.smtp.send_message(mensaje, from_addr=c["remitente_correo"], to_addrs=todos)
        except smtplib.SMTPRecipientsRefused:
            raise ErrorCorreo("El servidor rechazó la dirección: {}".format(", ".join(destinatarios)))
        except Exception as exc:
            raise ErrorCorreo("Error al enviar: {}".format(exc))

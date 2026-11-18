"""Reading the Excel files: clients file and receivables ("cartera") report."""

import datetime
import os
import re
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional, Tuple

CENTAVO = Decimal("0.01")

# Simple email check: something@domain.tld. Rejects cut-off addresses such as
# "milempro@cordiller" (no TLD) or "ventas@" (no domain).
_RE_CORREO = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$")
_RE_SEPARADORES_CORREO = re.compile(r"[;,\s]+")


class ErrorLectura(Exception):
    """Error with a user-facing message (in Spanish)."""


@dataclass
class Factura:
    tipo: str
    numero: str
    dias_mora: int
    emision: Optional[datetime.date]
    vencimiento: Optional[datetime.date]
    comentario: str
    valor_deuda: Decimal


@dataclass
class ClienteDeuda:
    codigo: str
    nombre: str
    facturas: List[Factura] = field(default_factory=list)
    correos: List[str] = field(default_factory=list)
    encontrado: bool = False                  # matched in the clients file
    correos_ignorados: List[str] = field(default_factory=list)

    @property
    def total(self) -> Decimal:
        return sum((f.valor_deuda for f in self.facturas), Decimal("0")).quantize(CENTAVO)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def normalizar(texto) -> str:
    """Normalize a name for matching: uppercase, no accents, no '.'/',',
    other symbols become spaces, single spaces.
    "Proaño  Carrión S.A." -> "PROANO CARRION SA"."""
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper().replace(".", "").replace(",", "")
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return texto.strip()


def separar_correos(texto) -> Tuple[List[str], List[str]]:
    """Split a cell with one or more emails. Returns (valid, invalid)."""
    validos, invalidos, vistos = [], [], set()
    for parte in _RE_SEPARADORES_CORREO.split(str(texto or "")):
        parte = parte.strip()
        if not parte:
            continue
        if _RE_CORREO.match(parte):
            if parte.lower() not in vistos:
                vistos.add(parte.lower())
                validos.append(parte)
        else:
            invalidos.append(parte)
    return validos, invalidos


def a_decimal(valor) -> Optional[Decimal]:
    """Excel number or text ("1,234.56", "1234,56", "$ 10") -> Decimal."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return Decimal(str(valor)).quantize(CENTAVO)
    texto = str(valor).replace("$", "").replace(" ", "").strip()
    if not texto:
        return None
    if "," in texto and "." in texto:
        # The last separator is the decimal one.
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return Decimal(texto).quantize(CENTAVO)
    except InvalidOperation:
        return None


def _a_fecha(valor) -> Optional[datetime.date]:
    if isinstance(valor, datetime.datetime):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    if isinstance(valor, str) and valor.strip():
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%Y.%m.%d"):
            try:
                return datetime.datetime.strptime(valor.strip(), formato).date()
            except ValueError:
                pass
    return None


def _a_texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def leer_filas(ruta: str) -> List[List[List[object]]]:
    """Read every sheet of an .xls/.xlsx. Returns [sheet][row][column] with
    dates as datetime, numbers as float/int and text as str."""
    if not os.path.isfile(ruta):
        raise ErrorLectura("No se encontró el archivo:\n{}".format(ruta))
    extension = os.path.splitext(ruta)[1].lower()
    try:
        if extension == ".xls":
            return _leer_xls(ruta)
        if extension in (".xlsx", ".xlsm"):
            return _leer_xlsx(ruta)
    except ErrorLectura:
        raise
    except Exception as exc:  # corrupt file, open in another program, etc.
        raise ErrorLectura("No se pudo leer el archivo:\n{}\n\nDetalle: {}".format(ruta, exc))
    raise ErrorLectura("Formato no soportado ({}). Use un archivo .xls o .xlsx.".format(extension))


def _leer_xls(ruta):
    import xlrd
    libro = xlrd.open_workbook(ruta)
    hojas = []
    for hoja in libro.sheets():
        filas = []
        for r in range(hoja.nrows):
            fila = []
            for c in range(hoja.ncols):
                tipo = hoja.cell_type(r, c)
                valor = hoja.cell_value(r, c)
                if tipo == xlrd.XL_CELL_DATE:
                    valor = xlrd.xldate_as_datetime(valor, libro.datemode)
                elif tipo in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    valor = None
                fila.append(valor)
            filas.append(fila)
        hojas.append(filas)
    return hojas


def _leer_xlsx(ruta):
    import openpyxl
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        return [[list(fila) for fila in hoja.iter_rows(values_only=True)] for hoja in libro.worksheets]
    finally:
        libro.close()


def _buscar_encabezado(filas, requeridos, max_filas=40):
    """Find the header row containing every normalized name in `requeridos`.
    Returns (row_index, {normalized_name: column}) or None."""
    for i, fila in enumerate(filas[:max_filas]):
        columnas = {}
        for c, valor in enumerate(fila):
            nombre = normalizar(valor)
            if nombre and nombre not in columnas:
                columnas[nombre] = c
        if all(r in columnas for r in requeridos):
            return i, columnas
    return None


# --------------------------------------------------------------------------
# Clients file
# --------------------------------------------------------------------------

def leer_clientes(ruta: str, columna_nombre: str, columna_correo: str) -> Dict[str, Tuple[List[str], List[str]]]:
    """Returns {normalized_name: (valid_emails, discarded_emails)}."""
    col_nombre, col_correo = normalizar(columna_nombre), normalizar(columna_correo)
    for filas in leer_filas(ruta):
        encontrado = _buscar_encabezado(filas, (col_nombre, col_correo))
        if not encontrado:
            continue
        inicio, columnas = encontrado
        i_nombre, i_correo = columnas[col_nombre], columnas[col_correo]
        clientes = {}
        for fila in filas[inicio + 1:]:
            nombre = normalizar(fila[i_nombre] if i_nombre < len(fila) else None)
            if not nombre:
                continue
            validos, invalidos = separar_correos(fila[i_correo] if i_correo < len(fila) else None)
            previos_v, previos_i = clientes.get(nombre, ([], []))
            # Same client listed twice: merge emails.
            for correo in validos:
                if correo.lower() not in (p.lower() for p in previos_v):
                    previos_v.append(correo)
            previos_i.extend(invalidos)
            clientes[nombre] = (previos_v, previos_i)
        return clientes
    raise ErrorLectura(
        "El archivo de clientes no tiene las columnas \"{}\" y \"{}\".\n{}".format(
            columna_nombre, columna_correo, ruta))


# --------------------------------------------------------------------------
# Receivables report ("Resumen de Cartera")
# --------------------------------------------------------------------------

_COLUMNAS_CARTERA = {
    "tipo": "TIPO",
    "numero": "NUMERO",
    "dias_mora": "DIAS MORA",
    "emision": "EMISION",
    "vencimiento": "VENCIMIENTO",
    "comentario": "COMENTARIO",
    "valor_deuda": "VALOR DEUDA",
}


def leer_cartera(ruta: str, dias_minimo: int, hoy: Optional[datetime.date] = None) -> Tuple[List[ClienteDeuda], int]:
    """Read the receivables report and return the clients with invoices more
    than `dias_minimo` days overdue (only those invoices), plus the total
    number of invoice rows read.

    Expected layout: report header, then a row starting with "Cliente" with
    the column headers; the client name is in the column right after
    "Cliente". "Total" rows, group rows and blank rows are skipped."""
    hoy = hoy or datetime.date.today()
    requeridos = ("CLIENTE", "DIAS MORA", "VALOR DEUDA")
    for filas in leer_filas(ruta):
        encontrado = _buscar_encabezado(filas, requeridos)
        if not encontrado:
            continue
        inicio, columnas = encontrado
        i_codigo = columnas["CLIENTE"]
        i_nombre = i_codigo + 1
        idx = {clave: columnas.get(nombre) for clave, nombre in _COLUMNAS_CARTERA.items()}

        def celda(fila, i):
            return fila[i] if i is not None and i < len(fila) else None

        clientes = {}  # code -> ClienteDeuda, keeps report order
        leidas = 0
        for fila in filas[inicio + 1:]:
            codigo = _a_texto(celda(fila, i_codigo))
            valor = a_decimal(celda(fila, idx["valor_deuda"]))
            if not codigo or valor is None:
                continue  # blank, group ("Clientes Locales") or "Total" row
            vencimiento = _a_fecha(celda(fila, idx["vencimiento"]))
            dias = a_decimal(celda(fila, idx["dias_mora"]))
            if dias is None and vencimiento is not None:
                dias = Decimal((hoy - vencimiento).days)
            if dias is None:
                continue
            leidas += 1
            if int(dias) <= dias_minimo or valor <= 0:
                continue
            nombre = _a_texto(celda(fila, i_nombre))
            cliente = clientes.get(codigo)
            if cliente is None:
                cliente = clientes[codigo] = ClienteDeuda(codigo=codigo, nombre=" ".join(nombre.split()))
            cliente.facturas.append(Factura(
                tipo=_a_texto(celda(fila, idx["tipo"])),
                numero=_a_texto(celda(fila, idx["numero"])),
                dias_mora=int(dias),
                emision=_a_fecha(celda(fila, idx["emision"])),
                vencimiento=vencimiento,
                comentario=_a_texto(celda(fila, idx["comentario"])),
                valor_deuda=valor,
            ))
        return list(clientes.values()), leidas
    raise ErrorLectura(
        "No se reconoce el formato del archivo de cuentas por cobrar.\n"
        "Se esperaba una fila de encabezados con las columnas \"Cliente\", "
        "\"Dias Mora\" y \"Valor Deuda\".\n\n{}".format(ruta))


def asignar_correos(deudores: List[ClienteDeuda], clientes: Dict[str, Tuple[List[str], List[str]]]) -> None:
    """Fill in emails by matching on the normalized name."""
    for deudor in deudores:
        datos = clientes.get(normalizar(deudor.nombre))
        if datos is not None:
            deudor.encontrado = True
            deudor.correos = list(datos[0])
            deudor.correos_ignorados = list(datos[1])

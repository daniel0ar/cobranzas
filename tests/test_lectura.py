"""Tests with synthetic data (no real client data in the repo).

Run:  python -m unittest discover -s tests
"""

import datetime
import os
import sys
import tempfile
import unittest
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from correo import generar_correo  # noqa: E402
from lectura import (a_decimal, asignar_correos, leer_cartera, leer_clientes, normalizar,  # noqa: E402
                     separar_correos)

D = datetime.datetime


def _guardar(filas):
    libro = openpyxl.Workbook()
    hoja = libro.active
    for fila in filas:
        hoja.append(fila)
    archivo = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    archivo.close()
    libro.save(archivo.name)
    return archivo.name


# Same layout as the ERP's "Resumen de Cartera" report.
CARTERA = [
    [None], ["Resúmen de Cartera"], ["Empresa:", "EMPRESA DEMO"], ["Desde:", "01/01/2025"],
    [None],
    ["Cliente", None, "Tipo", "Numero", "Dias Mora", "Emisión", "Vencimiento", "Comentario", None,
     "Total  Facturado", "Valor Deuda", "Saldo Actual"],
    [None],
    ["Clientes Locales"],
    ["CLI001", "ACME  QUÍMICA  S.A.", "FA", 100, 75, D(2026, 8, 4), D(2026, 9, 3), "LISTA 1", None, 434.13, 425.45, 425.45],
    [None, None, None, None, None, None, None, None, "Total", 434.13, 425.45, 425.45],
    ["CLI002", "JUAN PÉREZ", "FA", 200, 305, D(2026, 1, 16), D(2026, 1, 16), None, None, 402.93, 10.01, 10.01],
    ["CLI002", "JUAN PÉREZ", "FA", 201, 134, D(2026, 7, 6), D(2026, 7, 6), None, None, 387.09, 78.84, 88.85],
    ["CLI002", "JUAN PÉREZ", "FA", 202, 30, D(2026, 10, 18), D(2026, 10, 18), None, None, 50.00, 50.00, 138.85],
    [None, None, None, None, None, None, None, None, "Total", 840.02, 138.85, 138.85],
    ["CLI003", "AL DIA CIA LTDA", "FA", 300, 10, D(2026, 11, 7), D(2026, 11, 7), None, None, 99.0, 99.0, 99.0],
    ["CLI004", "SIN REGISTRO SA", "DC", 2, 654, D(2025, 2, 1), D(2025, 2, 1), None, None, 834.73, 834.73, 834.73],
    [None, None, None, None, None, None, None, None, "TOTAL GENERAL", 1.0, 1.0, None],
]

CLIENTES = [
    ["Fecha Registro", "Nombre Cliente ", "Contacto", "RUC", "E-Mail"],
    ["2020.03.10", "Acme Química S.A.", "X", "1", "pagos@acme.com.ec;ventas@acme"],
    ["2020.03.10", "Juan  Perez", "X", "2", "juan@correo.com"],
    ["2020.03.10", "Al Dia Cia. Ltda.", "X", "3", None],
]


class TestUtilidades(unittest.TestCase):

    def test_normalizar(self):
        self.assertEqual(normalizar("Proaño  Carrión S.A. "), "PROANO CARRION SA")
        self.assertEqual(normalizar("LABORATORIOS DR.  A. BJARNER C.A."), normalizar("Laboratorios Dr. A. Bjarner CA"))
        self.assertEqual(normalizar(None), "")

    def test_separar_correos_ignora_cortados(self):
        validos, invalidos = separar_correos("a@b.com.ec;milempro@cordiller, A@B.com.ec ; c@d.co")
        self.assertEqual(validos, ["a@b.com.ec", "c@d.co"])
        self.assertEqual(invalidos, ["milempro@cordiller"])

    def test_a_decimal(self):
        self.assertEqual(a_decimal(425.45), Decimal("425.45"))
        self.assertEqual(a_decimal("1,234.56"), Decimal("1234.56"))
        self.assertEqual(a_decimal("1.234,56"), Decimal("1234.56"))
        self.assertEqual(a_decimal("$ 10,5"), Decimal("10.50"))
        self.assertIsNone(a_decimal("abc"))
        self.assertIsNone(a_decimal(""))


class TestArchivos(unittest.TestCase):

    def setUp(self):
        self.ruta_cartera = _guardar(CARTERA)
        self.ruta_clientes = _guardar(CLIENTES)

    def tearDown(self):
        os.remove(self.ruta_cartera)
        os.remove(self.ruta_clientes)

    def test_cartera_filtra_mas_de_30_dias_y_suma_valor_deuda(self):
        deudores, leidas = leer_cartera(self.ruta_cartera, 30)
        self.assertEqual(leidas, 6)
        por_codigo = {d.codigo: d for d in deudores}
        self.assertEqual(list(por_codigo), ["CLI001", "CLI002", "CLI004"])  # CLI003 is not overdue
        self.assertEqual(por_codigo["CLI002"].total, Decimal("88.85"))     # exactly 30 days is excluded
        self.assertEqual(len(por_codigo["CLI002"].facturas), 2)
        self.assertEqual(por_codigo["CLI001"].nombre, "ACME QUÍMICA S.A.")

    def test_coincidencia_por_nombre_normalizado(self):
        deudores, _ = leer_cartera(self.ruta_cartera, 30)
        asignar_correos(deudores, leer_clientes(self.ruta_clientes, "Nombre Cliente", "E-Mail"))
        por_codigo = {d.codigo: d for d in deudores}
        self.assertEqual(por_codigo["CLI001"].correos, ["pagos@acme.com.ec"])
        self.assertEqual(por_codigo["CLI001"].correos_ignorados, ["ventas@acme"])
        self.assertEqual(por_codigo["CLI002"].correos, ["juan@correo.com"])
        self.assertFalse(por_codigo["CLI004"].encontrado)

    def test_correo_incluye_total_y_tabla(self):
        deudores, _ = leer_cartera(self.ruta_cartera, 30)
        juan = [d for d in deudores if d.codigo == "CLI002"][0]
        asunto, cuerpo = generar_correo(juan, juan.total)
        self.assertIn("JUAN PÉREZ", asunto)
        self.assertIn("$88.85", cuerpo)
        self.assertIn("FA 200", cuerpo)
        self.assertIn("FA 201", cuerpo)
        self.assertNotIn("FA 202", cuerpo)
        self.assertNotIn("{", cuerpo)


if __name__ == "__main__":
    unittest.main()

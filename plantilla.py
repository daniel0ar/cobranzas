"""
Default email template (plain text). Edit and rebuild the .exe.

Available placeholders (write them with braces, exactly as shown):
  {cliente}         client name as it appears in the receivables report
  {empresa}         NOMBRE_EMPRESA from config.py
  {total}           total owed (sum of overdue invoices, or the value edited
                    in the preview screen), e.g. "$1,234.56"
  {num_facturas}    number of overdue invoices
  {dias_minimo}     DIAS_MORA_MINIMO from config.py
  {tabla_facturas}  table with the overdue invoices and the total
  {fecha}           today's date
"""

ASUNTO = "Recordatorio de pago - Facturas vencidas - {cliente}"

CUERPO = """\
Estimado(a) {cliente}:

Reciba un cordial saludo de parte de {empresa}.

Le escribimos para recordarle que, según nuestros registros al {fecha}, mantiene {num_facturas} factura(s) con más de {dias_minimo} días de vencimiento, por un valor total de {total}.

A continuación, el detalle:

{tabla_facturas}

Le agradeceremos realizar el pago a la brevedad posible. Si ya realizó el pago, por favor envíenos el comprobante respondiendo a este correo y haga caso omiso de este mensaje.

Para cualquier consulta, no dude en contactarnos.

Atentamente,

Departamento de Cobranzas
{empresa}
"""

# Invoice table column headers.
TABLA_COLUMNAS = ("Documento", "Emisión", "Vencimiento", "Días vencida", "Valor pendiente")
TABLA_ETIQUETA_TOTAL = "TOTAL ADEUDADO"

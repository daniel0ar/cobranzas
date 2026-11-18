"""
Cobranzas: sends payment reminders to clients with overdue invoices.

Run:    python app.py
Build:  build.bat  (creates dist\\Cobranzas.exe)
"""

import csv
import datetime
import math
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import config
from correo import ErrorCorreo, ServidorCorreo, cargar_credenciales, formato_moneda, generar_correo
from lectura import ErrorLectura, a_decimal, asignar_correos, leer_cartera, leer_clientes, separar_correos

C = config.COLORES


def carpeta_app() -> str:
    """Folder of the .exe (when packaged) or of this file (when run with python)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def ruta_en_app(ruta: str) -> str:
    return ruta if os.path.isabs(ruta) else os.path.join(carpeta_app(), ruta)


def ruta_recurso(ruta: str) -> str:
    """File bundled with the app (e.g. assets/): inside the .exe when packaged, else next to this file."""
    if os.path.isabs(ruta):
        return ruta
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, ruta)


def escalar_imagen(imagen, ancho_max, alto_max):
    """Shrink a PhotoImage to fit the box, using the best zoom/subsample fraction (Tk has no resize)."""
    escala = min(ancho_max / imagen.width(), alto_max / imagen.height())
    if escala >= 1:
        return imagen
    mejor = (1, math.ceil(1 / escala))   # plain subsample, always fits
    for divisor in range(2, 21):
        multiplo = int(escala * divisor)
        if multiplo >= 1 and multiplo / divisor > mejor[0] / mejor[1]:
            mejor = (multiplo, divisor)
    multiplo, divisor = mejor
    return (imagen.zoom(multiplo) if multiplo > 1 else imagen).subsample(divisor)


class Fila:
    """One client in the preview table (data + widgets)."""

    def __init__(self, cliente):
        self.cliente = cliente
        self.var_correos = tk.StringVar(value="; ".join(cliente.correos))
        self.var_total = tk.StringVar(value="{:.2f}".format(cliente.total))
        self.asunto_editado = None   # set when the user saves the email in the preview
        self.cuerpo_editado = None
        self.enviado = False
        self.lbl_estado = None
        self.btn_previsualizar = None
        self.btn_enviar = None

    @property
    def estado_inicial(self):
        """Status based on what is currently typed in the email field."""
        validos, invalidos = separar_correos(self.var_correos.get())
        if invalidos:
            return "Correo no válido: {}".format(", ".join(invalidos)), C["error"]
        if validos:
            return "Listo para enviar", C["texto"]
        if not self.cliente.encontrado:
            return "No encontrado: ingrese el correo", C["error"]
        return "Sin correo válido: ingréselo", C["error"]

    def poner_estado(self, texto, color):
        if self.cuerpo_editado is not None and not self.enviado:
            texto += " (editado)"
        self.lbl_estado.configure(text=texto, fg=color)


class MarcoDesplazable(tk.Frame):
    """Frame with a vertical scrollbar; widgets go into `self.interior`."""

    def __init__(self, padre, bg):
        super().__init__(padre, bg=bg, highlightthickness=1, highlightbackground=C["borde"])
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        barra = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.interior = tk.Frame(self.canvas, bg=bg)
        ventana = self.canvas.create_window((0, 0), window=self.interior, anchor="nw")
        self.interior.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(ventana, width=e.width))
        self.canvas.configure(yscrollcommand=barra.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        for evento in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(evento, self._rueda, add="+")

    def _rueda(self, evento):
        try:
            widget = self.winfo_containing(evento.x_root, evento.y_root)
        except (KeyError, tk.TclError):
            return
        if widget is None or not str(widget).startswith(str(self.canvas)):
            return
        if evento.num == 4:
            paso = -1
        elif evento.num == 5:
            paso = 1
        else:
            paso = -int(evento.delta / 120) or (-1 if evento.delta > 0 else 1)
        self.canvas.yview_scroll(paso, "units")


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Cobranzas - {}".format(config.NOMBRE_EMPRESA))
        self.geometry(config.TAMANO_VENTANA)
        self.minsize(950, 560)
        self.configure(bg=C["fondo"])
        self.report_callback_exception = self._error_inesperado
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

        self.filas = []
        self.cola = queue.Queue()
        self.enviando = False
        self.en_curso = {}
        self.resultados = []
        self.botones_globales = []
        self._logo = None

        self._estilos()
        self._encabezado()
        if config.MODO_PRUEBA:
            tk.Label(self, text="MODO PRUEBA activo: los correos NO se envían a los clientes, sino a la cuenta "
                                "de prueba. (Desactivar en config.py: MODO_PRUEBA = False)",
                     bg=C["aviso_prueba"], fg=C["texto"], font=config.FUENTE_BASE, pady=6).pack(fill="x")
        self.contenido = tk.Frame(self, bg=C["fondo"])
        self.contenido.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        self.mostrar_bienvenida()

    # ------------------------------------------------------------------ UI base

    def _estilos(self):
        estilo = ttk.Style(self)
        if config.TEMA_TTK in estilo.theme_names():
            estilo.theme_use(config.TEMA_TTK)
        estilo.configure("TButton", font=config.FUENTE_BASE, padding=(10, 4))
        estilo.configure("Primario.TButton", font=config.FUENTE_NEGRITA, padding=(14, 6),
                         background=C["primario"], foreground="#FFFFFF", bordercolor=C["primario"])
        estilo.map("Primario.TButton",
                   background=[("disabled", C["borde"]), ("active", C["primario_hover"])],
                   foreground=[("disabled", C["texto_suave"])])
        estilo.configure("Chico.TButton", font=config.FUENTE_BASE, padding=(6, 2))

    def _encabezado(self):
        barra = tk.Frame(self, bg=C["panel"], highlightthickness=1, highlightbackground=C["borde"])
        barra.pack(fill="x")
        # Logo area (fixed size). The logo is bundled from assets/ (see config.LOGO_ARCHIVO).
        marco_logo = tk.Frame(barra, width=config.LOGO_ANCHO, height=config.LOGO_ALTO, bg=C["panel"])
        marco_logo.pack_propagate(False)
        marco_logo.pack(side="left", padx=20, pady=10)
        ruta_logo = ruta_recurso(config.LOGO_ARCHIVO)
        if os.path.isfile(ruta_logo):
            try:
                self._logo = escalar_imagen(tk.PhotoImage(file=ruta_logo), config.LOGO_ANCHO, config.LOGO_ALTO)
            except tk.TclError:
                self._logo = None
        if self._logo is not None:
            tk.Label(marco_logo, image=self._logo, bg=C["panel"]).pack(expand=True)

        ttk.Button(barra, text="Salir", command=self._cerrar).pack(side="right", padx=20, pady=10, anchor="n")

        textos = tk.Frame(barra, bg=C["panel"])
        textos.pack(side="left", pady=10)
        tk.Label(textos, text="Asistente de Cobranzas", bg=C["panel"], fg=C["texto"],
                 font=config.FUENTE_TITULO).pack(anchor="w")
        tk.Label(textos, text=config.NOMBRE_EMPRESA, bg=C["panel"], fg=C["texto_suave"],
                 font=config.FUENTE_SUBTITULO).pack(anchor="w")

    def _limpiar(self):
        for widget in self.contenido.winfo_children():
            widget.destroy()
        self.filas = []
        self.botones_globales = []

    def _error_inesperado(self, tipo, valor, tb):
        detalle = "".join(traceback.format_exception(tipo, valor, tb))
        messagebox.showerror("Error inesperado", "Ocurrió un error inesperado:\n\n{}\n\n{}".format(
            valor, detalle[-1500:]), parent=self)

    def _cerrar(self):
        if self.enviando and not messagebox.askyesno(
                "Envío en curso", "Se están enviando correos. ¿Desea salir de todos modos?", parent=self):
            return
        self.destroy()

    # ------------------------------------------------------------------ Welcome screen

    def mostrar_bienvenida(self):
        self._limpiar()
        hora = datetime.datetime.now().hour
        saludo = "Buenos días" if hora < 12 else ("Buenas tardes" if hora < 19 else "Buenas noches")

        tarjeta = tk.Frame(self.contenido, bg=C["panel"], highlightthickness=1,
                           highlightbackground=C["borde"], padx=40, pady=30)
        tarjeta.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(tarjeta, text="¡{}!".format(saludo), bg=C["panel"], fg=C["texto"],
                 font=config.FUENTE_TITULO).pack(anchor="w")
        tk.Label(tarjeta, bg=C["panel"], fg=C["texto"], font=config.FUENTE_SUBTITULO, justify="left",
                 wraplength=560,
                 text="Bienvenido(a) al asistente de cobranzas. Este programa revisa las cuentas por cobrar "
                      "y prepara un correo para cada cliente con facturas vencidas por más de {} días."
                      .format(config.DIAS_MORA_MINIMO)).pack(anchor="w", pady=(6, 20))
        tk.Label(tarjeta, text="Por favor, selecciona el archivo de cuentas por cobrar", bg=C["panel"],
                 fg=C["texto"], font=config.FUENTE_NEGRITA).pack(anchor="w")
        ttk.Button(tarjeta, text="Seleccionar archivo...", style="Primario.TButton",
                   command=self.seleccionar_archivo).pack(anchor="w", pady=(10, 20))

        ruta_clientes = ruta_en_app(config.RUTA_CLIENTES)
        existe = os.path.isfile(ruta_clientes)
        tk.Label(tarjeta, bg=C["panel"], fg=C["texto_suave"] if existe else C["error"],
                 font=config.FUENTE_BASE, justify="left", wraplength=560,
                 text="Base de clientes: {}{}".format(
                     ruta_clientes, "" if existe else "\nNo se encontró el archivo. Revise RUTA_CLIENTES en config.py.")
                 ).pack(anchor="w")

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            parent=self, title="Por favor, selecciona el archivo de cuentas por cobrar",
            filetypes=[("Archivos de Excel", "*.xls *.xlsx"), ("Todos los archivos", "*.*")])
        if not ruta:
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            try:
                clientes = leer_clientes(ruta_en_app(config.RUTA_CLIENTES),
                                         config.COLUMNA_NOMBRE_CLIENTE, config.COLUMNA_CORREO)
            except ErrorLectura as exc:
                messagebox.showerror("Archivo de clientes", "Error en el archivo de clientes "
                                     "(RUTA_CLIENTES en config.py).\n\n{}".format(exc), parent=self)
                return
            try:
                deudores, leidas = leer_cartera(ruta, config.DIAS_MORA_MINIMO)
            except ErrorLectura as exc:
                messagebox.showerror("Cuentas por cobrar", str(exc), parent=self)
                return
        finally:
            self.configure(cursor="")

        asignar_correos(deudores, clientes)
        if not deudores:
            messagebox.showinfo("Sin facturas vencidas",
                                "Se revisaron {} facturas y ninguna tiene más de {} días de vencimiento."
                                .format(leidas, config.DIAS_MORA_MINIMO), parent=self)
            return
        self.mostrar_previsualizacion(ruta, deudores, leidas)

    # ------------------------------------------------------------------ Preview screen

    def mostrar_previsualizacion(self, ruta, deudores, leidas):
        self._limpiar()
        self.filas = [Fila(d) for d in deudores]

        no_encontrados = sum(1 for d in deudores if not d.encontrado)
        sin_correo = sum(1 for d in deudores if d.encontrado and not d.correos)
        ignorados = sum(len(d.correos_ignorados) for d in deudores)
        total = sum((d.total for d in deudores), 0)

        # Summary + global buttons
        arriba = tk.Frame(self.contenido, bg=C["fondo"])
        arriba.pack(fill="x", pady=(0, 10))
        resumen = tk.Frame(arriba, bg=C["fondo"])
        resumen.pack(side="left", fill="x", expand=True)
        tk.Label(resumen, text="Archivo: {}".format(os.path.basename(ruta)), bg=C["fondo"], fg=C["texto_suave"],
                 font=config.FUENTE_BASE).pack(anchor="w")
        tk.Label(resumen, bg=C["fondo"], fg=C["texto"], font=config.FUENTE_NEGRITA,
                 text="{} clientes con facturas vencidas por más de {} días  ·  Total: {}  ·  ({} facturas revisadas)"
                 .format(len(deudores), config.DIAS_MORA_MINIMO, formato_moneda(total), leidas)).pack(anchor="w")
        leyenda = tk.Frame(resumen, bg=C["fondo"])
        leyenda.pack(anchor="w", pady=(4, 0))
        self._muestra_color(leyenda, C["sin_coincidencia"],
                            "No encontrado en la base de clientes ({})".format(no_encontrados))
        self._muestra_color(leyenda, C["sin_correo"], "Encontrado, pero sin correo válido ({})".format(sin_correo))
        if ignorados:
            tk.Label(leyenda, text="  ·  {} correo(s) incompleto(s) ignorado(s)".format(ignorados),
                     bg=C["fondo"], fg=C["texto_suave"], font=config.FUENTE_BASE).pack(side="left")

        botones = tk.Frame(arriba, bg=C["fondo"])
        botones.pack(side="right", anchor="n")
        b_otro = ttk.Button(botones, text="Cargar otro archivo", command=self._cargar_otro)
        b_otro.pack(side="left", padx=(0, 8))
        b_todos = ttk.Button(botones, text="Enviar a todos", style="Primario.TButton", command=self.enviar_todos)
        b_todos.pack(side="left")
        self.botones_globales = [b_otro, b_todos]

        # Table
        tabla = MarcoDesplazable(self.contenido, C["panel"])
        tabla.pack(fill="both", expand=True)
        interior = tabla.interior
        encabezados = ("Cliente", "Correos (separados por ;)", "Facturas", "Total", "Estado", "Acciones")
        for col, texto in enumerate(encabezados):
            tk.Label(interior, text=texto, bg=C["encabezado_tabla"], fg=C["texto"], font=config.FUENTE_NEGRITA,
                     anchor="w", padx=8, pady=6).grid(row=0, column=col, sticky="nsew")
        interior.grid_columnconfigure(0, weight=2, minsize=220)
        interior.grid_columnconfigure(1, weight=3, minsize=260)
        interior.grid_columnconfigure(4, minsize=200)

        for i, fila in enumerate(self.filas, start=1):
            self._construir_fila(interior, i, fila)

    def _muestra_color(self, padre, color, texto):
        tk.Label(padre, text="    ", bg=color, highlightthickness=1,
                 highlightbackground=C["borde"]).pack(side="left", padx=(0, 4))
        tk.Label(padre, text=texto, bg=C["fondo"], fg=C["texto"], font=config.FUENTE_BASE).pack(side="left",
                                                                                                padx=(0, 14))

    def _construir_fila(self, padre, r, fila):
        cliente = fila.cliente
        if not cliente.encontrado:
            fondo = C["sin_coincidencia"]
        elif not cliente.correos:
            fondo = C["sin_correo"]
        else:
            fondo = C["fila_alterna"] if r % 2 == 0 else C["panel"]

        def celda(col):
            marco = tk.Frame(padre, bg=fondo, padx=8, pady=6)
            marco.grid(row=r, column=col, sticky="nsew")
            return marco

        tk.Label(celda(0), text=cliente.nombre, bg=fondo, fg=C["texto"], font=config.FUENTE_BASE,
                 anchor="w", justify="left", wraplength=260).pack(fill="x")

        entrada = tk.Entry(celda(1), textvariable=fila.var_correos, font=config.FUENTE_BASE, relief="solid", bd=1,
                           bg="#FFFFFF", fg=C["texto"])
        entrada.pack(fill="x", ipady=2)

        tk.Label(celda(2), text=str(len(cliente.facturas)), bg=fondo, fg=C["texto"],
                 font=config.FUENTE_BASE).pack()

        tk.Entry(celda(3), textvariable=fila.var_total, font=config.FUENTE_BASE, relief="solid", bd=1,
                 width=11, justify="right", bg="#FFFFFF", fg=C["texto"]).pack(ipady=2)

        fila.lbl_estado = tk.Label(celda(4), bg=fondo, font=config.FUENTE_BASE, anchor="w", justify="left",
                                   wraplength=200)
        fila.lbl_estado.pack(fill="x")
        fila.poner_estado(*fila.estado_inicial)

        def al_editar_correo(*_):
            if not fila.enviado and not self.enviando:
                fila.poner_estado(*fila.estado_inicial)
        fila.var_correos.trace_add("write", al_editar_correo)

        acciones = celda(5)
        fila.btn_previsualizar = ttk.Button(acciones, text="Previsualizar correo", style="Chico.TButton",
                                            command=lambda: self.previsualizar(fila))
        fila.btn_previsualizar.pack(side="left", padx=(0, 6))
        fila.btn_enviar = ttk.Button(acciones, text="Enviar", style="Chico.TButton",
                                     command=lambda: self.enviar_uno(fila))
        fila.btn_enviar.pack(side="left")

    def _cargar_otro(self):
        if any(f.cuerpo_editado is not None and not f.enviado for f in self.filas) and not messagebox.askyesno(
                "Cargar otro archivo", "Hay correos editados sin enviar que se perderán. ¿Continuar?", parent=self):
            return
        self.mostrar_bienvenida()

    # ------------------------------------------------------------------ Email text per row

    def _total(self, fila):
        total = a_decimal(fila.var_total.get())
        if total is None or total <= 0:
            raise ValueError("El total de {} no es válido: \"{}\"".format(fila.cliente.nombre, fila.var_total.get()))
        return total

    def _texto(self, fila):
        """(subject, body): the version edited by the user, or the template's."""
        if fila.cuerpo_editado is not None:
            return fila.asunto_editado, fila.cuerpo_editado
        return generar_correo(fila.cliente, self._total(fila))

    def _destinatarios(self, fila):
        validos, invalidos = separar_correos(fila.var_correos.get())
        if invalidos:
            raise ValueError("Correo(s) no válido(s) para {}: {}".format(fila.cliente.nombre, ", ".join(invalidos)))
        if not validos:
            raise ValueError("{} no tiene correo. Ingréselo en la tabla.".format(fila.cliente.nombre))
        return validos

    def previsualizar(self, fila):
        try:
            asunto, cuerpo = self._texto(fila)
        except ValueError as exc:
            messagebox.showerror("Previsualizar correo", str(exc), parent=self)
            return
        para = fila.var_correos.get().strip() or "(sin correo: ingréselo en la tabla)"

        v = tk.Toplevel(self)
        v.title("Previsualizar correo - {}".format(fila.cliente.nombre))
        v.configure(bg=C["fondo"], padx=16, pady=16)
        v.geometry("860x640")
        v.minsize(600, 400)
        v.transient(self)
        v.columnconfigure(1, weight=1)
        v.rowconfigure(3, weight=1)

        tk.Label(v, text="Para:", bg=C["fondo"], fg=C["texto"], font=config.FUENTE_NEGRITA).grid(
            row=0, column=0, sticky="w", pady=(0, 6))
        tk.Label(v, text=para, bg=C["fondo"], fg=C["texto"], font=config.FUENTE_BASE, anchor="w").grid(
            row=0, column=1, sticky="we", pady=(0, 6))
        if config.MODO_PRUEBA:
            tk.Label(v, text="MODO PRUEBA: este correo se enviará a la cuenta de prueba, no al cliente.",
                     bg=C["aviso_prueba"], fg=C["texto"], font=config.FUENTE_BASE, anchor="w", padx=6).grid(
                row=1, column=0, columnspan=2, sticky="we", pady=(0, 6))
        tk.Label(v, text="Asunto:", bg=C["fondo"], fg=C["texto"], font=config.FUENTE_NEGRITA).grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        ent_asunto = tk.Entry(v, font=config.FUENTE_BASE, relief="solid", bd=1)
        ent_asunto.insert(0, asunto)
        ent_asunto.grid(row=2, column=1, sticky="we", pady=(0, 8), ipady=3)

        marco_texto = tk.Frame(v, bg=C["borde"], bd=0)
        marco_texto.grid(row=3, column=0, columnspan=2, sticky="nsew")
        texto = tk.Text(marco_texto, font=config.FUENTE_MONO, wrap="word", relief="flat", padx=10, pady=10,
                        bg=C["panel"], fg=C["texto"], undo=True)
        barra = ttk.Scrollbar(marco_texto, orient="vertical", command=texto.yview)
        texto.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        texto.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        texto.insert("1.0", cuerpo)

        def guardar():
            fila.asunto_editado = ent_asunto.get().strip()
            fila.cuerpo_editado = texto.get("1.0", "end-1c")
            if not fila.enviado:
                fila.poner_estado(*fila.estado_inicial)
            v.destroy()

        botones = tk.Frame(v, bg=C["fondo"])
        botones.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botones, text="Cancelar", command=v.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(botones, text="Guardar", style="Primario.TButton", command=guardar).pack(side="left")
        v.bind("<Escape>", lambda e: v.destroy())

        def capturar():
            try:
                v.grab_set()  # modal: blocks the main window until closed
            except tk.TclError:
                v.after(50, capturar)  # window not visible yet
        capturar()
        texto.focus_set()

    # ------------------------------------------------------------------ Sending

    def _preparar(self, fila, credenciales):
        """Validate the row and return the job (recipients, subject, body, total). Raises ValueError."""
        destinatarios = self._destinatarios(fila)
        total = self._total(fila)
        asunto, cuerpo = self._texto(fila)
        if not asunto:
            raise ValueError("El asunto del correo para {} está vacío.".format(fila.cliente.nombre))
        if config.MODO_PRUEBA:
            prueba = config.CORREO_PRUEBA or credenciales["remitente_correo"]
            cuerpo = "[PRUEBA - destinatarios reales: {}]\n\n{}".format(", ".join(destinatarios), cuerpo)
            asunto = "[PRUEBA] " + asunto
            destinatarios = [prueba]
        return destinatarios, asunto, cuerpo, total

    def _credenciales(self):
        try:
            return cargar_credenciales(carpeta_app())
        except ErrorCorreo as exc:
            messagebox.showerror("Credenciales de correo", str(exc), parent=self)
            return None

    def enviar_uno(self, fila):
        credenciales = self._credenciales()
        if credenciales is None:
            return
        try:
            trabajo = self._preparar(fila, credenciales)
        except ValueError as exc:
            messagebox.showerror("Enviar", str(exc), parent=self)
            return
        aviso = "\n\n(MODO PRUEBA: se enviará a {})".format(trabajo[0][0]) if config.MODO_PRUEBA else ""
        if not messagebox.askyesno("Confirmar envío", "¿Enviar el correo a {}?\n\nPara: {}\nTotal: {}{}".format(
                fila.cliente.nombre, fila.var_correos.get(), formato_moneda(trabajo[3]), aviso), parent=self):
            return
        self._iniciar_envio([(fila, trabajo)], credenciales)

    def enviar_todos(self):
        credenciales = self._credenciales()
        if credenciales is None:
            return
        listos, omitidos = [], []
        for fila in self.filas:
            if fila.enviado:
                continue
            try:
                listos.append((fila, self._preparar(fila, credenciales)))
            except ValueError as exc:
                omitidos.append(str(exc))
        if not listos:
            messagebox.showinfo("Enviar a todos", "No hay correos listos para enviar.\n\n" + "\n".join(omitidos[:15]),
                                parent=self)
            return
        mensaje = "Se enviarán {} correo(s).".format(len(listos))
        if omitidos:
            mensaje += "\n\nSe omitirán {} cliente(s):\n- {}".format(len(omitidos), "\n- ".join(omitidos[:10]))
            if len(omitidos) > 10:
                mensaje += "\n... y {} más".format(len(omitidos) - 10)
        if config.MODO_PRUEBA:
            mensaje += "\n\n(MODO PRUEBA: todos irán a {})".format(listos[0][1][0][0])
        if messagebox.askyesno("Confirmar envío", mensaje + "\n\n¿Continuar?", parent=self):
            self._iniciar_envio(listos, credenciales)

    def _bloquear(self, bloquear):
        estado = ["disabled"] if bloquear else ["!disabled"]
        for boton in self.botones_globales:
            boton.state(estado)
        for fila in self.filas:
            fila.btn_previsualizar.state(estado)
            fila.btn_enviar.state(["disabled"] if bloquear or fila.enviado else ["!disabled"])

    def _iniciar_envio(self, trabajos, credenciales):
        self.enviando = True
        self.resultados = []
        self.en_curso = {}
        datos = []
        for indice, (fila, (destinatarios, asunto, cuerpo, total)) in enumerate(trabajos):
            self.en_curso[indice] = (fila, destinatarios, total)
            datos.append((indice, destinatarios, asunto, cuerpo))
            fila.poner_estado("Enviando...", C["texto_suave"])
        self._bloquear(True)
        self.configure(cursor="watch")
        threading.Thread(target=self._trabajador, args=(datos, credenciales), daemon=True).start()
        self.after(100, self._revisar_cola)

    def _trabajador(self, datos, credenciales):
        """Runs in a separate thread so the window doesn't freeze. Doesn't touch Tk."""
        try:
            with ServidorCorreo(credenciales) as servidor:
                for indice, destinatarios, asunto, cuerpo in datos:
                    try:
                        servidor.enviar(destinatarios, asunto, cuerpo)
                        self.cola.put(("ok", indice, ""))
                    except ErrorCorreo as exc:
                        self.cola.put(("error", indice, str(exc)))
        except ErrorCorreo as exc:
            self.cola.put(("conexion", None, str(exc)))
        except Exception as exc:
            self.cola.put(("conexion", None, "Error inesperado: {}".format(exc)))
        finally:
            self.cola.put(("fin", None, ""))

    def _revisar_cola(self):
        try:
            while True:
                tipo, indice, detalle = self.cola.get_nowait()
                if tipo == "fin":
                    self._fin_envio()
                    return
                if tipo == "conexion":
                    self.resultados.append(("conexion", None, detalle))
                    continue
                fila, destinatarios, total = self.en_curso.pop(indice)
                if tipo == "ok":
                    if config.MODO_PRUEBA:
                        fila.poner_estado("Prueba enviada", C["exito"])
                    else:
                        fila.enviado = True
                        fila.poner_estado("Enviado", C["exito"])
                else:
                    fila.poner_estado("Error al enviar", C["error"])
                self.resultados.append((tipo, fila, detalle))
                self._registrar(fila, destinatarios, total, tipo, detalle)
        except queue.Empty:
            pass
        self.after(100, self._revisar_cola)

    def _fin_envio(self):
        # Rows still pending (connection failed before sending them).
        for fila, destinatarios, total in self.en_curso.values():
            fila.poner_estado("No enviado", C["error"])
        self.en_curso = {}
        self.enviando = False
        self.configure(cursor="")
        self._bloquear(False)

        enviados = sum(1 for r in self.resultados if r[0] == "ok")
        errores = ["- {}: {}".format(r[1].cliente.nombre, r[2]) for r in self.resultados if r[0] == "error"]
        conexion = [r[2] for r in self.resultados if r[0] == "conexion"]
        if conexion:
            messagebox.showerror("Error de conexión", "\n\n".join(conexion), parent=self)
            return
        mensaje = "Correos enviados: {}".format(enviados)
        if errores:
            mensaje += "\nCon error: {}\n\n{}".format(len(errores), "\n".join(errores[:10]))
            messagebox.showwarning("Envío terminado", mensaje, parent=self)
        else:
            messagebox.showinfo("Envío terminado", mensaje, parent=self)

    def _registrar(self, fila, destinatarios, total, resultado, detalle):
        """Append the result to envios.csv (opens in Excel)."""
        ruta = ruta_en_app(config.ARCHIVO_REGISTRO)
        nuevo = not os.path.isfile(ruta)
        try:
            # BOM only when creating the file, so Excel shows accents correctly.
            with open(ruta, "a", newline="", encoding="utf-8-sig" if nuevo else "utf-8") as archivo:
                escritor = csv.writer(archivo, delimiter=";")
                if nuevo:
                    escritor.writerow(["fecha_hora", "modo", "codigo", "cliente", "destinatarios", "facturas",
                                       "total", "resultado", "detalle"])
                escritor.writerow([
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "PRUEBA" if config.MODO_PRUEBA else "REAL",
                    fila.cliente.codigo, fila.cliente.nombre, "; ".join(destinatarios),
                    len(fila.cliente.facturas), "{:.2f}".format(total),
                    "enviado" if resultado == "ok" else "error", detalle,
                ])
        except OSError:
            pass  # e.g. envios.csv open in Excel: logging must not stop the sending


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

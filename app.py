"""Interfaz de escritorio del Calculador de Kilometraje (Windows y Mac).

Pensada para trabajo repetitivo: se cargan varios Excel (o una carpeta completa),
un botón los calcula todos con las mismas fórmulas de `procesador` y cada
resultado se guarda solo, sin diálogos intermedios.
"""
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

import procesador

TITULO = "Calculador de Kilometraje y Rendimiento"
RUTA_CONFIG = Path.home() / ".calculador_km.json"
EXTENSIONES = (".xlsx", ".xlsm", ".xls")
SUFIJO_RESULTADO = " - RESULTADO"
MAX_FILAS_VISTA = 500
FUENTE_MONO = "Menlo" if sys.platform == "darwin" else "Consolas"
GRIS = "#888888"  # se lee tanto en modo claro como en oscuro

COLUMNAS_ARCHIVOS = [
    ("archivo", "Archivo", 250, "w"),
    ("estado", "Estado", 110, "w"),
    ("tickets", "Tickets", 70, "e"),
    ("folios", "Folios", 60, "e"),
    ("km", "Km recorridos", 110, "e"),
    ("falt_lts", "Faltante (L)", 100, "e"),
    ("falt_pesos", "Faltante ($)", 110, "e"),
    ("revisar", "Revisar", 65, "e"),
]
VISTAS = ["Resumen por folio", "Detalle por día", "Revisar"]


def abrir_con_sistema(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta], check=False)
    else:
        subprocess.run(["xdg-open", ruta], check=False)


def mostrar_en_carpeta(ruta):
    if sys.platform == "win32":
        subprocess.run(["explorer", "/select,", str(ruta)], check=False)
    elif sys.platform == "darwin":
        subprocess.run(["open", "-R", str(ruta)], check=False)
    else:
        abrir_con_sistema(str(Path(ruta).parent))


def es_excel_de_entrada(ruta):
    """Excel válido para procesar: no es temporal de Excel ni un resultado ya generado."""
    ruta = Path(ruta)
    return (ruta.suffix.lower() in EXTENSIONES
            and not ruta.name.startswith("~$")
            and not ruta.stem.endswith(SUFIJO_RESULTADO))


def formatear(valor):
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ""
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, bool):
        return "Sí" if valor else ""
    if isinstance(valor, float):
        return f"{valor:.0f}" if valor.is_integer() else f"{valor:,.2f}"
    return str(valor)


def cargar_config():
    try:
        return json.loads(RUTA_CONFIG.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_config(config):
    try:
        RUTA_CONFIG.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass  # La configuración es una comodidad; no debe interrumpir el trabajo.


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("1040x720")
        self.minsize(860, 600)

        self.config_usuario = cargar_config()
        self.archivos = {}  # ruta (str) -> {"resultado", "salida", "error", "meta"}
        self.cola = queue.Queue()
        self.trabajando = False

        self.meta = tk.StringVar(value=str(self.config_usuario.get("meta", procesador.META_DEFAULT)))
        self.carpeta_salida = tk.StringVar(value=self.config_usuario.get("carpeta_salida", ""))
        self.abrir_al_terminar = tk.BooleanVar(value=self.config_usuario.get("abrir_al_terminar", False))

        marco = ttk.Frame(self, padding=16)
        marco.pack(fill="both", expand=True)
        self._crear_encabezado(marco)
        self._crear_barra_acciones(marco)
        self._crear_opciones(marco)

        self._crear_pie(marco)  # antes que el panel: así conserva su lugar al achicar la ventana

        panel = ttk.Panedwindow(marco, orient="vertical")
        panel.pack(fill="both", expand=True, pady=(10, 0))
        panel.add(self._crear_lista_archivos(panel), weight=1)
        panel.add(self._crear_vistas(panel), weight=2)
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self._actualizar_botones()

    # ---------- construcción de la ventana ----------

    def _crear_encabezado(self, padre):
        ttk.Label(padre, text=TITULO, font=("", 17, "bold")).pack(anchor="w")
        ttk.Label(padre, foreground=GRIS,
                  text="Agrega uno o varios Excel de cargas de diésel y presiona «Calcular todo». "
                       "Cada resultado se guarda solo como «nombre - RESULTADO.xlsx».").pack(anchor="w", pady=(2, 12))

    def _crear_barra_acciones(self, padre):
        fila = ttk.Frame(padre)
        fila.pack(fill="x")
        self.btn_agregar = ttk.Button(fila, text="📄  Agregar Excel…", command=self.agregar_archivos)
        self.btn_agregar.pack(side="left")
        self.btn_carpeta = ttk.Button(fila, text="📂  Agregar carpeta…", command=self.agregar_carpeta)
        self.btn_carpeta.pack(side="left", padx=6)
        self.btn_quitar = ttk.Button(fila, text="Quitar", command=self.quitar_seleccion)
        self.btn_quitar.pack(side="left")
        self.btn_limpiar = ttk.Button(fila, text="Limpiar lista", command=self.limpiar)
        self.btn_limpiar.pack(side="left", padx=6)
        self.btn_calcular = ttk.Button(fila, text="▶  Calcular todo", command=self.calcular)
        self.btn_calcular.pack(side="right")

    def _crear_opciones(self, padre):
        caja = ttk.LabelFrame(padre, text="Opciones (se recuerdan para la próxima vez)", padding=10)
        caja.pack(fill="x", pady=(10, 0))

        ttk.Label(caja, text="Rendimiento meta por defecto:").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(caja, from_=0.5, to=10, increment=0.1, width=6, textvariable=self.meta,
                    format="%.2f").grid(row=0, column=1, sticky="w", padx=(6, 4))
        ttk.Label(caja, text="km/L  (para tractos que no están en el catálogo)",
                  foreground=GRIS).grid(row=0, column=2, sticky="w")

        ttk.Label(caja, text="Guardar resultados en:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.lbl_salida = ttk.Label(caja)
        self.lbl_salida.grid(row=1, column=1, columnspan=2, sticky="w", padx=(6, 0), pady=(8, 0))
        ttk.Button(caja, text="Cambiar…", command=self.elegir_carpeta_salida).grid(row=1, column=3, padx=6, pady=(8, 0))
        ttk.Button(caja, text="Junto al original", command=self.salida_junto_al_original).grid(row=1, column=4, pady=(8, 0))

        ttk.Checkbutton(caja, text="Abrir los resultados en Excel al terminar",
                        variable=self.abrir_al_terminar).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        caja.columnconfigure(2, weight=1)
        self._refrescar_etiqueta_salida()

    def _crear_lista_archivos(self, padre):
        caja = ttk.Frame(padre)
        self.tabla = ttk.Treeview(caja, columns=[c[0] for c in COLUMNAS_ARCHIVOS], show="headings",
                                  selectmode="extended", height=5)
        for clave, titulo, ancho, alineacion in COLUMNAS_ARCHIVOS:
            self.tabla.heading(clave, text=titulo, anchor=alineacion)
            self.tabla.column(clave, width=ancho, anchor=alineacion, stretch=(clave == "archivo"))
        self.tabla.tag_configure("error", foreground="#E5484D")
        self.tabla.tag_configure("listo", foreground="#2E9E44")
        barra = ttk.Scrollbar(caja, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self._mostrar_seleccion())
        self.tabla.bind("<Double-1>", lambda _e: self.abrir_resultado())
        return caja

    def _crear_vistas(self, padre):
        self.pestanas = ttk.Notebook(padre)
        self.txt = tk.Text(self.pestanas, wrap="word", relief="flat", padx=10, pady=10, height=8,
                           font=(FUENTE_MONO, 11))
        self.txt.configure(state="disabled")
        self.pestanas.add(self.txt, text="Resumen")

        self.vistas = {}
        for nombre in VISTAS:
            caja = ttk.Frame(self.pestanas)
            arbol = ttk.Treeview(caja, show="headings", height=8)
            arbol.tag_configure("faltante", background="#F8D7DA", foreground="#000")
            vertical = ttk.Scrollbar(caja, orient="vertical", command=arbol.yview)
            horizontal = ttk.Scrollbar(caja, orient="horizontal", command=arbol.xview)
            arbol.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
            arbol.grid(row=0, column=0, sticky="nsew")
            vertical.grid(row=0, column=1, sticky="ns")
            horizontal.grid(row=1, column=0, sticky="ew")
            caja.rowconfigure(0, weight=1)
            caja.columnconfigure(0, weight=1)
            self.pestanas.add(caja, text=nombre)
            self.vistas[nombre] = arbol
        self._escribir("Agrega un Excel para empezar.")
        return self.pestanas

    def _crear_pie(self, padre):
        fila = ttk.Frame(padre)
        fila.pack(side="bottom", fill="x")
        self.barra = ttk.Progressbar(padre, mode="determinate")
        self.barra.pack(side="bottom", fill="x", pady=(10, 6))
        self.lbl_estado = ttk.Label(fila, text="")
        self.lbl_estado.pack(side="left")
        self.btn_ver_carpeta = ttk.Button(fila, text="Mostrar en carpeta", command=self.mostrar_resultado_en_carpeta)
        self.btn_ver_carpeta.pack(side="right")
        self.btn_ver = ttk.Button(fila, text="Abrir resultado en Excel", command=self.abrir_resultado)
        self.btn_ver.pack(side="right", padx=6)

    # ---------- lista de archivos ----------

    def agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona uno o varios Excel", initialdir=self.config_usuario.get("ultima_carpeta"),
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if rutas:
            self.config_usuario["ultima_carpeta"] = str(Path(rutas[0]).parent)
            self._agregar(rutas)

    def agregar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los Excel",
                                          initialdir=self.config_usuario.get("ultima_carpeta"))
        if not carpeta:
            return
        self.config_usuario["ultima_carpeta"] = carpeta
        rutas = sorted(p for p in Path(carpeta).iterdir() if p.is_file() and es_excel_de_entrada(p))
        if not rutas:
            messagebox.showinfo("Sin archivos", "En esa carpeta no hay Excel para procesar "
                                                "(los que terminan en «- RESULTADO» se ignoran).")
            return
        self._agregar(rutas)

    def _agregar(self, rutas):
        nuevos = 0
        for ruta in rutas:
            clave = str(Path(ruta))
            if clave in self.archivos:
                continue
            self.archivos[clave] = {"resultado": None, "salida": None, "error": None, "meta": None}
            self.tabla.insert("", "end", iid=clave, values=(Path(clave).name, "Pendiente"))
            nuevos += 1
        if nuevos and not self.tabla.selection():
            self.tabla.selection_set(self.tabla.get_children()[0])
        self._actualizar_botones()
        self._actualizar_estado()

    def quitar_seleccion(self):
        for clave in self.tabla.selection():
            self.tabla.delete(clave)
            self.archivos.pop(clave, None)
        self._mostrar_seleccion()
        self._actualizar_botones()
        self._actualizar_estado()

    def limpiar(self):
        self.tabla.delete(*self.tabla.get_children())
        self.archivos.clear()
        self._mostrar_seleccion()
        self._actualizar_botones()
        self._actualizar_estado()

    # ---------- opciones ----------

    def elegir_carpeta_salida(self):
        carpeta = filedialog.askdirectory(title="Carpeta donde guardar los resultados",
                                          initialdir=self.carpeta_salida.get() or self.config_usuario.get("ultima_carpeta"))
        if carpeta:
            self.carpeta_salida.set(carpeta)
            self._refrescar_etiqueta_salida()

    def salida_junto_al_original(self):
        self.carpeta_salida.set("")
        self._refrescar_etiqueta_salida()

    def _refrescar_etiqueta_salida(self):
        self.lbl_salida.configure(text=self.carpeta_salida.get() or "La misma carpeta del Excel original")

    def _leer_meta(self):
        try:
            meta = float(self.meta.get().replace(",", "."))
        except ValueError:
            meta = 0
        if meta <= 0:
            messagebox.showerror("Meta inválida", "El rendimiento meta por defecto debe ser un número mayor que 0 (ej. 2.5).")
            return None
        return meta

    def _ruta_salida(self, entrada):
        sugerida = procesador.ruta_salida_para(entrada)
        carpeta = self.carpeta_salida.get()
        return Path(carpeta) / sugerida.name if carpeta else sugerida

    # ---------- cálculo ----------

    def calcular(self):
        if self.trabajando or not self.archivos:
            return
        meta = self._leer_meta()
        if meta is None:
            return
        # Solo lo que falta; si ya está todo calculado con esta meta, se recalcula todo.
        pendientes = [c for c, a in self.archivos.items() if a["resultado"] is None or a["meta"] != meta]
        claves = pendientes or list(self.archivos)
        trabajos = [(clave, self._ruta_salida(clave)) for clave in claves]

        self.trabajando = True
        self.barra.configure(maximum=len(trabajos), value=0)
        for clave, _salida in trabajos:
            self._poner_fila(clave, "En espera…")
        self._actualizar_botones()
        threading.Thread(target=self._trabajar, args=(trabajos, meta), daemon=True).start()
        self.after(80, self._atender_cola)

    def _trabajar(self, trabajos, meta):
        """Hilo secundario: solo calcula y avisa por la cola; no toca la ventana."""
        procesador.META_DEFAULT = meta
        for clave, salida in trabajos:
            self.cola.put(("inicio", clave))
            try:
                resultado = procesador.procesar(clave)
                salida.parent.mkdir(parents=True, exist_ok=True)
                procesador.guardar(resultado, salida)
                self.cola.put(("listo", clave, resultado, salida, meta))
            except PermissionError:
                self.cola.put(("error", clave, f"No se pudo guardar «{salida.name}». "
                                               "¿Está abierto en Excel? Ciérralo y vuelve a calcular."))
            except Exception as e:  # noqa: BLE001 - se muestra al usuario
                self.cola.put(("error", clave, str(e)))
        self.cola.put(("fin",))

    def _atender_cola(self):
        try:
            while True:
                evento = self.cola.get_nowait()
                tipo, clave = evento[0], (evento[1] if len(evento) > 1 else None)
                if tipo == "fin":
                    self._terminar()
                    return
                if clave not in self.archivos:
                    continue
                if tipo == "inicio":
                    self._poner_fila(clave, "Procesando…")
                    self.tabla.see(clave)
                elif tipo == "listo":
                    _t, _c, resultado, salida, meta = evento
                    self.archivos[clave].update(resultado=resultado, salida=salida, error=None, meta=meta)
                    self._poner_fila(clave, "✅ Listo", resultado["_totales"], "listo")
                    self.barra.step(1)
                elif tipo == "error":
                    self.archivos[clave].update(resultado=None, salida=None, error=evento[2], meta=None)
                    self._poner_fila(clave, "❌ Error", etiqueta="error")
                    self.barra.step(1)
                if clave in self.tabla.selection():
                    self._mostrar_seleccion()
        except queue.Empty:
            pass
        self.after(80, self._atender_cola)

    def _terminar(self):
        self.trabajando = False
        self.barra.configure(value=0)
        self._actualizar_botones()
        self._actualizar_estado()
        self._mostrar_seleccion()
        if self.abrir_al_terminar.get():
            for archivo in self.archivos.values():
                if archivo["salida"]:
                    abrir_con_sistema(str(archivo["salida"]))
        errores = sum(1 for a in self.archivos.values() if a["error"])
        if errores:
            messagebox.showwarning("Terminado con errores",
                                   f"{errores} archivo(s) no se pudieron procesar. "
                                   "Selecciónalos en la lista para ver el motivo.")

    def _poner_fila(self, clave, estado, totales=None, etiqueta=""):
        valores = [Path(clave).name, estado]
        if totales:
            valores += [f"{totales['tickets']:,}", f"{totales['folios']:,}", f"{totales['km']:,.0f}",
                        f"{totales['faltante_lts']:,.2f}", f"${totales['faltante_pesos']:,.2f}",
                        totales["revisar"]]
        self.tabla.item(clave, values=valores, tags=(etiqueta,) if etiqueta else ())

    # ---------- vistas del archivo seleccionado ----------

    def _seleccionado(self):
        seleccion = self.tabla.selection()
        return self.archivos.get(seleccion[0]) if seleccion else None

    def _mostrar_seleccion(self):
        archivo = self._seleccionado()
        resultado = archivo["resultado"] if archivo else None
        if not archivo:
            self._escribir("Agrega un Excel para empezar." if not self.archivos else "Selecciona un archivo de la lista.")
        elif archivo["error"]:
            self._escribir(f"❌ No se pudo procesar este archivo:\n\n{archivo['error']}")
        elif not resultado:
            self._escribir("Pendiente. Presiona «Calcular todo».")
        else:
            self._escribir(self._texto_resumen(archivo))
        for nombre, arbol in self.vistas.items():
            self._llenar_tabla(arbol, resultado.get(nombre) if resultado else None)
        self._actualizar_botones()

    def _texto_resumen(self, archivo):
        t = archivo["resultado"]["_totales"]
        sin_meta = ", ".join(map(str, t["tractos_sin_meta"])) or "ninguno"
        return (
            f"✅ Listo\n\n"
            f"Tickets leídos:            {t['tickets']:,}\n"
            f"Líneas por día:            {t['lineas']:,}\n"
            f"Folios de liquidación:     {t['folios']:,}\n"
            f"Km recorridos (folios):    {t['km']:,.0f}\n\n"
            f"Faltante total:            {t['faltante_lts']:,.2f} L\n"
            f"Faltante total ($):        ${t['faltante_pesos']:,.2f}\n\n"
            f"Líneas para revisar:       {t['revisar']}\n"
            f"Cargas sin km (remolque):  {t['sin_km']}\n"
            f"Tractos con meta {archivo['meta']} por defecto ({len(t['tractos_sin_meta'])}):\n  {sin_meta}\n\n"
            f"Guardado en:\n  {archivo['salida']}\n"
        )

    def _llenar_tabla(self, arbol, tabla):
        arbol.delete(*arbol.get_children())
        if tabla is None or tabla.empty:
            arbol.configure(columns=())
            return
        columnas = list(tabla.columns)
        arbol.configure(columns=columnas)
        for col in columnas:
            numerica = pd.api.types.is_numeric_dtype(tabla[col]) and not pd.api.types.is_bool_dtype(tabla[col])
            arbol.heading(col, text=col, anchor="e" if numerica else "w")
            arbol.column(col, width=min(max(len(col) * 9 + 20, 80), 240), anchor="e" if numerica else "w", stretch=False)
        tiene_faltante = "FALTANTE LTS" in columnas
        for fila in tabla.head(MAX_FILAS_VISTA).itertuples(index=False):
            valores = dict(zip(columnas, fila))
            faltante = tiene_faltante and pd.notna(valores["FALTANTE LTS"]) and valores["FALTANTE LTS"] > 0
            arbol.insert("", "end", values=[formatear(v) for v in fila], tags=("faltante",) if faltante else ())
        if len(tabla) > MAX_FILAS_VISTA:
            arbol.insert("", "end", values=[f"… {len(tabla) - MAX_FILAS_VISTA:,} filas más en el Excel generado"])

    def _escribir(self, texto):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", texto)
        self.txt.configure(state="disabled")

    # ---------- acciones sobre el resultado ----------

    def abrir_resultado(self):
        archivo = self._seleccionado()
        if archivo and archivo["salida"]:
            abrir_con_sistema(str(archivo["salida"]))

    def mostrar_resultado_en_carpeta(self):
        archivo = self._seleccionado()
        if archivo and archivo["salida"]:
            mostrar_en_carpeta(archivo["salida"])

    # ---------- estado general ----------

    def _actualizar_botones(self):
        libre = "disabled" if self.trabajando else "normal"
        for boton in (self.btn_agregar, self.btn_carpeta):
            boton.configure(state=libre)
        hay_archivos = bool(self.archivos) and not self.trabajando
        for boton in (self.btn_quitar, self.btn_limpiar, self.btn_calcular):
            boton.configure(state="normal" if hay_archivos else "disabled")
        archivo = self._seleccionado()
        hay_salida = bool(archivo and archivo["salida"])
        for boton in (self.btn_ver, self.btn_ver_carpeta):
            boton.configure(state="normal" if hay_salida else "disabled")

    def _actualizar_estado(self):
        listos = [a["resultado"]["_totales"] for a in self.archivos.values() if a["resultado"]]
        if not self.archivos:
            texto = ""
        elif not listos:
            texto = f"{len(self.archivos)} archivo(s) en la lista"
        else:
            texto = (f"{len(listos)} de {len(self.archivos)} calculados  ·  "
                     f"Faltante total: {sum(t['faltante_lts'] for t in listos):,.2f} L  /  "
                     f"${sum(t['faltante_pesos'] for t in listos):,.2f}")
        self.lbl_estado.configure(text=texto)

    def _cerrar(self):
        try:
            self.config_usuario["meta"] = float(self.meta.get().replace(",", "."))
        except ValueError:
            pass
        self.config_usuario["carpeta_salida"] = self.carpeta_salida.get()
        self.config_usuario["abrir_al_terminar"] = self.abrir_al_terminar.get()
        guardar_config(self.config_usuario)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()

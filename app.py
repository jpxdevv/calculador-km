"""Interfaz de escritorio del Calculador de Kilometraje (Windows y Mac)."""
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import procesador

TITULO = "Calculador de Kilometraje y Rendimiento"


def abrir_con_sistema(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta], check=False)
    else:
        subprocess.run(["xdg-open", ruta], check=False)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO)
        self.geometry("620x520")
        self.minsize(560, 480)
        self.resultado = None
        self.ruta_entrada = None
        self.ruta_guardada = None

        marco = ttk.Frame(self, padding=20)
        marco.pack(fill="both", expand=True)

        ttk.Label(marco, text=TITULO, font=("", 16, "bold")).pack(anchor="w")
        ttk.Label(marco, text="Selecciona el Excel de cargas de diésel. El programa calcula km inicial, "
                              "km final, km recorridos, rendimiento, litros meta y faltantes.",
                  wraplength=560, foreground="#555").pack(anchor="w", pady=(4, 16))

        fila = ttk.Frame(marco)
        fila.pack(fill="x")
        self.btn_abrir = ttk.Button(fila, text="📂  Seleccionar Excel…", command=self.seleccionar)
        self.btn_abrir.pack(side="left")
        self.lbl_archivo = ttk.Label(fila, text="Ningún archivo seleccionado", foreground="#777")
        self.lbl_archivo.pack(side="left", padx=12)

        self.barra = ttk.Progressbar(marco, mode="indeterminate")
        self.barra.pack(fill="x", pady=16)

        caja = ttk.LabelFrame(marco, text="Resultado", padding=12)
        caja.pack(fill="both", expand=True)
        self.txt = tk.Text(caja, height=12, wrap="word", relief="flat",
                           font=("Menlo" if sys.platform == "darwin" else "Consolas", 11))
        self.txt.pack(fill="both", expand=True)
        self.txt.configure(state="disabled")

        botones = ttk.Frame(marco)
        botones.pack(fill="x", pady=(16, 0))
        self.btn_guardar = ttk.Button(botones, text="💾  Guardar resultado…", command=self.guardar,
                                      state="disabled")
        self.btn_guardar.pack(side="left")
        self.btn_ver = ttk.Button(botones, text="Abrir Excel generado", command=self.abrir_guardado,
                                  state="disabled")
        self.btn_ver.pack(side="left", padx=8)

    def escribir(self, texto):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", texto)
        self.txt.configure(state="disabled")

    def seleccionar(self):
        ruta = filedialog.askopenfilename(title="Selecciona el Excel",
                                          filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not ruta:
            return
        self.ruta_entrada = ruta
        self.ruta_guardada = None
        self.lbl_archivo.configure(text=Path(ruta).name, foreground="#000")
        self.btn_abrir.configure(state="disabled")
        self.btn_guardar.configure(state="disabled")
        self.btn_ver.configure(state="disabled")
        self.escribir("Procesando…")
        self.barra.start(12)
        threading.Thread(target=self._procesar, args=(ruta,), daemon=True).start()

    def _procesar(self, ruta):
        try:
            resultado = procesador.procesar(ruta)
            self.after(0, self._terminado, resultado, None)
        except Exception as e:  # noqa: BLE001 - se muestra al usuario
            self.after(0, self._terminado, None, e)

    def _terminado(self, resultado, error):
        self.barra.stop()
        self.btn_abrir.configure(state="normal")
        if error:
            self.resultado = None
            self.escribir(f"❌ Error:\n{error}")
            messagebox.showerror("Error", str(error))
            return
        self.resultado = resultado
        t = resultado["_totales"]
        sin_meta = ", ".join(map(str, t["tractos_sin_meta"])) or "ninguno"
        self.escribir(
            f"✅ Listo\n\n"
            f"Tickets leídos:            {t['tickets']:,}\n"
            f"Líneas por día:            {t['lineas']:,}\n"
            f"Folios de liquidación:     {t['folios']:,}\n"
            f"Km recorridos (folios):    {t['km']:,.0f}\n\n"
            f"Faltante total:            {t['faltante_lts']:,.2f} L\n"
            f"Faltante total ($):        ${t['faltante_pesos']:,.2f}\n\n"
            f"Líneas para revisar:       {t['revisar']}\n"
            f"Cargas sin km (remolque):  {t['sin_km']}\n"
            f"Tractos con meta {procesador.META_DEFAULT} por defecto ({len(t['tractos_sin_meta'])}):\n  {sin_meta}\n"
        )
        self.btn_guardar.configure(state="normal")
        self.guardar()

    def guardar(self):
        if not self.resultado:
            return
        sugerida = procesador.ruta_salida_para(self.ruta_entrada)
        ruta = filedialog.asksaveasfilename(title="Guardar resultado", defaultextension=".xlsx",
                                            initialdir=str(sugerida.parent), initialfile=sugerida.name,
                                            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            procesador.guardar(self.resultado, ruta)
        except PermissionError:
            messagebox.showerror("Error", "No se pudo guardar. ¿El archivo está abierto en Excel? Ciérralo e intenta de nuevo.")
            return
        self.ruta_guardada = ruta
        self.btn_ver.configure(state="normal")
        if messagebox.askyesno("Guardado", f"Se guardó:\n{ruta}\n\n¿Abrirlo ahora?"):
            abrir_con_sistema(ruta)

    def abrir_guardado(self):
        if self.ruta_guardada:
            abrir_con_sistema(self.ruta_guardada)


if __name__ == "__main__":
    App().mainloop()

# -*- coding: utf-8 -*-
"""
Reparador de hipervinculos - Libro de Gobierno
Interfaz grafica sencilla para Windows.
Requiere: openpyxl
"""

import os
import csv
import threading
from pathlib import Path
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None


class ReparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Reparador de hipervinculos")
        self.root.geometry("690x500")
        self.root.minsize(620, 440)
        self.root.configure(bg="#f4f6f8")

        self.archivo_var = tk.StringVar()
        self.estado_var = tk.StringVar(value="Selecciona el archivo Excel para comenzar.")
        self.progreso_var = tk.DoubleVar(value=0)

        self._crear_interfaz()

    def _crear_interfaz(self):
        header = tk.Frame(self.root, bg="#17365D", padx=22, pady=18)
        header.pack(fill="x")
        tk.Label(
            header, text="REPARADOR DE HIPERVÍNCULOS",
            font=("Segoe UI", 17, "bold"), fg="white", bg="#17365D"
        ).pack(anchor="w")
        tk.Label(
            header, text="Libro de Gobierno · Búsqueda automática en subcarpetas",
            font=("Segoe UI", 9), fg="#e6edf5", bg="#17365D"
        ).pack(anchor="w", pady=(4, 0))

        body = tk.Frame(self.root, bg="#f4f6f8", padx=22, pady=18)
        body.pack(fill="both", expand=True)

        tk.Label(
            body, text="1. Selecciona el archivo Excel (.xlsx)",
            font=("Segoe UI", 11, "bold"), bg="#f4f6f8", fg="#243447"
        ).pack(anchor="w")

        file_row = tk.Frame(body, bg="#f4f6f8")
        file_row.pack(fill="x", pady=(8, 18))
        entry = tk.Entry(
            file_row, textvariable=self.archivo_var, font=("Segoe UI", 9),
            state="readonly", readonlybackground="white"
        )
        entry.pack(side="left", fill="x", expand=True, ipady=7)
        tk.Button(
            file_row, text="Examinar...", command=self.seleccionar_archivo,
            font=("Segoe UI", 9, "bold"), bg="#dce6f1", fg="#17365D",
            relief="flat", padx=14, pady=7, cursor="hand2"
        ).pack(side="left", padx=(8, 0))

        info = tk.Frame(body, bg="white", padx=14, pady=12,
                        highlightbackground="#d7dee7", highlightthickness=1)
        info.pack(fill="x", pady=(0, 16))
        tk.Label(
            info, text="¿Qué hará el programa?",
            font=("Segoe UI", 10, "bold"), bg="white", fg="#243447"
        ).pack(anchor="w")
        bullets = [
            "• Recorre todas las hojas y celdas del libro.",
            "• Busca los PDF en la carpeta del Excel y todas sus subcarpetas.",
            "• Repara las rutas encontradas y reporta faltantes o nombres duplicados.",
            "• Guarda un Excel nuevo; nunca sobrescribe el archivo original."
        ]
        for item in bullets:
            tk.Label(
                info, text=item, font=("Segoe UI", 9), bg="white",
                fg="#455468", anchor="w"
            ).pack(anchor="w", pady=2)

        self.run_button = tk.Button(
            body, text="INICIAR REPARACIÓN", command=self.iniciar,
            font=("Segoe UI", 10, "bold"), bg="#217346", fg="white",
            activebackground="#185c37", activeforeground="white",
            relief="flat", padx=18, pady=11, cursor="hand2", state="disabled"
        )
        self.run_button.pack(anchor="w", pady=(0, 14))

        self.progress = ttk.Progressbar(
            body, variable=self.progreso_var, maximum=100, mode="determinate"
        )
        self.progress.pack(fill="x", pady=(0, 8))

        tk.Label(
            body, textvariable=self.estado_var, font=("Segoe UI", 9),
            bg="#f4f6f8", fg="#455468", wraplength=630, justify="left"
        ).pack(anchor="w")

        self.log = tk.Text(
            body, height=8, font=("Consolas", 8), bg="#ffffff",
            fg="#263238", relief="solid", bd=1, state="disabled"
        )
        self.log.pack(fill="both", expand=True, pady=(12, 0))

        footer = tk.Label(
            self.root,
            text="Importante: trabaja con una copia de respaldo. Revisa el resultado antes de usarlo como libro oficial.",
            font=("Segoe UI", 8), bg="#e9edf2", fg="#5b6572", padx=12, pady=8
        )
        footer.pack(fill="x")

    def log_line(self, text):
        self.root.after(0, self._append_log, text)

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def seleccionar_archivo(self):
        path = filedialog.askopenfilename(
            title="Selecciona el Libro de Gobierno",
            filetypes=[("Libro de Excel", "*.xlsx")]
        )
        if path:
            self.archivo_var.set(path)
            self.run_button.configure(state="normal")
            self.estado_var.set("Listo para revisar el libro seleccionado.")

    def iniciar(self):
        if not self.archivo_var.get():
            return
        self.run_button.configure(state="disabled")
        self.progreso_var.set(0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.estado_var.set("Procesando. No cierres el programa...")
        thread = threading.Thread(target=self.procesar, daemon=True)
        thread.start()

    def procesar(self):
        try:
            if load_workbook is None:
                raise RuntimeError("No se encontró openpyxl. Contacta a soporte.")

            excel = Path(self.archivo_var.get())
            raiz = excel.parent
            salida = excel.with_name(excel.stem + "_REPARADO.xlsx")
            reporte_path = excel.with_name(excel.stem + "_REPORTE.csv")

            if salida.exists():
                raise FileExistsError(
                    f"Ya existe el archivo de salida:\n{salida.name}\n"
                    "Renómbralo o muévelo antes de volver a ejecutar."
                )

            self.log_line(f"Carpeta principal: {raiz}")
            self.log_line("Buscando archivos PDF en todas las subcarpetas...")

            indice = defaultdict(list)
            total_pdf = 0
            for carpeta, subcarpetas, archivos in os.walk(raiz):
                subcarpetas[:] = [
                    s for s in subcarpetas
                    if not s.startswith(".") and s not in ("$RECYCLE.BIN", "System Volume Information")
                ]
                for nombre in archivos:
                    if nombre.lower().endswith(".pdf"):
                        ruta = (Path(carpeta) / nombre).resolve()
                        indice[nombre.casefold()].append(ruta)
                        total_pdf += 1

            self.log_line(f"PDF indexados: {total_pdf}")
            self.log_line("Abriendo libro de Excel...")
            wb = load_workbook(excel, keep_links=True)

            reparados = 0
            no_encontrados = 0
            duplicados = 0
            revisados = 0
            reporte = []

            hojas = wb.worksheets
            total_hojas = max(len(hojas), 1)

            for numero_hoja, ws in enumerate(hojas, start=1):
                self.log_line(f"Revisando hoja: {ws.title}")
                for fila in ws.iter_rows():
                    for celda in fila:
                        valor = celda.value
                        if not isinstance(valor, str):
                            continue

                        texto = valor.strip().strip('"').strip("'")
                        if not texto.lower().endswith(".pdf"):
                            continue

                        revisados += 1
                        normalizada = texto.replace("/", "\\")
                        partes = [p for p in normalizada.split("\\") if p and p != "."]

                        if partes and partes[0].casefold() == raiz.name.casefold():
                            partes = partes[1:]

                        candidato = raiz.joinpath(*partes) if partes else None

                        encontrado = None
                        if candidato is not None and candidato.is_file():
                            encontrado = candidato.resolve()
                        else:
                            nombre_pdf = Path(partes[-1]).name.casefold() if partes else ""
                            coincidencias = indice.get(nombre_pdf, [])

                            if len(coincidencias) == 1:
                                encontrado = coincidencias[0]
                            elif len(coincidencias) > 1:
                                duplicados += 1
                                reporte.append([
                                    ws.title, celda.coordinate, texto,
                                    "DUPLICADO - REQUIERE REVISION", ""
                                ])
                                continue
                            else:
                                no_encontrados += 1
                                reporte.append([
                                    ws.title, celda.coordinate, texto,
                                    "NO ENCONTRADO", ""
                                ])
                                continue

                        # Reemplaza/crea el destino del hipervínculo.
                        celda.hyperlink = encontrado.as_uri()
                        reparados += 1
                        reporte.append([
                            ws.title, celda.coordinate, texto,
                            "REPARADO", str(encontrado)
                        ])

                porcentaje = (numero_hoja / total_hojas) * 85
                self.root.after(0, self.progreso_var.set, porcentaje)

            self.log_line("Guardando una copia reparada...")
            wb.save(salida)
            wb.close()

            self.log_line("Generando reporte CSV...")
            with open(reporte_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Hoja", "Celda", "Ruta original",
                    "Resultado", "Documento localizado"
                ])
                writer.writerows(reporte)

            resumen = (
                f"Proceso terminado.\n\n"
                f"PDF encontrados en carpetas: {total_pdf}\n"
                f"Celdas con rutas PDF revisadas: {revisados}\n"
                f"Hipervínculos creados/reparados: {reparados}\n"
                f"Documentos no encontrados: {no_encontrados}\n"
                f"Coincidencias duplicadas: {duplicados}\n\n"
                f"Archivos generados:\n{salida.name}\n{reporte_path.name}\n\n"
                "Revisa el libro reparado y el reporte antes de utilizarlo."
            )

            self.root.after(0, self.progreso_var.set, 100)
            self.root.after(0, self.estado_var.set, "Finalizado. Revisa los resultados.")
            self.log_line(f"Terminado: {reparados} reparados, {no_encontrados} faltantes, {duplicados} duplicados.")
            self.root.after(0, lambda: messagebox.showinfo("Reparación terminada", resumen))

        except Exception as exc:
            self.log_line(f"ERROR: {exc}")
            self.root.after(0, self.estado_var.set, "Ocurrió un error. Revisa el mensaje.")
            self.root.after(0, lambda: messagebox.showerror("No se pudo completar", str(exc)))

        finally:
            self.root.after(0, lambda: self.run_button.configure(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ReparadorApp(root)
    root.mainloop()

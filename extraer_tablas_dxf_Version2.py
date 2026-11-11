from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ezdxf
from ezdxf.entities.acad_table import read_acad_table_content
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def nombre_hoja(nombre_base: str, usados: set[str]) -> str:
    """Genera un nombre de hoja válido y único para Excel."""
    nombre = re.sub(r"[\[\]:*?/\\]", "_", nombre_base).strip("'")
    nombre = (nombre or "Tabla")[:31]

    candidato = nombre
    contador = 2

    while candidato.lower() in usados:
        sufijo = f"_{contador}"
        candidato = f"{nombre[:31 - len(sufijo)]}{sufijo}"
        contador += 1

    usados.add(candidato.lower())
    return candidato


def ajustar_columnas(hoja) -> None:
    for columna in hoja.columns:
        letra = get_column_letter(columna[0].column)
        ancho = max(
            (len(str(celda.value)) for celda in columna if celda.value is not None),
            default=0,
        )
        hoja.column_dimensions[letra].width = min(max(ancho + 2, 10), 60)


def extraer_tablas(carpeta: Path, salida: Path, recursivo: bool) -> tuple[int, list[str]]:
    patron = "**/*.dxf" if recursivo else "*.dxf"
    archivos = sorted(carpeta.glob(patron))

    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos DXF en:\n{carpeta}")

    libro = Workbook()
    libro.remove(libro.active)
    usados: set[str] = set()
    total_tablas = 0
    errores = []

    for archivo in archivos:
        try:
            documento = ezdxf.readfile(archivo)
        except Exception as error:
            errores.append(f"{archivo.name}: no se pudo leer ({error})")
            continue

        numero_tabla = 0

        # Busca tablas tanto en modelspace como en las presentaciones.
        for layout in documento.layouts:
            for tabla in layout.query("ACAD_TABLE"):
                numero_tabla += 1

                try:
                    filas = read_acad_table_content(tabla)
                except Exception as error:
                    errores.append(
                        f"{archivo.name}, presentación {layout.name}: "
                        f"no se pudo leer una tabla ({error})"
                    )
                    continue

                nombre = f"{archivo.stem}_T{numero_tabla}"
                hoja = libro.create_sheet(nombre_hoja(nombre, usados))

                # Cada fila de la tabla DXF se escribe en una fila de Excel.
                for fila in filas:
                    valores = ["" if valor is None else str(valor) for valor in fila]
                    hoja.append(valores)

                    # Evita que Excel interprete el texto extraído como fórmula.
                    for celda in hoja[hoja.max_row]:
                        celda.data_type = "s"

                if hoja.max_row:
                    for celda in hoja[1]:
                        celda.font = Font(bold=True)
                    hoja.freeze_panes = "A2"
                    hoja.auto_filter.ref = hoja.dimensions
                    ajustar_columnas(hoja)

                total_tablas += 1

    if total_tablas == 0:
        hoja = libro.create_sheet("Resultado")
        hoja["A1"] = "No se encontraron tablas nativas ACAD_TABLE."

    salida.parent.mkdir(parents=True, exist_ok=True)
    libro.save(salida)

    return total_tablas, errores


class Aplicacion:
    def __init__(self, ventana: tk.Tk):
        self.ventana = ventana
        self.ventana.title("Extraer tablas de DXF a Excel")
        self.ventana.resizable(False, False)

        self.carpeta = tk.StringVar()
        self.recursivo = tk.BooleanVar(value=False)

        marco = ttk.Frame(ventana, padding=16)
        marco.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            marco,
            text="Selecciona la carpeta que contiene los archivos DXF:",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self.entrada = ttk.Entry(marco, textvariable=self.carpeta, width=55)
        self.entrada.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Button(
            marco,
            text="Examinar…",
            command=self.elegir_carpeta,
        ).grid(row=1, column=1)

        ttk.Checkbutton(
            marco,
            text="Buscar también en subcarpetas",
            variable=self.recursivo,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=10)

        self.boton = ttk.Button(
            marco,
            text="Extraer tablas a Excel",
            command=self.iniciar_extraccion,
        )
        self.boton.grid(row=3, column=0, columnspan=2, sticky="ew")

        self.estado = ttk.Label(marco, text="El Excel se guardará en la carpeta elegida.")
        self.estado.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def elegir_carpeta(self) -> None:
        seleccion = filedialog.askdirectory(title="Elige la carpeta con archivos DXF")
        if seleccion:
            self.carpeta.set(seleccion)
            self.estado.config(text="Carpeta seleccionada.")

    def iniciar_extraccion(self) -> None:
        ruta = self.carpeta.get().strip()

        if not ruta:
            messagebox.showwarning("Falta la carpeta", "Primero selecciona una carpeta.")
            return

        carpeta = Path(ruta)
        if not carpeta.is_dir():
            messagebox.showerror("Carpeta no válida", f"No existe esta carpeta:\n{carpeta}")
            return

        salida = carpeta / "tablas_extraidas.xlsx"
        self.boton.config(state="disabled")
        self.estado.config(text="Procesando archivos DXF…")
        self.ventana.update_idletasks()

        try:
            cantidad, errores = extraer_tablas(
                carpeta,
                salida,
                self.recursivo.get(),
            )
        except Exception as error:
            messagebox.showerror("No se pudo completar", str(error))
            self.estado.config(text="No se completó la extracción.")
            return
        finally:
            self.boton.config(state="normal")

        resumen = (
            f"Tablas extraídas: {cantidad}\n"
            f"Archivo Excel:\n{salida}"
        )

        if errores:
            resumen += f"\n\nAvisos ({len(errores)}):\n" + "\n".join(errores[:10])
            if len(errores) > 10:
                resumen += f"\n… y {len(errores) - 10} aviso(s) más."
            messagebox.showwarning("Extracción terminada con avisos", resumen)
        else:
            messagebox.showinfo("Extracción terminada", resumen)

        self.estado.config(text=f"Excel guardado en: {salida}")


if __name__ == "__main__":
    ventana = tk.Tk()
    Aplicacion(ventana)
    ventana.mainloop()
"""
Generador de base de datos de categorías MercadoLibre Argentina — Autopartes
Fuente de datos: capturas de pantalla del árbol de categorías de MeLi AR
Uso: python generar_excel.py
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_FILE = "categorias_autopartes_meli.xlsx"

HEADERS = [
    "Categoria_Nivel_1",
    "Categoria_Nivel_2",
    "Categoria_Nivel_3",
    "Categoria_Nivel_4",
    "Categoria_Nivel_5",
    "ID_Categoria_MLA",
    "Notas",
]

# ─────────────────────────────────────────────
# DATOS — agregar filas aquí a medida que
# se cargan desde las capturas de pantalla.
# Cada fila = (N1, N2, N3, N4, N5, ID_MLA, Notas)
# Dejar en "" los niveles que no apliquen.
# ─────────────────────────────────────────────
N1 = "Accesorios para Vehículos"
N2 = "Repuestos Autos y Camionetas"

CATEGORIAS = [
    # ── Nivel 3: categorías de Repuestos Autos y Camionetas ──────────────────
    (N1, N2, "Baterías",                    "", "", "", "Categoría terminal — sin subcategorías (va directo a productos)"),
    (N1, N2, "Carrocería",                  "", "", "", ""),
    (N1, N2, "Cerraduras y Llaves",         "", "", "", ""),
    (N1, N2, "Climatización",               "", "", "", ""),
    (N1, N2, "Componentes de Seguridad",    "", "", "", ""),
    (N1, N2, "Conducción Asistida Avanzada","", "", "", ""),
    (N1, N2, "Electroventiladores",         "", "", "", ""),
    (N1, N2, "Eléctricos, Híbridos y PHEV", "", "", "", ""),
    (N1, N2, "Encendido",                   "", "", "", ""),
    (N1, N2, "Escapes",                     "", "", "", ""),
    (N1, N2, "Filtros",                     "", "", "", ""),
    (N1, N2, "Frenos",                      "", "", "", ""),
    (N1, N2, "Iluminación",                 "", "", "", ""),
    (N1, N2, "Instalaciones Eléctricas",    "", "", "", ""),
    (N1, N2, "Inyección",                   "", "", "", ""),
    (N1, N2, "Motor",                       "", "", "", ""),
    (N1, N2, "Otros",                       "", "", "", ""),
    (N1, N2, "Repuestos de Exterior",       "", "", "", ""),
    (N1, N2, "Repuestos de Habitáculo",     "", "", "", ""),
    (N1, N2, "Suspensión y Dirección",      "", "", "", ""),
    (N1, N2, "Transmisión",                 "", "", "", ""),
    (N1, N2, "Ventanas y Sellos",           "", "", "", ""),
]


# ─────────────────────────────────────────────
# Estilos
# ─────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1A56DB")   # Azul MeLi
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)
HEADER_ALIGN  = Alignment(horizontal="center", vertical="center", wrap_text=True)

LEVEL_COLORS = {
    0: "DBEAFE",  # N1 — azul muy claro
    1: "EFF6FF",  # N2 — azul pastel
    2: "F8FAFF",  # N3 — casi blanco
    3: "FFFFFF",  # N4
    4: "FFFFFF",  # N5
    5: "FEF9C3",  # ID  — amarillo suave
    6: "F3F4F6",  # Notas — gris claro
}

THIN_BORDER = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)

COL_WIDTHS = [35, 35, 35, 30, 30, 18, 45]


def build_excel(categorias, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Categorias MeLi AR"

    # Cabecera
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 30

    # Datos
    for row_idx, row_data in enumerate(categorias, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = PatternFill("solid", fgColor=LEVEL_COLORS[col_idx - 1])
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

        ws.row_dimensions[row_idx].height = 20

    # Anchos de columna
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Filtro automático
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    # Panel fijo en fila de cabecera
    ws.freeze_panes = "A2"

    wb.save(output_file)
    print(f"[OK] Archivo generado: {output_file}  ({len(categorias)} fila/s de datos)")


if __name__ == "__main__":
    build_excel(CATEGORIAS, OUTPUT_FILE)

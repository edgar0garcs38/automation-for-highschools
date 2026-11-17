"""
Analizador de Cartera Vencida - CxC
Aplicación de escritorio para análisis de cuentas por cobrar.
Requiere: pip install pandas openpyxl matplotlib
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import os
from datetime import datetime
import unicodedata
from mpl_toolkits.mplot3d import Axes3D

# ─────────────────────────────────────────────
# COLORES Y ESTILO
# ─────────────────────────────────────────────
COLOR_BG        = "#F0F4F8"
COLOR_SIDEBAR   = "#1E3A5F"
COLOR_ACCENT    = "#2E86AB"
COLOR_ACCENT2   = "#A8DADC"
COLOR_BTN       = "#2E86AB"
COLOR_BTN_HOV   = "#1B6C8A"
COLOR_SUCCESS   = "#28A745"
COLOR_WHITE     = "#FFFFFF"
COLOR_TEXT      = "#1A1A2E"
COLOR_MUTED     = "#6C757D"
COLOR_CARD      = "#FFFFFF"
COLOR_BORDER    = "#DEE2E6"

MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
         7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
MESES_CORTOS = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
                7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

PALETTE = ["#2E86AB","#E84855","#3BB273","#F4A261","#9B5DE5",
           "#00B4D8","#F72585","#06D6A0","#FFB703","#A8DADC",
           "#E9C46A","#264653","#E76F51","#457B9D"]

LOGICA_PENDIENTES = "Cuentas por Cobrar Pendientes"
LOGICA_VENCIDAS = "Cuentas por Cobrar Vencidas"
LOGICAS_DISPONIBLES = (LOGICA_PENDIENTES, LOGICA_VENCIDAS)


# ─────────────────────────────────────────────
# PROCESAMIENTO DE DATOS
# ─────────────────────────────────────────────
def cargar_datos(ruta: str, logica: str = "AUTO") -> pd.DataFrame:
    df = pd.read_excel(
        ruta,
        dtype={
            "IDENTIFICACIÓN": str,
            "CEDULA": str,
            "CÉDULA": str,
            "TELÉFONO": str,
            "TELEFONO": str,
            "TELÉFONO_CONV": str,
            "TELEFONO_CONV": str,
        }
    )
    df.columns = df.columns.str.strip()
    df = normalizar_plantilla_rubros(df, logica)
    df["FECHA A PAGAR"] = pd.to_datetime(df["FECHA A PAGAR"], errors="coerce")
    df["MES"]  = df["FECHA A PAGAR"].dt.month
    df["AÑO"]  = df["FECHA A PAGAR"].dt.year
    df["CARRERA"] = df["CARRERA"].str.strip()
    df["APELLIDOS Y NOMBRES"] = df["APELLIDOS Y NOMBRES"].astype(str).str.strip()
    df["TELÉFONO"] = df["TELÉFONO"].fillna("").astype(str).str.strip()
    df["SESION"] = df["SESION"].fillna("").astype(str).str.strip()
    df["TIPO_SESION"] = df["SESION"].apply(clasificar_sesion)
    df["RUBRO_TIPO"] = df["RUBRO"].apply(clasificar_rubro)
    return df


def normalizar_plantilla_rubros(df: pd.DataFrame, logica: str = "AUTO") -> pd.DataFrame:
    """Convierte plantillas nuevas y antiguas al formato interno de la app."""
    df = df.copy()
    columnas = {_clave_columna(c): c for c in df.columns}
    aliases = _aliases_por_logica(logica)

    faltantes = []
    for destino, posibles in aliases.items():
        origen = next((columnas[p] for p in posibles if p in columnas), None)
        if origen:
            df[destino] = df[origen]
        elif destino in ("CORREO", "PERIODO", "VALOR PAGADO"):
            df[destino] = ""
        else:
            faltantes.append(f"{destino} ({', '.join(posibles)})")

    if faltantes:
        plantilla = "nueva" if logica == LOGICA_VENCIDAS else "anterior"
        raise ValueError(
            f"El archivo no coincide con la plantilla {plantilla} "
            f"({logica}). Columnas necesarias no encontradas: "
            + "; ".join(faltantes)
        )

    for col in ("VALOR A PAGAR", "VALOR TOTAL", "VALOR PAGADO", "VALOR PENDIENTE"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["IDENTIFICACIÓN"] = df["IDENTIFICACIÓN"].apply(_limpiar_identificador)
    df["TELÉFONO"] = df["TELÉFONO"].apply(_limpiar_identificador)

    return df


def _aliases_por_logica(logica: str):
    aliases_pendientes = {
        "IDENTIFICACIÓN": ("IDENTIFICACION",),
        "APELLIDOS Y NOMBRES": ("APELLIDOS Y NOMBRES",),
        "CARRERA": ("CARRERA",),
        "SESION": ("SESION",),
        "RUBRO": ("RUBRO",),
        "TELÉFONO": ("TELEFONO",),
        "CORREO": ("CORREO",),
        "FECHA A PAGAR": ("FECHA A PAGAR",),
        "VALOR A PAGAR": ("VALOR A PAGAR",),
        "VALOR TOTAL": ("VALOR TOTAL",),
        "VALOR PAGADO": ("VALOR PAGADO",),
        "VALOR PENDIENTE": ("VALOR PENDIENTE",),
        "PERIODO": ("PERIODO",),
    }

    aliases_vencidas = {
        "IDENTIFICACIÓN": ("CEDULA",),
        "APELLIDOS Y NOMBRES": ("NOMBRES Y APELLIDOS",),
        "CARRERA": ("CARRERA",),
        "SESION": ("SESION",),
        "RUBRO": ("RUBRO",),
        "TELÉFONO": ("TELEFONO",),
        "CORREO": ("EMAIL", "EMAIL INST", "EMAIL_INST"),
        "FECHA A PAGAR": ("FECHA DE VENCIMIENTO",),
        "VALOR A PAGAR": ("VALOR",),
        "VALOR TOTAL": ("TOTAL",),
        "VALOR PAGADO": ("ABONO",),
        "VALOR PENDIENTE": ("SALDO",),
        "PERIODO": ("NIVEL",),
    }

    if logica == LOGICA_PENDIENTES:
        return aliases_pendientes
    if logica == LOGICA_VENCIDAS:
        return aliases_vencidas

    combinado = {}
    for destino in aliases_pendientes:
        combinado[destino] = aliases_pendientes[destino] + aliases_vencidas[destino]
    return combinado


def _clave_columna(columna) -> str:
    texto = str(columna).strip().replace("_", " ")
    texto = " ".join(texto.split())
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.upper()


def _limpiar_identificador(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]
    return texto


def clasificar_sesion(sesion: str) -> str:
    s = str(sesion).upper()
    if "VALIDACIÓN" in s or "VALIDACION" in s:
        return "VALIDACIÓN"
    return "CARRERA"


def filtrar_modalidad(df: pd.DataFrame, modalidad: str) -> pd.DataFrame:
    if modalidad in ("VALIDACIÓN", "CARRERA"):
        return df[df["TIPO_SESION"] == modalidad].copy()
    return df.copy()


def clasificar_rubro(rubro: str) -> str:
    r = str(rubro).upper()
    if "MATRIC" in r:
        return "MATRÍCULA"
    if "ARANCEL" in r or "CUOTA" in r:
        return "ARANCEL"
    if "INSCR" in r:
        return "INSCRIPCIÓN"
    return "OTRO"


def obtener_años(df: pd.DataFrame):
    return sorted(df["AÑO"].dropna().unique().astype(int).tolist())


def calcular_pivots(df: pd.DataFrame, año: int, modalidad: str = "TODOS"):
    """Retorna (pivot_valores, pivot_estudiantes, todas_carreras)."""
    df = filtrar_modalidad(df, modalidad)
    todas = sorted(df["CARRERA"].dropna().unique().tolist())
    idx_full = pd.Index(todas, name="CARRERA")
    meses_idx = pd.Index(range(1, 13), name="MES")

    sub = df[df["AÑO"] == año].copy()

    pv = sub.pivot_table(index="CARRERA", columns="MES",
                         values="VALOR PENDIENTE", aggfunc="sum", fill_value=0)
    pv = pv.reindex(index=idx_full, columns=meses_idx, fill_value=0)

    pe = sub.pivot_table(index="CARRERA", columns="MES",
                         values="IDENTIFICACIÓN", aggfunc=pd.Series.nunique,
                         fill_value=0)
    pe = pe.reindex(index=idx_full, columns=meses_idx, fill_value=0)

    return pv, pe, todas


# ─────────────────────────────────────────────
# EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────
def exportar_excel(pv: pd.DataFrame, pe: pd.DataFrame,
                   año: int, ruta_salida: str):

    wb = Workbook()
    wb.remove(wb.active)

    _hoja_resumen(wb, pv, pe, año)
    _hoja_valores(wb, pv, año)
    _hoja_estudiantes(wb, pe, año)

    wb.save(ruta_salida)
    return ruta_salida


def exportar_validacion_carrera(df: pd.DataFrame, año: int, modalidad: str,
                                ruta_salida: str):
    """Genera un reporte tipo Validacion - Carrera con teléfono por estudiante."""
    sub = filtrar_modalidad(df[df["AÑO"] == año].copy(), modalidad)
    sub = sub[sub["VALOR PENDIENTE"].fillna(0) != 0].copy()
    if sub.empty:
        raise ValueError("No hay datos para el año y modalidad seleccionados.")

    wb = Workbook()
    wb.remove(wb.active)

    resumen_items = []
    for (carrera, tipo), grupo in sub.groupby(["CARRERA", "TIPO_SESION"], sort=True):
        ws_name = _nombre_hoja_reporte(wb, carrera, tipo)
        ws = wb.create_sheet(ws_name)
        total_row = _hoja_detalle_validacion_carrera(ws, grupo, carrera, tipo, año)
        resumen_items.append((carrera, tipo, ws_name, total_row))

    _hoja_resumen_validacion_carrera(wb, resumen_items, año, modalidad)
    wb.move_sheet(wb["Resumen"], offset=-(len(wb.worksheets) - 1))
    wb.save(ruta_salida)
    return ruta_salida


# ── estilos helpers ──────────────────────────
def _borde_fino():
    thin = Side(style="thin", color="B0B8C1")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color.lstrip("#"))

def _centrado(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _izq():
    return Alignment(horizontal="left", vertical="center")


def _fijar_alturas_filas(ws, fila_inicio=1, fila_fin=None, alto=18):
    fila_fin = fila_fin or ws.max_row
    for row in range(fila_inicio, fila_fin + 1):
        ws.row_dimensions[row].height = alto


def _limpiar_nombre_hoja(texto):
    invalidos = '[]:*?/\\'
    limpio = "".join("_" if c in invalidos else c for c in str(texto))
    return limpio.strip() or "Hoja"


def _nombre_hoja_reporte(wb, carrera, tipo):
    base = _limpiar_nombre_hoja(str(carrera).title())
    sufijo = " (Valid)" if tipo == "VALIDACIÓN" else " (Carr)"
    nombre = f"{base[:31-len(sufijo)]}{sufijo}"
    if nombre not in wb.sheetnames:
        return nombre

    for i in range(2, 100):
        extra = f" {i}"
        candidato = f"{nombre[:31-len(extra)]}{extra}"
        if candidato not in wb.sheetnames:
            return candidato
    raise ValueError("No se pudo crear un nombre único de hoja.")


def _nombre_para_formula(nombre_hoja):
    return "'" + nombre_hoja.replace("'", "''") + "'"


def _meses_presentes(grupo):
    meses = sorted(int(m) for m in grupo["MES"].dropna().unique())
    return [m for m in meses if 1 <= m <= 12] or list(range(1, 13))


def _pivot_estudiantes_con_telefono(grupo, meses):
    base = (
        grupo.pivot_table(
            index=["APELLIDOS Y NOMBRES", "TELÉFONO"],
            columns="MES",
            values="VALOR PENDIENTE",
            aggfunc="sum",
            fill_value=0
        )
        .reindex(columns=meses, fill_value=0)
        .reset_index()
        .sort_values("APELLIDOS Y NOMBRES")
    )
    return base


def _hoja_detalle_validacion_carrera(ws, grupo, carrera, tipo, año):
    meses = _meses_presentes(grupo)
    data = _pivot_estudiantes_con_telefono(grupo, meses)
    last_col = 3 + len(meses)
    last_col_letter = get_column_letter(last_col)

    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last_col)
    titulo = ws["A5"]
    titulo.value = f"CUENTAS POR COBRAR {año} ({tipo})"
    titulo.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    titulo.fill = _fill("1E3A5F")
    titulo.alignment = _centrado()
    ws.row_dimensions[5].height = 24

    ws["A7"].value = f"CARRERA: {carrera}"
    ws["A7"].font = Font(name="Arial", bold=True, size=10)

    headers = ["Estudiante", "Teléfono"] + [MESES_CORTOS[m].lower() for m in meses] + ["Total"]
    header_row = 9
    for col, header in enumerate(headers, 1):
        cell = ws.cell(header_row, col)
        cell.value = header
        cell.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill = _fill("2E86AB")
        cell.alignment = _centrado(wrap=True)
        cell.border = _borde_fino()

    start_row = header_row + 1
    for idx, (_, row_data) in enumerate(data.iterrows()):
        row = start_row + idx
        alt = _fill("F8FAFC") if idx % 2 == 0 else _fill("FFFFFF")
        ws.cell(row, 1).value = row_data["APELLIDOS Y NOMBRES"]
        ws.cell(row, 2).value = row_data["TELÉFONO"]
        for col in (1, 2):
            ws.cell(row, col).font = Font(name="Arial", size=9)
            ws.cell(row, col).fill = alt
            ws.cell(row, col).alignment = _izq() if col == 1 else _centrado()
            ws.cell(row, col).border = _borde_fino()

        for pos, mes in enumerate(meses, 3):
            val = float(row_data.get(mes, 0) or 0)
            cell = ws.cell(row, pos)
            cell.value = val if val else None
            cell.number_format = '#,##0.00'
            cell.font = Font(name="Arial", size=9)
            cell.fill = alt
            cell.alignment = _centrado()
            cell.border = _borde_fino()

        total_cell = ws.cell(row, last_col)
        total_cell.value = f"=SUM(C{row}:{get_column_letter(last_col-1)}{row})"
        total_cell.number_format = '#,##0.00'
        total_cell.font = Font(name="Arial", size=9, bold=True)
        total_cell.fill = _fill("DCF0F8")
        total_cell.alignment = _centrado()
        total_cell.border = _borde_fino()

    total_row = start_row + len(data)
    ws.cell(total_row, 1).value = "Total general"
    ws.cell(total_row, 1).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(total_row, 1).fill = _fill("1E3A5F")
    ws.cell(total_row, 1).alignment = _izq()
    ws.cell(total_row, 1).border = _borde_fino()
    ws.cell(total_row, 2).fill = _fill("1E3A5F")
    ws.cell(total_row, 2).border = _borde_fino()
    for col in range(3, last_col + 1):
        col_letter = get_column_letter(col)
        cell = ws.cell(total_row, col)
        cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{total_row-1})"
        cell.number_format = '#,##0.00'
        cell.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill = _fill("1E3A5F")
        cell.alignment = _centrado()
        cell.border = _borde_fino()

    if len(data) > 0:
        ref = f"A{header_row}:{last_col_letter}{total_row}"
        table_name = f"Tbl_{ws.parent.worksheets.index(ws) + 1:02d}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=True,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(table)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 15
    for col in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 11
    _fijar_alturas_filas(ws, 10, max(total_row - 1, 10), 18)
    ws.row_dimensions[9].height = 22
    ws.row_dimensions[total_row].height = 22
    ws.freeze_panes = "C10"
    ws.sheet_view.showGridLines = False
    return total_row


def _hoja_resumen_validacion_carrera(wb, resumen_items, año, modalidad):
    ws = wb.create_sheet("Resumen")
    ws.merge_cells("A5:N5")
    titulo = ws["A5"]
    titulo.value = f"RESUMEN CUENTAS POR COBRAR {año} - {modalidad}"
    titulo.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    titulo.fill = _fill("1E3A5F")
    titulo.alignment = _centrado()
    ws.row_dimensions[5].height = 24

    ws.merge_cells("A8:A9")
    ws.cell(8, 1).value = "CARRERA"
    for start_col, label in ((2, f"Trim.I - {año}"), (5, f"Trim.II - {año}"),
                             (8, f"Trim.III - {año}"), (11, f"Trim.IV - {año}")):
        ws.merge_cells(start_row=8, start_column=start_col, end_row=8, end_column=start_col + 2)
        ws.cell(8, start_col).value = label
    ws.merge_cells("N8:N9")
    ws.cell(8, 14).value = f"Total {año}"

    for col in range(1, 15):
        cell = ws.cell(8, col)
        cell.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill = _fill("2E86AB" if col < 14 else "1B6C8A")
        cell.alignment = _centrado(wrap=True)
        cell.border = _borde_fino()
    for col, mes in enumerate(range(1, 13), 2):
        cell = ws.cell(9, col)
        cell.value = MESES_CORTOS[mes].lower()
        cell.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill = _fill("2E86AB")
        cell.alignment = _centrado()
        cell.border = _borde_fino()

    resumen_items = sorted(resumen_items, key=lambda x: (x[0], x[1]))
    for idx, (carrera, tipo, hoja, total_row) in enumerate(resumen_items, 10):
        etiqueta = f"{carrera} ({tipo})" if modalidad == "TODOS" else carrera
        alt = _fill("F8FAFC") if idx % 2 == 0 else _fill("FFFFFF")
        ws.cell(idx, 1).value = etiqueta
        ws.cell(idx, 1).font = Font(name="Arial", size=9)
        ws.cell(idx, 1).alignment = _izq()
        ws.cell(idx, 1).border = _borde_fino()
        ws.cell(idx, 1).fill = alt

        sheet_ref = _nombre_para_formula(hoja)
        for mes in range(1, 13):
            col = mes + 1
            source_col = _columna_mes_en_hoja(wb[hoja], mes)
            cell = ws.cell(idx, col)
            if source_col:
                cell.value = f"={sheet_ref}!{get_column_letter(source_col)}{total_row}"
            cell.number_format = '#,##0.00'
            cell.font = Font(name="Arial", size=9)
            cell.alignment = _centrado()
            cell.border = _borde_fino()
            cell.fill = alt

        total = ws.cell(idx, 14)
        total.value = f"=SUM(B{idx}:M{idx})"
        total.number_format = '#,##0.00'
        total.font = Font(name="Arial", bold=True, size=9)
        total.fill = _fill("DCF0F8")
        total.alignment = _centrado()
        total.border = _borde_fino()

    total_row = 10 + len(resumen_items)
    ws.cell(total_row, 1).value = "TOTAL GENERAL"
    ws.cell(total_row, 1).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(total_row, 1).fill = _fill("1E3A5F")
    ws.cell(total_row, 1).alignment = _izq()
    ws.cell(total_row, 1).border = _borde_fino()
    for col in range(2, 15):
        col_letter = get_column_letter(col)
        cell = ws.cell(total_row, col)
        cell.value = f"=SUM({col_letter}10:{col_letter}{total_row-1})"
        cell.number_format = '#,##0.00'
        cell.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill = _fill("1E3A5F")
        cell.alignment = _centrado()
        cell.border = _borde_fino()

    ws.column_dimensions["A"].width = 58
    for col in range(2, 14):
        ws.column_dimensions[get_column_letter(col)].width = 11
    ws.column_dimensions["N"].width = 13
    _fijar_alturas_filas(ws, 10, max(total_row - 1, 10), 18)
    ws.row_dimensions[8].height = 22
    ws.row_dimensions[9].height = 20
    ws.row_dimensions[total_row].height = 22
    ws.freeze_panes = "B10"
    ws.sheet_view.showGridLines = False


def _columna_mes_en_hoja(ws, mes):
    buscado = MESES_CORTOS[mes].lower()
    for col in range(3, ws.max_column):
        if str(ws.cell(9, col).value).strip().lower() == buscado:
            return col
    return None

def _aplicar_tabla(ws, pv, pe, año, incluir_est=True):
    """Escribe encabezados + datos + totales en la hoja de resumen."""

    carreras = pv.index.tolist()
    n_car    = len(carreras)

    # ── Fila 1: título principal ──────────────
    ws.merge_cells("A1:AC1")
    c = ws["A1"]
    c.value     = f"ANÁLISIS MENSUAL CARTERA VENCIDA {año}"
    c.font      = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill      = _fill("1E3A5F")
    c.alignment = _centrado()
    ws.row_dimensions[1].height = 28

    # ── Fila 2: sub‑encabezados de meses ──────
    col_base = 3  # C
    ws["A2"].value = "#"
    ws["B2"].value = "Carrera"
    for m in range(1, 13):
        col_v = col_base + (m - 1) * (2 if incluir_est else 1)
        ws.cell(2, col_v).value = MESES[m]
        ws.cell(2, col_v).font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        ws.cell(2, col_v).fill  = _fill("2E86AB")
        ws.cell(2, col_v).alignment = _centrado(wrap=True)
        if incluir_est:
            ws.cell(2, col_v).value = f"{MESES[m]}\n($ / Est.)"
            ws.merge_cells(start_row=2, start_column=col_v,
                           end_row=2, end_column=col_v + 1)

    # columnas totales
    col_tot_v  = col_base + 12 * (2 if incluir_est else 1)
    col_tot_e  = col_tot_v + 1 if incluir_est else None

    ws.cell(2, col_tot_v).value = "TOTAL $"
    ws.cell(2, col_tot_v).font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
    ws.cell(2, col_tot_v).fill  = _fill("1B6C8A")
    ws.cell(2, col_tot_v).alignment = _centrado(wrap=True)
    if incluir_est:
        ws.cell(2, col_tot_e).value = "TOTAL Est."
        ws.cell(2, col_tot_e).font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        ws.cell(2, col_tot_e).fill  = _fill("1B6C8A")
        ws.cell(2, col_tot_e).alignment = _centrado(wrap=True)

    ws.row_dimensions[2].height = 32

    # ── Filas de datos ─────────────────────────
    for i, car in enumerate(carreras):
        row = 3 + i
        alt = _fill("F0F4F8") if i % 2 == 0 else _fill("FFFFFF")

        ws.cell(row, 1).value = i + 1
        ws.cell(row, 1).font  = Font(name="Arial", size=9, color="6C757D")
        ws.cell(row, 1).alignment = _centrado()

        ws.cell(row, 2).value = car
        ws.cell(row, 2).font  = Font(name="Arial", size=9, bold=False)
        ws.cell(row, 2).alignment = _izq()

        for m in range(1, 13):
            col_v = col_base + (m - 1) * (2 if incluir_est else 1)
            val_v = pv.at[car, m]
            ws.cell(row, col_v).value       = val_v if val_v != 0 else "-"
            ws.cell(row, col_v).number_format = '#,##0.00'
            ws.cell(row, col_v).font        = Font(name="Arial", size=9)
            ws.cell(row, col_v).fill        = alt
            ws.cell(row, col_v).alignment   = _centrado()
            ws.cell(row, col_v).border      = _borde_fino()

            if incluir_est:
                val_e = pe.at[car, m]
                ws.cell(row, col_v + 1).value       = int(val_e) if val_e != 0 else "-"
                ws.cell(row, col_v + 1).font        = Font(name="Arial", size=9,
                                                           color="2E86AB", bold=True)
                ws.cell(row, col_v + 1).fill        = alt
                ws.cell(row, col_v + 1).alignment   = _centrado()
                ws.cell(row, col_v + 1).border      = _borde_fino()

        # totales fila – fórmulas Excel
        col_letras_v = [get_column_letter(col_base + (m-1)*(2 if incluir_est else 1))
                        for m in range(1, 13)]
        formula_v = "=SUMIF({cols},{row})".format(
            cols=",".join(f"{c}{row}" for c in col_letras_v), row=row)
        # usamos SUM directo
        ws.cell(row, col_tot_v).value = \
            f"=SUM({col_letras_v[0]}{row}:{col_letras_v[-1]}{row})"
        ws.cell(row, col_tot_v).number_format = '#,##0.00'
        ws.cell(row, col_tot_v).font   = Font(name="Arial", size=9, bold=True)
        ws.cell(row, col_tot_v).fill   = _fill("DCF0F8")
        ws.cell(row, col_tot_v).alignment = _centrado()
        ws.cell(row, col_tot_v).border = _borde_fino()

        if incluir_est:
            col_letras_e = [get_column_letter(col_base + (m-1)*2 + 1)
                            for m in range(1, 13)]
            ws.cell(row, col_tot_e).value = \
                f"=SUM({col_letras_e[0]}{row}:{col_letras_e[-1]}{row})"
            ws.cell(row, col_tot_e).font   = Font(name="Arial", size=9, bold=True,
                                                  color="2E86AB")
            ws.cell(row, col_tot_e).fill   = _fill("DCF0F8")
            ws.cell(row, col_tot_e).alignment = _centrado()
            ws.cell(row, col_tot_e).border = _borde_fino()

    # ── Fila TOTAL ─────────────────────────────
    row_tot = 3 + n_car
    ws.cell(row_tot, 1).value = ""
    ws.cell(row_tot, 2).value = f"TOTAL {año}"
    ws.cell(row_tot, 2).font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(row_tot, 2).fill  = _fill("1E3A5F")
    ws.cell(row_tot, 2).alignment = _izq()

    for m in range(1, 13):
        col_v = col_base + (m - 1) * (2 if incluir_est else 1)
        col_l = get_column_letter(col_v)
        ws.cell(row_tot, col_v).value = f"=SUM({col_l}3:{col_l}{row_tot-1})"
        ws.cell(row_tot, col_v).number_format = '#,##0.00'
        ws.cell(row_tot, col_v).font  = Font(name="Arial", bold=True, size=9,
                                             color="FFFFFF")
        ws.cell(row_tot, col_v).fill  = _fill("1E3A5F")
        ws.cell(row_tot, col_v).alignment = _centrado()
        ws.cell(row_tot, col_v).border = _borde_fino()

        if incluir_est:
            col_l2 = get_column_letter(col_v + 1)
            ws.cell(row_tot, col_v+1).value = f"=SUM({col_l2}3:{col_l2}{row_tot-1})"
            ws.cell(row_tot, col_v+1).font  = Font(name="Arial", bold=True, size=9,
                                                   color="FFFFFF")
            ws.cell(row_tot, col_v+1).fill  = _fill("1E3A5F")
            ws.cell(row_tot, col_v+1).alignment = _centrado()
            ws.cell(row_tot, col_v+1).border = _borde_fino()

    col_l_tot = get_column_letter(col_tot_v)
    ws.cell(row_tot, col_tot_v).value = f"=SUM({col_l_tot}3:{col_l_tot}{row_tot-1})"
    ws.cell(row_tot, col_tot_v).number_format = '#,##0.00'
    ws.cell(row_tot, col_tot_v).font  = Font(name="Arial", bold=True, size=10,
                                             color="FFFFFF")
    ws.cell(row_tot, col_tot_v).fill  = _fill("1E3A5F")
    ws.cell(row_tot, col_tot_v).alignment = _centrado()

    if incluir_est:
        col_l_tot_e = get_column_letter(col_tot_e)
        ws.cell(row_tot, col_tot_e).value = \
            f"=SUM({col_l_tot_e}3:{col_l_tot_e}{row_tot-1})"
        ws.cell(row_tot, col_tot_e).font  = Font(name="Arial", bold=True, size=10,
                                                 color="FFFFFF")
        ws.cell(row_tot, col_tot_e).fill  = _fill("1E3A5F")
        ws.cell(row_tot, col_tot_e).alignment = _centrado()

    ws.row_dimensions[row_tot].height = 22

    # ── Anchos de columna ──────────────────────
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 38
    for m in range(1, 13):
        col_v = col_base + (m - 1) * (2 if incluir_est else 1)
        ws.column_dimensions[get_column_letter(col_v)].width = 11
        if incluir_est:
            ws.column_dimensions[get_column_letter(col_v + 1)].width = 7
    ws.column_dimensions[get_column_letter(col_tot_v)].width = 13
    if incluir_est:
        ws.column_dimensions[get_column_letter(col_tot_e)].width = 9

    _fijar_alturas_filas(ws, 3, max(row_tot - 1, 3), 18)
    ws.row_dimensions[row_tot].height = 22
    ws.freeze_panes = "C3"


def _hoja_resumen(wb, pv, pe, año):
    ws = wb.create_sheet(f"Resumen {año}")
    _aplicar_tabla(ws, pv, pe, año, incluir_est=True)


def _hoja_valores(wb, pv, año):
    ws = wb.create_sheet(f"Valores {año}")

    carreras = pv.index.tolist()
    n = len(carreras)

    ws.merge_cells("A1:N1")
    c = ws["A1"]
    c.value     = f"VALOR PENDIENTE MENSUAL POR CARRERA – {año}"
    c.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill      = _fill("1E3A5F")
    c.alignment = _centrado()
    ws.row_dimensions[1].height = 26

    headers = ["#", "Carrera"] + [MESES[m] for m in range(1, 13)] + ["TOTAL"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(2, col)
        cell.value = h
        cell.font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill  = _fill("2E86AB")
        cell.alignment = _centrado(wrap=True)
        cell.border = _borde_fino()
    ws.row_dimensions[2].height = 22

    for i, car in enumerate(carreras):
        row = 3 + i
        alt = _fill("F0F4F8") if i % 2 == 0 else _fill("FFFFFF")
        ws.cell(row, 1).value = i + 1
        ws.cell(row, 1).alignment = _centrado()
        ws.cell(row, 2).value = car
        ws.cell(row, 2).font  = Font(name="Arial", size=9)
        ws.cell(row, 2).fill  = alt
        ws.cell(row, 2).alignment = _izq()
        ws.cell(row, 2).border = _borde_fino()
        for m in range(1, 13):
            val = pv.at[car, m]
            ws.cell(row, 2 + m).value = val if val != 0 else "-"
            ws.cell(row, 2 + m).number_format = '#,##0.00'
            ws.cell(row, 2 + m).font   = Font(name="Arial", size=9)
            ws.cell(row, 2 + m).fill   = alt
            ws.cell(row, 2 + m).alignment = _centrado()
            ws.cell(row, 2 + m).border = _borde_fino()
        ws.cell(row, 15).value = f"=SUM(C{row}:N{row})"
        ws.cell(row, 15).number_format = '#,##0.00'
        ws.cell(row, 15).font  = Font(name="Arial", size=9, bold=True)
        ws.cell(row, 15).fill  = _fill("DCF0F8")
        ws.cell(row, 15).alignment = _centrado()
        ws.cell(row, 15).border = _borde_fino()

    row_t = 3 + n
    ws.cell(row_t, 2).value = f"TOTAL {año}"
    ws.cell(row_t, 2).font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(row_t, 2).fill  = _fill("1E3A5F")
    for m in range(1, 13):
        col_l = get_column_letter(2 + m)
        ws.cell(row_t, 2 + m).value = f"=SUM({col_l}3:{col_l}{row_t-1})"
        ws.cell(row_t, 2 + m).number_format = '#,##0.00'
        ws.cell(row_t, 2 + m).font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        ws.cell(row_t, 2 + m).fill  = _fill("1E3A5F")
        ws.cell(row_t, 2 + m).alignment = _centrado()
        ws.cell(row_t, 2 + m).border = _borde_fino()
    ws.cell(row_t, 15).value = f"=SUM(O3:O{row_t-1})"
    ws.cell(row_t, 15).number_format = '#,##0.00'
    ws.cell(row_t, 15).font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(row_t, 15).fill  = _fill("1E3A5F")
    ws.cell(row_t, 15).alignment = _centrado()
    ws.row_dimensions[row_t].height = 22

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 38
    for col in range(3, 16):
        ws.column_dimensions[get_column_letter(col)].width = 12
    _fijar_alturas_filas(ws, 3, max(row_t - 1, 3), 18)
    ws.row_dimensions[row_t].height = 22
    ws.freeze_panes = "C3"


def _hoja_estudiantes(wb, pe, año):
    ws = wb.create_sheet(f"Estudiantes {año}")
    carreras = pe.index.tolist()
    n = len(carreras)

    ws.merge_cells("A1:N1")
    c = ws["A1"]
    c.value     = f"ESTUDIANTES CON DEUDA MENSUAL POR CARRERA – {año}"
    c.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill      = _fill("1E3A5F")
    c.alignment = _centrado()
    ws.row_dimensions[1].height = 26

    headers = ["#", "Carrera"] + [MESES[m] for m in range(1, 13)] + ["TOTAL"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(2, col)
        cell.value = h
        cell.font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        cell.fill  = _fill("3BB273")
        cell.alignment = _centrado(wrap=True)
        cell.border = _borde_fino()
    ws.row_dimensions[2].height = 22

    for i, car in enumerate(carreras):
        row = 3 + i
        alt = _fill("F0F4F8") if i % 2 == 0 else _fill("FFFFFF")
        ws.cell(row, 1).value = i + 1
        ws.cell(row, 1).alignment = _centrado()
        ws.cell(row, 2).value = car
        ws.cell(row, 2).font  = Font(name="Arial", size=9)
        ws.cell(row, 2).fill  = alt
        ws.cell(row, 2).alignment = _izq()
        ws.cell(row, 2).border = _borde_fino()
        for m in range(1, 13):
            val = int(pe.at[car, m])
            ws.cell(row, 2 + m).value = val if val != 0 else "-"
            ws.cell(row, 2 + m).font  = Font(name="Arial", size=9)
            ws.cell(row, 2 + m).fill  = alt
            ws.cell(row, 2 + m).alignment = _centrado()
            ws.cell(row, 2 + m).border = _borde_fino()
        ws.cell(row, 15).value = f"=SUMIF(C{row}:N{row},\">0\")"
        ws.cell(row, 15).font  = Font(name="Arial", size=9, bold=True)
        ws.cell(row, 15).fill  = _fill("D4EDDA")
        ws.cell(row, 15).alignment = _centrado()
        ws.cell(row, 15).border = _borde_fino()

    row_t = 3 + n
    ws.cell(row_t, 2).value = f"TOTAL {año}"
    ws.cell(row_t, 2).font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(row_t, 2).fill  = _fill("1E3A5F")
    for m in range(1, 13):
        col_l = get_column_letter(2 + m)
        ws.cell(row_t, 2 + m).value = f"=SUM({col_l}3:{col_l}{row_t-1})"
        ws.cell(row_t, 2 + m).font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        ws.cell(row_t, 2 + m).fill  = _fill("1E3A5F")
        ws.cell(row_t, 2 + m).alignment = _centrado()
        ws.cell(row_t, 2 + m).border = _borde_fino()
    ws.cell(row_t, 15).value = f"=SUM(O3:O{row_t-1})"
    ws.cell(row_t, 15).font  = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    ws.cell(row_t, 15).fill  = _fill("1E3A5F")
    ws.cell(row_t, 15).alignment = _centrado()
    ws.row_dimensions[row_t].height = 22

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 38
    for col in range(3, 16):
        ws.column_dimensions[get_column_letter(col)].width = 10
    _fijar_alturas_filas(ws, 3, max(row_t - 1, 3), 18)
    ws.row_dimensions[row_t].height = 22
    ws.freeze_panes = "C3"


# ─────────────────────────────────────────────
# GRÁFICOS
# ─────────────────────────────────────────────
def _filtrar(df, año, carreras_sel, meses_sel, rubros_sel):
    mask = (df["AÑO"] == año)
    if carreras_sel:
        mask &= df["CARRERA"].isin(carreras_sel)
    if meses_sel:
        mask &= df["MES"].isin(meses_sel)
    if rubros_sel:
        mask &= df["RUBRO_TIPO"].isin(rubros_sel)
    return df[mask]

def _preparar_pivot(sub, carreras_sel, todas_carreras, tipo_valor):
    """
    Retorna pivot según métrica seleccionada.
    """

    if tipo_valor == "VALOR":
        pivot = sub.pivot_table(
            index="CARRERA",
            columns="MES",
            values="VALOR PENDIENTE",
            aggfunc="sum",
            fill_value=0
        )
    else:
        pivot = sub.pivot_table(
            index="CARRERA",
            columns="MES",
            values="IDENTIFICACIÓN",
            aggfunc=pd.Series.nunique,
            fill_value=0
        )

    pivot = pivot.reindex(columns=range(1, 13), fill_value=0)
    pivot = pivot[pivot.sum(axis=1) > 0]

    unificado = not carreras_sel

    if unificado and not pivot.empty:
        total_row = pivot.sum(axis=0).to_frame(name="TOTAL GENERAL").T
        return total_row, True

    return pivot, False

def _fmt_val(v, tipo_valor):
    if tipo_valor == "VALOR":
        if v >= 1_000:
            return f"${v/1_000:.1f}k"
        return f"${v:.0f}"

    return f"{int(v)}"

def grafico_barras(fig, df, año, carreras_sel, meses_sel, rubros_sel, tipo_valor):
    fig.clear()
    ax = fig.add_subplot(111)
    sub = _filtrar(df, año, carreras_sel, meses_sel, rubros_sel)

    if sub.empty:
        ax.text(0.5, 0.5, "Sin datos para los filtros seleccionados",
                ha="center", va="center", fontsize=12, color=COLOR_MUTED)
        fig.canvas.draw()
        return

    todas = sorted(df["CARRERA"].unique())
    pivot, unificado = _preparar_pivot(
        sub,
        carreras_sel,
        todas,
        tipo_valor
    )

    if pivot.empty:
        ax.text(0.5, 0.5, "Todas las carreras tienen valor $0",
                ha="center", va="center", fontsize=12, color=COLOR_MUTED)
        fig.canvas.draw()
        return

    meses_presentes = [m for m in range(1, 13) if pivot[m].sum() > 0]
    pivot = pivot[meses_presentes]

    x     = np.arange(len(meses_presentes))
    n_car = len(pivot)
    ancho = min(0.65, 0.8 / max(n_car, 1))

    for idx, (car, row_data) in enumerate(pivot.iterrows()):
        offset = (idx - n_car / 2) * ancho + ancho / 2
        color  = PALETTE[0] if unificado else PALETTE[idx % len(PALETTE)]
        bars   = ax.bar(x + offset, row_data.values, ancho,
                        label=car, color=color,
                        alpha=0.88, edgecolor="white", linewidth=0.5)
        # Etiqueta de valor sobre cada barra (solo si > 0)
        for bar, val in zip(bars, row_data.values):
            if val > 0:
                fs = 7 if n_car <= 4 else 6
                ax.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + ax.get_ylim()[1] * 0.005,
                        _fmt_val(val, tipo_valor),
                        ha="center", va="bottom", fontsize=fs,
                        color="#333333", rotation=0 if n_car <= 3 else 70,
                        clip_on=True)

    ax.set_xticks(x)
    ax.set_xticklabels([MESES_CORTOS[m] for m in meses_presentes], fontsize=9)
    if tipo_valor == "VALOR":
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda v, _: f"${v:,.0f}")
        )
    # Dar espacio extra arriba para las etiquetas
    ymax = ax.get_ylim()[1]
    ax.set_ylim(0, ymax * 1.18)
    titulo = (f"Cartera Total Mensual – {año}" if unificado
              else f"Valor Pendiente por Carrera y Mes – {año}")
    ax.set_title(titulo, fontsize=12, fontweight="bold", pad=12, color=COLOR_TEXT)
    ax.set_xlabel("Mes", fontsize=10)
    ylabel = (
        "Valor Pendiente ($)"
        if tipo_valor == "VALOR"
        else "Cantidad de Estudiantes"
    )

    ax.set_ylabel(ylabel, fontsize=10)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)

    if not unificado:
        if n_car <= 8:
            ax.legend(loc="upper right", fontsize=7, framealpha=0.8,
                      ncol=1 if n_car <= 5 else 2)
        else:
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15),
                      fontsize=7, framealpha=0.8, ncol=3)
            fig.subplots_adjust(bottom=0.22)

    fig.tight_layout()
    fig.canvas.draw()


def grafico_lineas(fig, df, año, carreras_sel, meses_sel, rubros_sel, tipo_valor):
    fig.clear()
    ax = fig.add_subplot(111)
    sub = _filtrar(df, año, carreras_sel, meses_sel, rubros_sel)

    if sub.empty:
        ax.text(0.5, 0.5, "Sin datos para los filtros seleccionados",
                ha="center", va="center", fontsize=12, color=COLOR_MUTED)
        fig.canvas.draw()
        return

    todas = sorted(df["CARRERA"].unique())
    pivot, unificado = _preparar_pivot(sub, carreras_sel, todas, tipo_valor)

    if pivot.empty:
        ax.text(0.5, 0.5, "Todas las carreras tienen valor $0",
                ha="center", va="center", fontsize=12, color=COLOR_MUTED)
        fig.canvas.draw()
        return

    meses_presentes = [m for m in range(1, 13) if pivot[m].sum() > 0]
    etiquetas_x = [MESES_CORTOS[m] for m in meses_presentes]

    for idx, (car, row_data) in enumerate(pivot.iterrows()):
        vals  = [row_data[m] for m in meses_presentes]
        color = PALETTE[0] if unificado else PALETTE[idx % len(PALETTE)]
        ax.plot(etiquetas_x, vals,
                marker="o", linewidth=2.2, markersize=6,
                color=color, label=car, alpha=0.9)
        # Etiqueta de valor junto a cada punto
        for xi, (etiq, val) in enumerate(zip(etiquetas_x, vals)):
            if val > 0:
                n_car = len(pivot)
                fs = 7 if n_car <= 4 else 6
                offset_y = ax.get_ylim()[1] * 0.025
                ax.annotate(_fmt_val(val, tipo_valor),
                            xy=(xi, val),
                            xytext=(0, 9), textcoords="offset points",
                            ha="center", va="bottom",
                            fontsize=fs, color=color,
                            clip_on=True)

    if tipo_valor == "VALOR":
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda v, _: f"${v:,.0f}")
        )
    ymax = ax.get_ylim()[1]
    ax.set_ylim(0, ymax * 1.15)
    titulo = (f"Tendencia Total Mensual – {año}" if unificado
              else f"Tendencia Mensual de Cartera Vencida – {año}")
    ax.set_title(titulo, fontsize=12, fontweight="bold", pad=12, color=COLOR_TEXT)
    ax.set_xlabel("Mes", fontsize=10)
    ax.set_ylabel("Valor Pendiente ($)", fontsize=10)
    ax.grid(linestyle="--", alpha=0.35)
    ax.spines[["top", "right"]].set_visible(False)

    n_car = len(pivot)
    if not unificado:
        if n_car <= 8:
            ax.legend(loc="upper right", fontsize=7, framealpha=0.8,
                      ncol=1 if n_car <= 5 else 2)
        else:
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15),
                      fontsize=7, framealpha=0.8, ncol=3)
            fig.subplots_adjust(bottom=0.22)

    fig.tight_layout()
    fig.canvas.draw()
  
# ─────────────────────────────────────────────
# INTERFAZ GRÁFICA
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Analizador de Cartera Vencida – CxC")
        self.geometry("1280x780")
        self.minsize(1024, 680)
        self.configure(bg=COLOR_BG)

        self.df        = None
        self.año_var   = tk.StringVar()
        self.tipo_graf = tk.StringVar(value="Barras")
        self.tipo_valor = tk.StringVar(value="VALOR")
        self.modalidad_var = tk.StringVar(value="TODOS")
        self.logica_var = tk.StringVar(value="")
        self.ruta_var  = tk.StringVar(value="Ningún archivo cargado")

        self._mostrar_menu_principal()

    def _limpiar_ventana(self):
        try:
            self.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        for widget in self.winfo_children():
            widget.destroy()

    def _mostrar_menu_principal(self):
        self._limpiar_ventana()
        menu = tk.Frame(self, bg=COLOR_BG)
        menu.pack(fill="both", expand=True)

        caja = tk.Frame(menu, bg=COLOR_BG)
        caja.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(caja, text="Cartera CxC", font=("Arial", 24, "bold"),
                 bg=COLOR_BG, fg=COLOR_TEXT).pack(pady=(0, 6))
        tk.Label(caja, text="Selecciona el tipo de análisis",
                 font=("Arial", 12), bg=COLOR_BG, fg=COLOR_MUTED).pack(pady=(0, 24))

        self._menu_btn(caja, LOGICA_PENDIENTES,
                       "Plantilla anterior de cartera pendiente",
                       lambda: self._seleccionar_logica(LOGICA_PENDIENTES))
        self._menu_btn(caja, LOGICA_VENCIDAS,
                       "Plantilla nueva de rubros vencidos",
                       lambda: self._seleccionar_logica(LOGICA_VENCIDAS))

    def _menu_btn(self, parent, titulo, subtitulo, cmd):
        frame = tk.Frame(parent, bg=COLOR_CARD, highlightbackground=COLOR_BORDER,
                         highlightthickness=1)
        frame.pack(fill="x", pady=8, ipadx=12, ipady=10)
        frame.bind("<Button-1>", lambda e: cmd())

        tk.Label(frame, text=titulo, font=("Arial", 13, "bold"),
                 bg=COLOR_CARD, fg=COLOR_TEXT, cursor="hand2").pack(anchor="w", padx=18, pady=(10, 0))
        tk.Label(frame, text=subtitulo, font=("Arial", 9),
                 bg=COLOR_CARD, fg=COLOR_MUTED, cursor="hand2").pack(anchor="w", padx=18, pady=(2, 10))
        for child in frame.winfo_children():
            child.bind("<Button-1>", lambda e: cmd())

    def _seleccionar_logica(self, logica):
        self.logica_var.set(logica)
        self.df = None
        self.año_var.set("")
        self.modalidad_var.set("TODOS")
        self.ruta_var.set("Ningún archivo cargado")
        self._construir_ui()

    # ── Layout principal ──────────────────────
    def _construir_ui(self):
        self._limpiar_ventana()
        # Sidebar izquierdo
        self.sidebar = tk.Frame(self, bg=COLOR_SIDEBAR, width=260)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Área derecha
        self.area_der = tk.Frame(self, bg=COLOR_BG)
        self.area_der.pack(side="left", fill="both", expand=True)

        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self):
        canvas = tk.Canvas(self.sidebar, bg=COLOR_SIDEBAR, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.sidebar, orient="vertical", command=canvas.yview)
        sb = tk.Frame(canvas, bg=COLOR_SIDEBAR)
        window_id = canvas.create_window((0, 0), window=sb, anchor="nw", width=244)

        def _configurar_scroll(_=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _ajustar_ancho(event):
            canvas.itemconfigure(window_id, width=event.width)

        sb.bind("<Configure>", _configurar_scroll)
        canvas.bind("<Configure>", _ajustar_ancho)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Logo / título
        tk.Label(sb, text="📊", font=("Arial", 28), bg=COLOR_SIDEBAR,
                 fg=COLOR_WHITE).pack(pady=(24, 4))
        titulo_sidebar = ("Cartera Pendiente" if self.logica_var.get() == LOGICA_PENDIENTES
                          else "Cartera Vencida")
        tk.Label(sb, text=titulo_sidebar, font=("Arial", 13, "bold"),
                 bg=COLOR_SIDEBAR, fg=COLOR_WHITE).pack()
        tk.Label(sb, text="Análisis CxC", font=("Arial", 10),
                 bg=COLOR_SIDEBAR, fg=COLOR_ACCENT2).pack(pady=(0, 20))

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=4)

        self._sb_section(sb, "LÓGICA")
        tk.Label(sb, text=self.logica_var.get(), bg=COLOR_SIDEBAR,
                 fg=COLOR_WHITE, font=("Arial", 8, "bold"), wraplength=210,
                 justify="left").pack(anchor="w", padx=16, pady=(0, 4))
        self._sb_btn(sb, "↩  Menú principal", self._mostrar_menu_principal,
                     color="#6C757D")

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=10)

        # ── Archivo ─────────────────────────
        self._sb_section(sb, "ARCHIVO")
        tk.Label(sb, textvariable=self.ruta_var, bg=COLOR_SIDEBAR,
                 fg=COLOR_ACCENT2, font=("Arial", 8), wraplength=220,
                 justify="left").pack(padx=16, pady=(0, 6))
        self._sb_btn(sb, "📂  Cargar Archivo", self._cargar_archivo)

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=10)

        # ── Año ──────────────────────────────
        self._sb_section(sb, "AÑO A ANALIZAR")
        self.combo_año = ttk.Combobox(sb, textvariable=self.año_var,
                                      state="readonly", width=20)
        self.combo_año.pack(padx=16, pady=(2, 10))
        self.combo_año.bind("<<ComboboxSelected>>", self._on_año_change)

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=4)

        # ── Modalidad ────────────────────────
        self._sb_section(sb, "MODALIDAD")
        self.combo_modalidad = ttk.Combobox(
            sb,
            textvariable=self.modalidad_var,
            values=("TODOS", "VALIDACIÓN", "CARRERA"),
            state="readonly",
            width=20
        )
        self.combo_modalidad.pack(padx=16, pady=(2, 10))
        self.combo_modalidad.bind("<<ComboboxSelected>>", self._on_modalidad_change)

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=4)

        # ── Exportar ─────────────────────────
        self._sb_section(sb, "EXPORTAR")
        self._sb_btn(sb, "💾  Generar Excel", self._exportar, color="#3BB273")
        self._sb_btn(sb, "📑  Reporte Val/Carr", self._exportar_validacion_carrera,
                     color="#F4A261")

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=10)
        
        # ── Tipo de valor ─────────────────────
        self._sb_section(sb, "MOSTRAR")

        tk.Radiobutton(
            sb,
            text="Valor $",
            variable=self.tipo_valor,
            value="VALOR",
            bg=COLOR_SIDEBAR,
            fg=COLOR_WHITE,
            selectcolor=COLOR_ACCENT,
            activebackground=COLOR_SIDEBAR,
            activeforeground=COLOR_WHITE,
            command=self._actualizar_grafico
        ).pack(anchor="w", padx=24)

        tk.Radiobutton(
            sb,
            text="Cantidad estudiantes",
            variable=self.tipo_valor,
            value="ESTUDIANTES",
            bg=COLOR_SIDEBAR,
            fg=COLOR_WHITE,
            selectcolor=COLOR_ACCENT,
            activebackground=COLOR_SIDEBAR,
            activeforeground=COLOR_WHITE,
            command=self._actualizar_grafico
        ).pack(anchor="w", padx=24)

        ttk.Separator(sb, orient="horizontal").pack(fill="x", padx=16, pady=10)

        # ── Gráficos ─────────────────────────
        self._sb_section(sb, "TIPO DE GRÁFICO")
        for t in ("Barras", "Líneas"):
            tk.Radiobutton(sb, text=t, variable=self.tipo_graf, value=t,
                           bg=COLOR_SIDEBAR, fg=COLOR_WHITE,
                           selectcolor=COLOR_ACCENT, activebackground=COLOR_SIDEBAR,
                           activeforeground=COLOR_WHITE,
                           command=self._actualizar_grafico).pack(anchor="w", padx=24)

        # Espaciador + versión
        tk.Frame(sb, bg=COLOR_SIDEBAR).pack(expand=True, fill="y")
        tk.Label(sb, text="v1.0 · 2026", font=("Arial", 8),
                 bg=COLOR_SIDEBAR, fg="#4A6B8A").pack(pady=12)

    def _sb_section(self, parent, texto):
        tk.Label(parent, text=texto, font=("Arial", 8, "bold"),
                 bg=COLOR_SIDEBAR, fg="#7BA7BC").pack(anchor="w", padx=16, pady=(8, 2))

    def _sb_btn(self, parent, texto, cmd, color=COLOR_BTN):
        btn = tk.Button(parent, text=texto, command=cmd,
                        bg=color, fg=COLOR_WHITE,
                        font=("Arial", 9, "bold"),
                        relief="flat", cursor="hand2",
                        padx=10, pady=7, bd=0)
        btn.pack(fill="x", padx=16, pady=3)
        btn.bind("<Enter>", lambda e: btn.config(bg=COLOR_BTN_HOV))
        btn.bind("<Leave>", lambda e: btn.config(bg=color))

    def _build_main(self):
        area = self.area_der

        # ── Tabs ────────────────────────────
        self.notebook = ttk.Notebook(area)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Pestaña Resumen
        self.tab_resumen = tk.Frame(self.notebook, bg=COLOR_BG)
        self.notebook.add(self.tab_resumen, text="  📋 Resumen  ")
        self._build_tab_resumen()

        # Pestaña Gráficos
        self.tab_graficos = tk.Frame(self.notebook, bg=COLOR_BG)
        self.notebook.add(self.tab_graficos, text="  📈 Gráficos  ")
        self._build_tab_graficos()

    # ── Tab Resumen ───────────────────────────
    def _build_tab_resumen(self):
        t = self.tab_resumen

        # KPI cards al tope
        self.frame_kpis = tk.Frame(t, bg=COLOR_BG)
        self.frame_kpis.pack(fill="x", padx=8, pady=(8, 4))

        # Tabla
        frame_tabla = tk.Frame(t, bg=COLOR_BG)
        frame_tabla.pack(fill="both", expand=True, padx=8, pady=4)

        cols = ("Carrera", "Total $", "Estudiantes",
                *[MESES_CORTOS[m] for m in range(1, 13)])

        self.tree = ttk.Treeview(frame_tabla, columns=cols,
                                 show="headings", height=20)
        for col in cols:
            anchor = "w" if col == "Carrera" else "center"
            width  = 240 if col == "Carrera" else 80
            self.tree.heading(col, text=col,
                              command=lambda c=col: self._sort_tree(c))
            self.tree.column(col, width=width, anchor=anchor, minwidth=50)

        vsb = ttk.Scrollbar(frame_tabla, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(frame_tabla, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)

        # Tags de color alternado
        self.tree.tag_configure("par",   background="#F8FAFC")
        self.tree.tag_configure("impar", background="#FFFFFF")
        self.tree.tag_configure("total", background="#1E3A5F",
                                foreground=COLOR_WHITE,
                                font=("Arial", 9, "bold"))

    def _build_tab_graficos(self):
        t = self.tab_graficos

        # Panel de filtros
        frame_filtros = tk.LabelFrame(t, text=" Filtros ",
                                      bg=COLOR_BG, fg=COLOR_TEXT,
                                      font=("Arial", 9, "bold"))
        frame_filtros.pack(fill="x", padx=10, pady=(8, 4))

        # ── Filtro Carreras ──────────────────
        tk.Label(frame_filtros, text="Carreras:", bg=COLOR_BG,
                 font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="nw",
                                                  padx=(8,4), pady=4)
        self.lb_carreras = tk.Listbox(frame_filtros, selectmode="multiple",
                                      height=5, width=32, exportselection=False,
                                      font=("Arial", 8))
        self.lb_carreras.grid(row=0, column=1, padx=4, pady=4, sticky="w")
        sb_c = ttk.Scrollbar(frame_filtros, command=self.lb_carreras.yview)
        sb_c.grid(row=0, column=2, sticky="ns", pady=4)
        self.lb_carreras.config(yscrollcommand=sb_c.set)

        # ── Filtro Meses ─────────────────────
        tk.Label(frame_filtros, text="Meses:", bg=COLOR_BG,
                 font=("Arial", 9, "bold")).grid(row=0, column=3, sticky="nw",
                                                  padx=(16,4), pady=4)
        frame_meses = tk.Frame(frame_filtros, bg=COLOR_BG)
        frame_meses.grid(row=0, column=4, padx=4, pady=4, sticky="w")
        self.vars_meses = {}
        for i, m in enumerate(range(1, 13)):
            var = tk.BooleanVar(value=True)
            self.vars_meses[m] = var
            chk = tk.Checkbutton(frame_meses, text=MESES_CORTOS[m],
                                 variable=var, bg=COLOR_BG,
                                 font=("Arial", 8),
                                 command=self._actualizar_grafico)
            chk.grid(row=i // 6, column=i % 6, sticky="w", padx=3)

        # ── Filtro Rubro ─────────────────────
        tk.Label(frame_filtros, text="Tipo rubro:", bg=COLOR_BG,
                 font=("Arial", 9, "bold")).grid(row=0, column=5, sticky="nw",
                                                  padx=(16,4), pady=4)
        frame_rub = tk.Frame(frame_filtros, bg=COLOR_BG)
        frame_rub.grid(row=0, column=6, padx=4, pady=4, sticky="nw")
        self.vars_rubros = {}
        for rub in ("MATRÍCULA", "ARANCEL", "INSCRIPCIÓN", "OTRO"):
            var = tk.BooleanVar(value=True)
            self.vars_rubros[rub] = var
            chk = tk.Checkbutton(frame_rub, text=rub, variable=var,
                                  bg=COLOR_BG, font=("Arial", 8),
                                  command=self._actualizar_grafico)
            chk.pack(anchor="w")

        # Botón aplicar
        tk.Button(frame_filtros, text="🔄 Actualizar",
                  command=self._actualizar_grafico,
                  bg=COLOR_BTN, fg=COLOR_WHITE,
                  font=("Arial", 9, "bold"), relief="flat",
                  cursor="hand2", padx=10, pady=5).grid(
            row=0, column=7, padx=12, pady=4)

        # Bind listbox
        self.lb_carreras.bind("<<ListboxSelect>>",
                              lambda e: self._actualizar_grafico())

        # ── Área del gráfico ─────────────────
        self.fig = Figure(figsize=(10, 5), facecolor=COLOR_BG)
        self.canvas = FigureCanvasTkAgg(self.fig, master=t)
        self.canvas.get_tk_widget().pack(fill="both", expand=True,
                                         padx=10, pady=4)

        toolbar_frame = tk.Frame(t, bg=COLOR_BG)
        toolbar_frame.pack(fill="x", padx=10)
        NavigationToolbar2Tk(self.canvas, toolbar_frame)

    # ── Acciones ─────────────────────────────
    def _cargar_archivo(self):
        logica = self.logica_var.get()
        plantilla = ("matriz_cartera_vencida.xlsx" if logica == LOGICA_PENDIENTES
                     else "matriz_rubros_vencidos*.xlsx")
        ruta = filedialog.askopenfilename(
            title=f"Seleccionar archivo - {logica}",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            self.df = cargar_datos(ruta, logica)
            años    = obtener_años(self.df)
            self.combo_año["values"] = años
            año_actual = datetime.now().year
            default = str(año_actual) if año_actual in años else str(años[-1])
            self.año_var.set(default)
            self.ruta_var.set(os.path.basename(ruta))
            self._on_año_change()
            messagebox.showinfo("Archivo cargado",
                                f"✓ {len(self.df):,} registros cargados.\n"
                                f"Lógica: {logica}\n"
                                f"Años disponibles: {', '.join(map(str, años))}")
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))

    def _on_año_change(self, *_):
        if self.df is None or not self.año_var.get():
            return
        año = int(self.año_var.get())
        modalidad = self.modalidad_var.get()
        pv, pe, carreras = calcular_pivots(self.df, año, modalidad)
        self._poblar_tabla(pv, pe, carreras)
        self._poblar_kpis(pv, pe, año)
        self._actualizar_listbox_carreras(carreras)
        self._actualizar_grafico()

    def _on_modalidad_change(self, *_):
        self._on_año_change()

    def _poblar_kpis(self, pv, pe, año):
        for w in self.frame_kpis.winfo_children():
            w.destroy()

        total_val = pv.values.sum()
        # Conteo real: estudiantes únicos por IDENTIFICACIÓN en el año
        sub_año = filtrar_modalidad(self.df[self.df["AÑO"] == año], self.modalidad_var.get())
        total_est = sub_año["IDENTIFICACIÓN"].nunique()
        mes_actual = datetime.now().month
        val_mes = pv[mes_actual].sum() if mes_actual in pv.columns else 0
        carrera_top = pv.sum(axis=1).idxmax() if pv.sum(axis=1).max() > 0 else "–"

        kpis = [
            ("💰 Total Cartera", f"${total_val:,.2f}", COLOR_ACCENT),
            ("👥 Estudiantes", f"{int(total_est):,}", "#3BB273"),
            (f"📅 {MESES.get(mes_actual,'Mes actual')}", f"${val_mes:,.2f}", "#E84855"),
            ("🏆 Mayor Cartera", carrera_top[:22] + "…" if len(carrera_top) > 22
             else carrera_top, "#F4A261"),
        ]
        for label, valor, color in kpis:
            card = tk.Frame(self.frame_kpis, bg=COLOR_CARD,
                            relief="flat", bd=1)
            card.pack(side="left", expand=True, fill="x", padx=5, pady=4)
            tk.Frame(card, bg=color, height=4).pack(fill="x")
            tk.Label(card, text=label, font=("Arial", 9),
                     bg=COLOR_CARD, fg=COLOR_MUTED).pack(pady=(6, 0))
            tk.Label(card, text=valor, font=("Arial", 13, "bold"),
                     bg=COLOR_CARD, fg=COLOR_TEXT, wraplength=180).pack(pady=(0, 8))

    def _poblar_tabla(self, pv, pe, carreras):
        for row in self.tree.get_children():
            self.tree.delete(row)

        for i, car in enumerate(carreras):
            total_v = pv.loc[car].sum()
            total_e = int(pe.loc[car].max())  # estudiantes únicos en algún mes
            vals_mes = [f"${pv.at[car,m]:,.0f}" if pv.at[car,m] > 0 else "–"
                        for m in range(1, 13)]
            tag = "par" if i % 2 == 0 else "impar"
            self.tree.insert("", "end",
                             values=(car, f"${total_v:,.2f}", total_e, *vals_mes),
                             tags=(tag,))

        # Fila total
        tot_v = pv.sum(axis=0)
        tot_e_mes = [int(pe[m].sum()) for m in range(1, 13)]
        vals_mes_tot = [f"${tot_v[m]:,.0f}" if tot_v[m] > 0 else "–"
                        for m in range(1, 13)]
        self.tree.insert("", "end",
                         values=("TOTAL", f"${pv.values.sum():,.2f}",
                                 int(pe.values.sum()), *vals_mes_tot),
                         tags=("total",))

    def _actualizar_listbox_carreras(self, carreras):
        self.lb_carreras.delete(0, "end")
        for car in carreras:
            self.lb_carreras.insert("end", car)

    def _actualizar_grafico(self, *_):

        if self.df is None or not self.año_var.get():
            return

        año = int(self.año_var.get())

        sel_idx = self.lb_carreras.curselection()

        if sel_idx:
            carreras_sel = [self.lb_carreras.get(i) for i in sel_idx]
        else:
            carreras_sel = []

        meses_sel = [
            m for m, v in self.vars_meses.items()
            if v.get()
        ]

        rubros_sel = [
            r for r, v in self.vars_rubros.items()
            if v.get()
        ]

        tipo_valor = self.tipo_valor.get()
        tipo_graf  = self.tipo_graf.get()

        if tipo_graf == "Barras":

            grafico_barras(
                self.fig,
                filtrar_modalidad(self.df, self.modalidad_var.get()),
                año,
                carreras_sel,
                meses_sel,
                rubros_sel,
                tipo_valor
            )

        elif tipo_graf == "Líneas":

            grafico_lineas(
                self.fig,
                filtrar_modalidad(self.df, self.modalidad_var.get()),
                año,
                carreras_sel,
                meses_sel,
                rubros_sel,
                tipo_valor
            )
   
    def _exportar(self):
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo.")
            return
        if not self.año_var.get():
            messagebox.showwarning("Sin año", "Selecciona un año.")
            return

        año = int(self.año_var.get())
        modalidad = self.modalidad_var.get()
        pv, pe, _ = calcular_pivots(self.df, año, modalidad)

        ruta = filedialog.asksaveasfilename(
            title="Guardar análisis como…",
            defaultextension=".xlsx",
            initialfile=f"Analisis_CxC_{self._slug_logica()}_{modalidad.title()}_{año}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            exportar_excel(pv, pe, año, ruta)
            messagebox.showinfo("¡Listo!",
                                f"Excel generado exitosamente:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))

    def _exportar_validacion_carrera(self):
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo.")
            return
        if not self.año_var.get():
            messagebox.showwarning("Sin año", "Selecciona un año.")
            return

        año = int(self.año_var.get())
        modalidad = self.modalidad_var.get()
        ruta = filedialog.asksaveasfilename(
            title="Guardar reporte Validación/Carrera como…",
            defaultextension=".xlsx",
            initialfile=f"Validacion_Carrera_{self._slug_logica()}_{modalidad.title()}_{año}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            exportar_validacion_carrera(self.df, año, modalidad, ruta)
            messagebox.showinfo("¡Listo!",
                                f"Reporte generado exitosamente:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error al generar reporte", str(e))

    def _slug_logica(self):
        return "Pendientes" if self.logica_var.get() == LOGICA_PENDIENTES else "Vencidas"

    def _sort_tree(self, col):
        """Ordena la tabla al hacer clic en un encabezado."""
        items = [(self.tree.set(k, col), k)
                 for k in self.tree.get_children("")
                 if "total" not in self.tree.item(k, "tags")]
        try:
            items.sort(key=lambda x: float(x[0].replace("$","").replace(",","")
                                           .replace("–","0")))
        except ValueError:
            items.sort()
        for idx, (_, k) in enumerate(items):
            self.tree.move(k, "", idx)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    plt.style.use("seaborn-v0_8-whitegrid")
    app = App()
    app.mainloop()

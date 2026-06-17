#!/usr/bin/env python3
"""
Dashboard de Entrega de Resultados
Procesa archivos XLS semanales y genera tablero de métricas en Excel + JSON.

Uso:
  python procesar.py                         # procesa todos los XLS en datos/
  python procesar.py --semana 24             # solo semana 24
  python procesar.py --semana 23 24          # semanas 23 y 24 combinadas
  python procesar.py --archivo datos/Semana_24.xls ...
"""

import sys
import os
import re
import json
import argparse
import datetime
from pathlib import Path

import xlrd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuración ─────────────────────────────────────────────────────────────

SUCURSALES = {
    "LM": "López Mateos",
    "AC": "Acueducto",
    "CH": "Chamizal",
    "GP": "Guadalupe",
    "CM": "Clínica de la Mujer",
    "MP": "Medicina Preventiva",
    "LA": "Las Águilas",
}

ORDEN_UNIDADES = ["General", "LM", "AC", "CH", "GP", "CM", "MP", "LA"]

MODALIDADES = {
    "Cardiología":  {"depto": "Cardiología",               "lim_ag": 7, "lim_te": 7},
    "Densitometría":{"depto": "Densitometría",             "lim_ag": 2, "lim_te": 2},
    "Ecosonografía":{"depto": "Ecosonografía",             "lim_ag": 7, "lim_te": 2},
    "Mamografía":   {"depto": "Mamografía",                "lim_ag": 3, "lim_te": 3},
    "Radiología":   {"depto": "Radiología",                "lim_ag": 2, "lim_te": 2},
    "RM":           {"depto": "Resonancia Magnética",      "lim_ag": 5, "lim_te": 5},
    "Tomografía":   {"depto": "Tomografía Axial Computada","lim_ag": 3, "lim_te": 3},
}

DEPTO_A_MOD = {v["depto"]: k for k, v in MODALIDADES.items()}

RM_48H  = ["tórax a muslo", "bodyscan", "valoración funcional", "oncológica",
           "funcional", "tractografía", "difusión", "perfusión"]
RM_5H   = ["columna", "cervical", "dorsal", "lumbar", "hombro", "cadera",
           "sacroilíaca", "tobillo", "codo", "rodilla"]
RM_5H_EXCL = ["muñeca", "mano"]

DATOS_DIR  = Path(__file__).parent / "datos"
OUTPUT_DIR = Path(__file__).parent / "output"

# ── Colores tema ──────────────────────────────────────────────────────────────

COLOR_AZUL_OSCURO = "1F3864"
COLOR_AZUL_MEDIO  = "2E5FA3"
COLOR_AZUL_CLARO  = "BDD7EE"
COLOR_VERDE       = "70AD47"
COLOR_AMARILLO    = "FFD966"
COLOR_ROJO        = "FF0000"
COLOR_GRIS        = "D9D9D9"
COLOR_BLANCO      = "FFFFFF"
COLOR_ENCABEZADO  = "203864"
COLOR_FILA_PAR    = "EBF3FB"
COLOR_VERDE_CLARO = "E2EFDA"


# ── Utilidades ────────────────────────────────────────────────────────────────

def parse_af_hours(val):
    """Convierte string A-F ' HH:MM' o '-HH:MM' a float horas. None si vacío."""
    if not val or not isinstance(val, str):
        return None
    val = val.strip()
    if not val:
        return None
    neg = val.startswith("-")
    val = val.lstrip("-").strip()
    m = re.match(r"(\d+):(\d+)", val)
    if not m:
        return None
    hours = int(m.group(1)) + int(m.group(2)) / 60
    return -hours if neg else hours


def classify_rm(estudio: str) -> str:
    """Clasifica un estudio RM en '5h', '24h' o '48h'."""
    s = estudio.lower()
    if any(kw in s for kw in RM_48H):
        return "48h"
    if any(kw in s for kw in RM_5H) and not any(ex in s for ex in RM_5H_EXCL):
        return "5h"
    return "24h"


def semana_de_archivo(path: Path) -> int | None:
    """Extrae número de semana del nombre del archivo (Semana_NN.xls)."""
    m = re.search(r"[Ss]emana[_\s]*(\d+)", path.stem)
    return int(m.group(1)) if m else None


def fmt_horas(val):
    """Formatea horas float a 'HH:MM' para visualización."""
    if val is None:
        return "-"
    neg = val < 0
    val = abs(val)
    h = int(val)
    mn = int(round((val - h) * 60))
    return f"{'-' if neg else ''}{h:02d}:{mn:02d}"


# ── Lectura de XLS ────────────────────────────────────────────────────────────

def leer_xls(filepath: Path, semana: int | None = None) -> list[dict]:
    """Lee un XLS y retorna lista de dicts. semana proviene del nombre del archivo."""
    wb = xlrd.open_workbook(str(filepath))
    ws = wb.sheet_by_index(0)
    if semana is None:
        semana = semana_de_archivo(filepath)

    rows = []
    for i in range(9, ws.nrows):
        row = ws.row_values(i)
        estudio  = str(row[2]).strip() if row[2] else ""
        depto    = str(row[3]).strip() if row[3] else ""
        sucursal = str(row[44]).strip() if row[44] else ""

        if not depto or not sucursal:
            continue

        ag = row[21]
        ag = float(ag) if ag else None

        fecha_alta_raw   = row[19]
        fecha_recibe_raw = row[39]
        af_str = str(row[41]).strip() if row[41] else ""

        te = None
        if (isinstance(fecha_recibe_raw, float) and isinstance(fecha_alta_raw, float)
                and fecha_recibe_raw and fecha_alta_raw):
            te = (fecha_recibe_raw - fecha_alta_raw) * 24

        af_hours = parse_af_hours(af_str)

        fecha = None
        if isinstance(fecha_alta_raw, float) and fecha_alta_raw:
            fecha = xlrd.xldate_as_datetime(fecha_alta_raw, wb.datemode).date()

        rows.append({
            "estudio":  estudio,
            "depto":    depto,
            "sucursal": sucursal,
            "ag":       ag,
            "af":       af_hours,
            "te":       te,
            "fecha":    fecha,
            "semana":   semana,
            "archivo":  filepath.name,
        })
    return rows


# ── Cálculo de métricas ───────────────────────────────────────────────────────

def es_incumplimiento(af):
    """Incumplimiento real: -24 < af < 0."""
    return af is not None and -24 < af < 0


def calcular_metricas(rows: list[dict]) -> dict:
    """Calcula todas las métricas para un conjunto de filas."""
    modal_rows = [r for r in rows if r["depto"] in DEPTO_A_MOD]
    total_modal = len(modal_rows)
    incump = sum(1 for r in modal_rows if es_incumplimiento(r["af"]))
    indicador = ((total_modal - incump) / total_modal * 100) if total_modal > 0 else None

    result = {
        "indicador":          indicador,
        "estudios":           total_modal,
        "incumplimiento":     incump,
        "modalidades_total":  len(rows),
    }

    for mod, cfg in MODALIDADES.items():
        depto   = cfg["depto"]
        lim_ag  = cfg["lim_ag"]
        lim_te  = cfg["lim_te"]

        mod_rows = [r for r in rows if r["depto"] == depto]
        ok_rows  = [r for r in mod_rows if r["ag"] is not None and r["ag"] <= lim_ag]
        count    = len(ok_rows)

        te_vals  = [r["te"] for r in ok_rows
                    if r["te"] is not None and 0 <= r["te"] <= lim_te]
        te_avg   = (sum(te_vals) / len(te_vals)) if te_vals else None

        result[mod]          = count
        result[f"TE_{mod}"]  = te_avg

        if mod == "RM":
            result["RM_5h"]  = sum(1 for r in ok_rows if classify_rm(r["estudio"]) == "5h")
            result["RM_24h"] = sum(1 for r in ok_rows if classify_rm(r["estudio"]) == "24h")
            result["RM_48h"] = sum(1 for r in ok_rows if classify_rm(r["estudio"]) == "48h")

    return result


def generar_tabla(rows: list[dict], fecha_reporte=None) -> list[dict]:
    """Genera lista de filas del dashboard (una por unidad)."""
    tabla = []
    for unidad in ORDEN_UNIDADES:
        if unidad == "General":
            subset = rows
            fecha  = str(fecha_reporte) if fecha_reporte else ""
        else:
            nombre_suc = SUCURSALES[unidad]
            subset = [r for r in rows if r["sucursal"] == nombre_suc]
            fecha  = ""

        m = calcular_metricas(subset)
        tabla.append({"Unidad": unidad, "Fecha": fecha, **m})
    return tabla


# ── Escritura Excel ───────────────────────────────────────────────────────────

def thin_border():
    t = Side(style="thin", color="AAAAAA")
    return Border(left=t, right=t, top=t, bottom=t)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def bold_font(size=11, color=COLOR_BLANCO):
    return Font(bold=True, size=size, color=color)


def normal_font(size=10, color="000000"):
    return Font(size=size, color=color)


def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def left():
    return Alignment(horizontal="left", vertical="center")


COLUMNAS_DASH = [
    ("Fecha",             16),
    ("Unidad",            10),
    ("Indicador %",       12),
    ("Estudios",          10),
    ("Incumplimiento",    14),
    ("Total Registros",   13),
    ("Cardiología",       12),
    ("TE Cardiología",    14),
    ("Densitometría",     13),
    ("TE Densitometría",  15),
    ("Ecosonografía",     13),
    ("TE Ecosonografía",  15),
    ("Mamografía",        12),
    ("TE Mamografía",     14),
    ("Radiología",        12),
    ("TE Radiología",     14),
    ("RM",                8),
    ("TE RM",             10),
    ("RM 5h",             9),
    ("RM 24h",            9),
    ("RM 48h",            9),
    ("Tomografía",        12),
    ("TE Tomografía",     14),
]

# mapa de nombre de columna → clave en el dict de métricas
COL_KEY = {
    "Indicador %":      "indicador",
    "Estudios":         "estudios",
    "Incumplimiento":   "incumplimiento",
    "Total Registros":  "modalidades_total",
    "Cardiología":      "Cardiología",
    "TE Cardiología":   "TE_Cardiología",
    "Densitometría":    "Densitometría",
    "TE Densitometría": "TE_Densitometría",
    "Ecosonografía":    "Ecosonografía",
    "TE Ecosonografía": "TE_Ecosonografía",
    "Mamografía":       "Mamografía",
    "TE Mamografía":    "TE_Mamografía",
    "Radiología":       "Radiología",
    "TE Radiología":    "TE_Radiología",
    "RM":               "RM",
    "TE RM":            "TE_RM",
    "RM 5h":            "RM_5h",
    "RM 24h":           "RM_24h",
    "RM 48h":           "RM_48h",
    "Tomografía":       "Tomografía",
    "TE Tomografía":    "TE_Tomografía",
}


def escribir_hoja_dashboard(ws_dash, titulo: str, tabla: list[dict]):
    """Escribe la hoja Dashboard con formato completo."""
    # — Título —
    ws_dash.merge_cells(f"A1:{get_column_letter(len(COLUMNAS_DASH))}1")
    celda = ws_dash.cell(1, 1, titulo)
    celda.fill      = fill(COLOR_AZUL_OSCURO)
    celda.font      = Font(bold=True, size=14, color=COLOR_BLANCO)
    celda.alignment = center()
    ws_dash.row_dimensions[1].height = 30

    # — Encabezados —
    for col_idx, (nombre, ancho) in enumerate(COLUMNAS_DASH, 1):
        celda = ws_dash.cell(2, col_idx, nombre)
        celda.fill      = fill(COLOR_AZUL_MEDIO)
        celda.font      = bold_font(10)
        celda.alignment = center()
        celda.border    = thin_border()
        ws_dash.column_dimensions[get_column_letter(col_idx)].width = ancho
    ws_dash.row_dimensions[2].height = 40

    # — Filas de datos —
    for fila_idx, fila in enumerate(tabla, 3):
        es_general = fila["Unidad"] == "General"
        color_fondo = COLOR_AZUL_CLARO if es_general else (
            COLOR_FILA_PAR if fila_idx % 2 == 0 else COLOR_BLANCO
        )

        for col_idx, (nombre, _) in enumerate(COLUMNAS_DASH, 1):
            if nombre == "Fecha":
                val = fila.get("Fecha", "")
            elif nombre == "Unidad":
                val = fila["Unidad"]
            else:
                key  = COL_KEY.get(nombre)
                val  = fila.get(key)

            # Formateo de valores
            if nombre == "Indicador %":
                texto = f"{val:.1f}%" if val is not None else "-"
                # Color según cumplimiento
                if val is not None:
                    color_ind = COLOR_VERDE if val >= 95 else (
                        COLOR_AMARILLO if val >= 90 else "FF4444")
                else:
                    color_ind = None
            elif nombre.startswith("TE "):
                texto = fmt_horas(val)
                color_ind = None
            elif val is None:
                texto = "-"
                color_ind = None
            elif isinstance(val, float) and not nombre.startswith("RM"):
                texto = f"{val:.0f}"
                color_ind = None
            else:
                texto = str(val) if val is not None else "-"
                color_ind = None

            celda = ws_dash.cell(fila_idx, col_idx, texto)
            celda.fill      = fill(color_fondo)
            celda.border    = thin_border()
            celda.alignment = center() if col_idx > 2 else left()

            if nombre == "Indicador %" and color_ind:
                celda.font = Font(bold=True, size=11, color=color_ind)
            elif es_general:
                celda.font = Font(bold=True, size=10, color="1F3864")
            else:
                celda.font = normal_font()

        ws_dash.row_dimensions[fila_idx].height = 22

    # Freeze encabezado
    ws_dash.freeze_panes = "C3"


def escribir_hoja_detalle(ws_det, rows: list[dict]):
    """Hoja con detalle de registros usados en el cálculo."""
    encabezados = ["Semana", "Fecha", "Sucursal", "Departamento", "Estudio",
                   "A-G (h)", "A-F", "TE (h)", "Incumplimiento", "Clasificación RM"]
    anchos      = [9, 12, 18, 22, 50, 10, 10, 10, 14, 16]

    for col_idx, (enc, ancho) in enumerate(zip(encabezados, anchos), 1):
        celda = ws_det.cell(1, col_idx, enc)
        celda.fill      = fill(COLOR_ENCABEZADO)
        celda.font      = bold_font(10)
        celda.alignment = center()
        celda.border    = thin_border()
        ws_det.column_dimensions[get_column_letter(col_idx)].width = ancho
    ws_det.row_dimensions[1].height = 30

    for fila_idx, r in enumerate(rows, 2):
        incump_val = "Sí" if es_incumplimiento(r["af"]) else "No"
        rm_cls     = classify_rm(r["estudio"]) if r["depto"] == "Resonancia Magnética" else ""
        color_fondo = COLOR_FILA_PAR if fila_idx % 2 == 0 else COLOR_BLANCO

        valores = [
            r.get("semana", ""),
            str(r["fecha"]) if r["fecha"] else "",
            r["sucursal"],
            r["depto"],
            r["estudio"],
            f"{r['ag']:.1f}" if r["ag"] is not None else "-",
            fmt_horas(r["af"]),
            f"{r['te']:.2f}" if r["te"] is not None else "-",
            incump_val,
            rm_cls,
        ]
        for col_idx, val in enumerate(valores, 1):
            celda = ws_det.cell(fila_idx, col_idx, val)
            celda.fill      = fill(color_fondo)
            celda.font      = normal_font(9)
            celda.alignment = left()
            celda.border    = thin_border()
            if incump_val == "Sí" and col_idx == 9:
                celda.font = Font(size=9, color=COLOR_ROJO, bold=True)

    ws_det.freeze_panes = "A2"
    ws_det.auto_filter.ref = ws_det.dimensions


def escribir_hoja_semanas(ws_sem, todas_las_filas: list[dict], semanas: list[int]):
    """Resumen por semana: una sub-tabla por semana."""
    fila_actual = 1
    for semana in sorted(semanas):
        rows_sem = [r for r in todas_las_filas if r.get("semana") == semana]
        if not rows_sem:
            continue
        fechas = [r["fecha"] for r in rows_sem if r["fecha"]]
        rango  = f"{min(fechas)} al {max(fechas)}" if fechas else ""
        titulo = f"Semana {semana}  —  {rango}"

        tabla = generar_tabla(rows_sem, max(fechas) if fechas else None)

        # Escribir subtítulo
        ws_sem.merge_cells(
            start_row=fila_actual, start_column=1,
            end_row=fila_actual, end_column=len(COLUMNAS_DASH))
        celda = ws_sem.cell(fila_actual, 1, titulo)
        celda.fill      = fill(COLOR_AZUL_MEDIO)
        celda.font      = bold_font(11)
        celda.alignment = center()
        ws_sem.row_dimensions[fila_actual].height = 25
        fila_actual += 1

        # Encabezados
        for col_idx, (nombre, ancho) in enumerate(COLUMNAS_DASH, 1):
            celda = ws_sem.cell(fila_actual, col_idx, nombre)
            celda.fill      = fill(COLOR_AZUL_CLARO)
            celda.font      = Font(bold=True, size=9, color="1F3864")
            celda.alignment = center()
            celda.border    = thin_border()
            ws_sem.column_dimensions[get_column_letter(col_idx)].width = ancho
        ws_sem.row_dimensions[fila_actual].height = 30
        fila_actual += 1

        # Datos
        for fila in tabla:
            es_general  = fila["Unidad"] == "General"
            color_fondo = COLOR_VERDE_CLARO if es_general else (
                COLOR_FILA_PAR if fila_actual % 2 == 0 else COLOR_BLANCO)

            for col_idx, (nombre, _) in enumerate(COLUMNAS_DASH, 1):
                if nombre == "Fecha":
                    val = fila.get("Fecha", "")
                elif nombre == "Unidad":
                    val = fila["Unidad"]
                else:
                    key = COL_KEY.get(nombre)
                    val = fila.get(key)

                if nombre == "Indicador %":
                    texto = f"{val:.1f}%" if val is not None else "-"
                elif nombre.startswith("TE "):
                    texto = fmt_horas(val)
                elif val is None:
                    texto = "-"
                else:
                    texto = str(val) if not isinstance(val, float) else f"{val:.0f}"

                celda = ws_sem.cell(fila_actual, col_idx, texto)
                celda.fill      = fill(color_fondo)
                celda.border    = thin_border()
                celda.alignment = center() if col_idx > 2 else left()
                celda.font      = (Font(bold=True, size=9, color="1F3864")
                                   if es_general else Font(size=9))

            ws_sem.row_dimensions[fila_actual].height = 18
            fila_actual += 1

        fila_actual += 2  # espacio entre semanas


# ── Exportar JSON ─────────────────────────────────────────────────────────────

def exportar_json(tabla_acum: list[dict], tabla_por_semana: dict,
                  rows: list[dict], ruta: Path):
    """Genera JSON de respaldo con todos los datos."""
    def clean(obj):
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(v) for v in obj]
        if isinstance(obj, datetime.date):
            return obj.isoformat()
        if isinstance(obj, float) and obj != obj:  # NaN
            return None
        return obj

    data = {
        "generado":      datetime.datetime.now().isoformat(),
        "acumulado":     tabla_acum,
        "por_semana":    tabla_por_semana,
        "total_registros": len(rows),
    }
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(clean(data), f, ensure_ascii=False, indent=2, default=str)


def exportar_web(tabla_acum, tabla_por_semana, todas_las_filas,
                 semanas_presentes, fecha_min, fecha_max):
    """Genera docs/data.js con los datos para el dashboard web (GitHub Pages)."""
    def clean(obj):
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(v) for v in obj]
        if isinstance(obj, datetime.date):
            return obj.isoformat()
        if isinstance(obj, float) and obj != obj:
            return None
        return obj

    # Rango de fechas por semana
    rangos = {}
    for sem in semanas_presentes:
        fechas = [r["fecha"] for r in todas_las_filas
                  if r.get("semana") == sem and r["fecha"]]
        if fechas:
            rangos[str(sem)] = {"inicio": min(fechas).isoformat(),
                                "fin": max(fechas).isoformat()}

    data = {
        "generado":   datetime.datetime.now().isoformat(),
        "rango":      {"inicio": fecha_min.isoformat(), "fin": fecha_max.isoformat()},
        "semanas":    sorted(str(s) for s in semanas_presentes),
        "rangos":     rangos,
        "acumulado":  tabla_acum,
        "por_semana": tabla_por_semana,
        "sucursales": SUCURSALES,
        "total_registros": len(todas_las_filas),
    }

    DOCS_DIR = Path(__file__).parent / "docs"
    DOCS_DIR.mkdir(exist_ok=True)
    ruta = DOCS_DIR / "data.js"
    with open(ruta, "w", encoding="utf-8") as f:
        f.write("// Generado automáticamente por procesar.py — no editar a mano\n")
        f.write("window.DASHBOARD_DATA = ")
        json.dump(clean(data), f, ensure_ascii=False, indent=2, default=str)
        f.write(";\n")
    return ruta


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Dashboard de Entrega de Resultados")
    parser.add_argument("--semana",   type=int, nargs="+",
                        help="Número(s) de semana a procesar (default: todas)")
    parser.add_argument("--archivo",  nargs="+",
                        help="Archivos XLS específicos a procesar")
    parser.add_argument("--salida",   default=None,
                        help="Nombre del archivo de salida (sin extensión)")
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(exist_ok=True)

    # Determinar archivos a leer
    if args.archivo:
        archivos = [Path(a) for a in args.archivo]
    else:
        archivos = sorted(DATOS_DIR.glob("*.xls")) + sorted(DATOS_DIR.glob("*.xlsx"))

    if not archivos:
        print("No se encontraron archivos XLS en datos/. "
              "Coloca los archivos ahí o usa --archivo.")
        sys.exit(1)

    # Filtrar por semana si se especificó
    if args.semana:
        archivos = [a for a in archivos
                    if semana_de_archivo(a) in args.semana or semana_de_archivo(a) is None]

    print(f"Procesando {len(archivos)} archivo(s):")
    todas_las_filas = []
    semanas_presentes = set()
    for a in archivos:
        print(f"  • {a.name}", end="  ")
        filas = leer_xls(a)
        sem   = semana_de_archivo(a)
        if sem:
            semanas_presentes.add(sem)
        print(f"({len(filas)} registros)")
        todas_las_filas.extend(filas)

    if not todas_las_filas:
        print("No se obtuvieron registros válidos.")
        sys.exit(1)

    # Fechas para nombre de archivo
    fechas_validas = [r["fecha"] for r in todas_las_filas if r["fecha"]]
    fecha_max  = max(fechas_validas) if fechas_validas else datetime.date.today()
    fecha_min  = min(fechas_validas) if fechas_validas else fecha_max

    # Título del reporte
    if semanas_presentes:
        sems_str = ", ".join(f"Sem {s}" for s in sorted(semanas_presentes))
        titulo   = f"Dashboard de Entrega de Resultados  —  {sems_str}  ({fecha_min} al {fecha_max})"
    else:
        titulo   = f"Dashboard de Entrega de Resultados  —  {fecha_min} al {fecha_max}"

    # Tabla acumulada
    tabla_acum = generar_tabla(todas_las_filas, fecha_max)

    # Tabla por semana
    tabla_por_semana = {}
    for sem in sorted(semanas_presentes):
        rows_sem = [r for r in todas_las_filas if r.get("semana") == sem]
        tabla_por_semana[str(sem)] = generar_tabla(rows_sem)

    # Nombre de salida
    sufijo = args.salida or f"Dashboard_{fecha_max.strftime('%d_%m_%Y')}"
    if args.semana and len(args.semana) == 1:
        sufijo = f"Dashboard_Semana{args.semana[0]}_{fecha_max.strftime('%d_%m_%Y')}"

    ruta_xlsx = OUTPUT_DIR / f"{sufijo}.xlsx"
    ruta_json = OUTPUT_DIR / f"{sufijo}.json"

    # Generar Excel
    wb = openpyxl.Workbook()

    # Hoja 1: Acumulado
    ws_acum = wb.active
    ws_acum.title = "Acumulado"
    escribir_hoja_dashboard(ws_acum, titulo, tabla_acum)

    # Hoja 2: Por semana (si hay más de una o si se quiere el desglose)
    if semanas_presentes:
        ws_sem = wb.create_sheet("Por Semana")
        escribir_hoja_semanas(ws_sem, todas_las_filas, list(semanas_presentes))

    # Hoja 3: Detalle de registros
    ws_det = wb.create_sheet("Detalle Registros")
    escribir_hoja_detalle(ws_det, todas_las_filas)

    wb.save(ruta_xlsx)
    print(f"\n✓ Excel generado: {ruta_xlsx}")

    # Generar JSON
    exportar_json(tabla_acum, tabla_por_semana, todas_las_filas, ruta_json)
    print(f"✓ JSON  generado: {ruta_json}")

    # Generar datos para el dashboard web (GitHub Pages)
    ruta_web = exportar_web(tabla_acum, tabla_por_semana, todas_las_filas,
                            semanas_presentes, fecha_min, fecha_max)
    print(f"✓ Web   generado: {ruta_web}")

    # Resumen rápido en consola
    general = tabla_acum[0]
    print(f"\n── Resumen Acumulado ──────────────────────────")
    print(f"  Registros totales :  {general['modalidades_total']}")
    print(f"  Estudios modal.   :  {general['estudios']}")
    print(f"  Incumplimientos   :  {general['incumplimiento']}")
    ind = general["indicador"]
    print(f"  Indicador global  :  {ind:.2f}%" if ind else "  Indicador global  :  -")
    print(f"────────────────────────────────────────────────\n")


if __name__ == "__main__":
    main()

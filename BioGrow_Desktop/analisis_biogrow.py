"""
analisis_biogrow.py — Une la app de escritorio con el motor del dashboard (motor_biogrow.py).

- Convierte los registros de la base SQLite al mismo formato que lee el dashboard
  (una semilla por fila: lote, variedad, tratamiento, temperatura, raíz, tallo, estado, fecha)
  y los pasa por el MISMO lector que los Excel, así ambos se analizan igual.
- Ejecuta el mismo flujo de análisis que dashboardst.py: grupos, descartes, promedio por caja
  (o por semilla), descriptivos, ANOVA, Tukey, LSD, letras y cuartiles.
"""
from __future__ import annotations

import io
import os
import re
from typing import Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

import motor_biogrow as M

NOMBRE_BD = "Base de datos BioGrow.xlsx"
# Columnas del machote (plantilla para subir Excel). Es también el formato del Excel exportado.
ENCABEZADOS_BD = ["Lote Origen", "Variedad de Semilla", "No. Muestra", "Tipo de tratamiento", "Tratamiento",
                  "Temperatura (°C)", "Tipo de gas", "Flujo (L/min)", "Orientación", "Morfología",
                  "Cant. Hojas", "Raíz (mm)", "Tallo (mm)", "Estado", "Fecha de Registro"]
DESCRIPCION_COLUMNAS = [
    ("Lote Origen", "Nombre de la caja / lote. Todas las semillas de un lote llevan el mismo nombre.",
     "Lote 16-Mayo Blanco 45s"),
    ("Variedad de Semilla", "Maíz blanco, rosa, cacahuazintle, criollo o lenteja (el nombre técnico o científico "
     "entre paréntesis es opcional).", "Semilla de maíz blanco (PBB-DH-183)"),
    ("No. Muestra", "Número de la semilla dentro del lote (1, 2, 3…).", "1"),
    ("Tipo de tratamiento", "Plasma, Salino, Térmico o Solución (vacío = Plasma).", "Plasma"),
    ("Tratamiento", "Tiempo de tratamiento: Control, 30s, 45s, 60s, 2min…", "45s"),
    ("Temperatura (°C)", "Campo libre: un número en °C, un rango (35-40 °C), Ambiente u otro texto.", "35"),
    ("Tipo de gas", "Sin gas, Aire, Oxígeno, Helio, Argón o Nitrógeno.", "Argón"),
    ("Flujo (L/min)", "Litros por minuto (número). Vacío si no hubo flujo.", "2"),
    ("Orientación", "Aleatorio, Embrión o Endospermo.", "Embrión"),
    ("Morfología", "Plana o Bola.", "Plana"),
    ("Cant. Hojas", "Número entero (0 si no germinó).", "2"),
    ("Raíz (mm)", "Longitud de raíz en milímetros (0 si no germinó).", "84.5"),
    ("Tallo (mm)", "Longitud de tallo en milímetros (0 si no germinó).", "52.1"),
    ("Estado", "Germinó o No germinó.", "Germinó"),
    ("Fecha de Registro", "AAAA-MM-DD HH:MM:SS", "2026-10-28 09:00:00"),
]

# "Comparar por": etiqueta visible -> columnas del motor (None = automático según el experimento)
OPCIONES_COMPARAR = {
    "Automático (según el experimento)": None,
    "Variedad + Tratamiento": ["Variedad", "Tratamiento"],
    "Tratamiento": ["Tratamiento"],
    "Variedad": ["Variedad"],
    "Temperatura": ["Temperatura"],
    "Variedad + Temperatura": ["Variedad", "Temperatura"],
    "Variedad + Combinación de tratamiento": ["Variedad", "Combinación"],
    "Combinación de tratamiento": ["Combinación"],
    "Tipo de tratamiento (plasma, salino…)": ["Tipo de tratamiento"],
    "Variedad + Tipo de tratamiento": ["Variedad", "Tipo de tratamiento"],
    "Tipo de gas": ["Gas"],
    "Variedad + Tipo de gas": ["Variedad", "Gas"],
    "Flujo (L/min)": ["Flujo"],
    "Variedad + Flujo": ["Variedad", "Flujo"],
    "Orientación de la semilla": ["Orientación"],
    "Variedad + Orientación": ["Variedad", "Orientación"],
    "Morfología de la semilla": ["Morfología"],
    "Variedad + Morfología": ["Variedad", "Morfología"],
    "Parte expuesta (Embrión / Endospermo / Patrón)": ["Clase nombre"],
    "Fecha / hoja": ["Hoja"],
    "Archivo / origen": ["Archivo"],
}
ETIQUETAS_DIM = {
    "Temperatura": "Temperatura", "Tratamiento": "Tratamiento", "Clase nombre": "Parte expuesta",
    "Variedad": "Variedad", "Condición": "Condición", "Tirada": "Tirada", "Bloque": "Bloque",
    "Hoja": "Fecha / hoja", "Archivo": "Archivo", "Fecha etiqueta": "Fecha", "Corrida": "Caja",
    "Combinación": "Combinación de tratamiento", "Gas": "Tipo de gas", "Flujo": "Flujo",
    "Orientación": "Orientación", "Morfología": "Morfología", "Tipo de tratamiento": "Tipo de tratamiento",
}


# ---------------------------------------------------------------------------
# Lectura de datos
# ---------------------------------------------------------------------------

def excel_desde_registros(registros: Iterable[dict]) -> bytes:
    """Arma en memoria un Excel 'una semilla por fila' con los registros de la base."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Registros"
    ws.append(ENCABEZADOS_BD)
    for r in registros:
        raiz = r.get("longitud_raiz")
        tallo = r.get("longitud_tallo")
        raiz = float(raiz) if raiz is not None else None
        tallo = float(tallo) if tallo is not None else None
        germino = r.get("germino")
        if germino is None:
            germino = 1 if ((raiz or 0) > 0 or (tallo or 0) > 0) else 0
        fecha = r.get("fecha_registro")
        fecha = fecha.strftime("%Y-%m-%d %H:%M:%S") if hasattr(fecha, "strftime") else str(fecha or "")
        ws.append(_fila_machote(r, raiz, tallo, germino, fecha))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _fila_machote(r, raiz, tallo, germino, fecha) -> list:
    hojas = r.get("cantidad_hojas")
    flujo = r.get("flujo")
    return [
        str(r.get("nombre_lote") or "").strip(),
        str(r.get("tipo_semilla") or "").strip(),
        int(r.get("num_muestra") or 0),
        str(r.get("tipo_tratamiento") or "Plasma").strip(),
        str(r.get("tiempo_tratamiento") or "").strip(),
        str(r.get("temperatura") or "Ambiente").strip(),
        r.get("gas") or None,
        float(flujo) if flujo not in (None, "") else None,
        r.get("orientacion") or None,
        r.get("morfologia") or None,
        int(hojas) if hojas not in (None, "") else None,
        raiz, tallo,
        "Germinó" if int(germino) == 1 else "No germinó",
        fecha,
    ]


def excel_machote(registros: Iterable[dict] = (), titulo: Optional[str] = None) -> bytes:
    """Excel con el formato del machote: hoja 'Datos' (encabezados en la fila 1, una semilla por
    fila) y hoja 'Léeme' con la explicación de cada columna. Sin registros es la plantilla vacía."""
    from datetime import datetime
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    registros = list(registros)
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"
    ws.append(ENCABEZADOS_BD)
    verde, borde = PatternFill("solid", fgColor="059669"), Border(*(Side(style="thin", color="CBD5E1"),) * 4)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = verde
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borde
    ws.row_dimensions[1].height = 30
    for r in registros:
        raiz = float(r["longitud_raiz"]) if r.get("longitud_raiz") is not None else None
        tallo = float(r["longitud_tallo"]) if r.get("longitud_tallo") is not None else None
        germino = r.get("germino")
        if germino is None:
            germino = 1 if ((raiz or 0) > 0 or (tallo or 0) > 0) else 0
        fecha = r.get("fecha_registro")
        fecha = fecha.strftime("%Y-%m-%d %H:%M:%S") if hasattr(fecha, "strftime") else str(fecha or "")
        ws.append(_fila_machote(r, raiz, tallo, germino, fecha))
        fila = ws.max_row
        for c in ws[fila]:
            c.border = borde
            if fila % 2 == 1:
                c.fill = PatternFill("solid", fgColor="F8FAFC")
        ws.cell(fila, 13).font = Font(bold=True, color="059669" if int(germino) == 1 else "DC2626")
        for col in (11, 12):
            ws.cell(fila, col).number_format = "0.00"
    for i, ancho in enumerate([30, 28, 11, 12, 15, 13, 12, 13, 12, 11, 11, 11, 13, 20], start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"

    le = wb.create_sheet("Léeme")
    le["A1"] = titulo or "PLANTILLA BIOGROW — una fila por semilla"
    le["A1"].font = Font(bold=True, size=13, color="0F172A")
    if registros:
        total = len(registros)
        germ = sum(1 for r in registros if int(r.get("germino") or 0) == 1)
        le["A2"] = (f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Muestras: {total}  |  "
                    f"Germinadas: {germ} ({germ / total * 100:.1f} %)")
    else:
        le["A2"] = ("Llena la hoja 'Datos' desde la fila 2. No cambies los encabezados. Esta hoja ('Léeme') "
                    "la ignora el programa.")
    le["A2"].font = Font(italic=True, color="475569")
    le.append([])
    le.append(["Columna", "Qué poner", "Ejemplo"])
    for c in le[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = verde
    for fila in DESCRIPCION_COLUMNAS:
        le.append(list(fila))
    le.column_dimensions["A"].width = 22
    le.column_dimensions["B"].width = 80
    le.column_dimensions["C"].width = 26
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def leer_base(registros: List[dict]) -> Tuple[pd.DataFrame, List[str]]:
    if not registros:
        return pd.DataFrame(), []
    df, avisos = M.cargar_varios([(NOMBRE_BD, excel_desde_registros(registros))])
    if not df.empty:
        df["Origen"] = "Base de datos"
    return df, [a.replace(f"{NOMBRE_BD}: ", "Base de datos: ") for a in avisos]


def _hojas_tabla_larga(ruta) -> dict:
    """Cantidad de hojas por (lote, n.º de semilla) en tablas con una semilla por fila.
    El motor del dashboard no analiza hojas; solo se usa al importar a la base."""
    salida = {}
    try:
        libros = pd.read_excel(ruta, sheet_name=None)
    except Exception:
        return salida
    for hoja in libros.values():
        cols = {c: str(c).strip().lower() for c in hoja.columns}
        c_hojas = next((c for c, t in cols.items() if "hoja" in t and "cant" in t or t in ("hojas", "n.º hojas")), None)
        c_lote = next((c for c, t in cols.items() if re.search(M._COLUMNAS_LARGO["caja"], t)), None)
        c_no = next((c for c, t in cols.items() if re.search(M._COLUMNAS_LARGO["no"], t)), None)
        if not (c_hojas and c_lote and c_no):
            continue
        for lote, no, h in zip(hoja[c_lote], hoja[c_no], hoja[c_hojas]):
            if pd.notna(lote) and pd.notna(no) and pd.notna(h):
                try:
                    salida[(re.sub(r"\s+", " ", str(lote).strip()), int(no))] = int(h)
                except (TypeError, ValueError):
                    pass
    return salida


def leer_excels(rutas: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    if not rutas:
        return pd.DataFrame(), []
    df, avisos = M.cargar_varios(list(rutas))
    if not df.empty:
        df["Origen"] = "Excel"
        df["Hojas"] = np.nan
        for ruta in rutas:
            nombre = os.path.basename(str(ruta))
            m = (df["Archivo"] == nombre) & (df["Tipo de experimento"] == M.TIPO_LARGO)
            if m.any():
                mapa = _hojas_tabla_larga(ruta)
                if mapa:
                    df.loc[m, "Hojas"] = [mapa.get((b, int(n)), np.nan)
                                          for b, n in zip(df.loc[m, "Bloque"], df.loc[m, "No"])]
    return df, avisos


def combinar(df_bd: pd.DataFrame, df_xl: pd.DataFrame, modo: str,
             nombres_tecnicos: Optional[dict] = None) -> pd.DataFrame:
    """modo: 'Solo MySQL' (base de datos), 'Solo Excel' o 'MySQL + Excel (Combinados)'."""
    if modo.startswith("Solo MySQL"):
        marcos = [df_bd]
    elif modo.startswith("Solo Excel"):
        marcos = [df_xl]
    else:
        marcos = [df_bd, df_xl]
    marcos = [m for m in marcos if m is not None and not m.empty]
    if not marcos:
        return pd.DataFrame()
    df = pd.concat(marcos, ignore_index=True)
    df = aplicar_nombres_tecnicos(M.variedad_por_color(df), nombres_tecnicos or {})
    return M.agregar_combinacion(df)


# ---------------------------------------------------------------------------
# Nombres técnicos de las variedades (catálogo editable, se guarda en un .json)
# ---------------------------------------------------------------------------

# Variedades de la captura con su nombre técnico (híbrido) o científico, igual que en la versión de Hibrain.
TIPOS_SEMILLA = [
    "Semilla de maíz blanco (PBB-DH-183)",
    "Semilla de maíz rosa (H-392 A)",
    "Semilla de maíz cacahuazintle (Zea mays var. amylacea)",
    "Semilla de maíz criollo (Zea mays nativo)",
    "Semilla de lenteja (Lens culinaris)",
]

NOMBRES_TECNICOS_DEFECTO = {"Maíz blanco": "PBB-DH-183", "Maíz rosa": "H-392 A",
                            "Maíz cacahuazintle": "Zea mays var. amylacea", "Maíz criollo": "Zea mays nativo",
                            "Lenteja": "Lens culinaris"}
_VERSION_NOMBRES = 2


def separar_nombre(texto) -> Tuple[str, str]:
    """'Maíz blanco (PBB-DH-183)' -> ('Maíz blanco', 'PBB-DH-183')."""
    t = re.sub(r"\s+", " ", str(texto or "")).strip()
    m = re.match(r"^(.*?)\s*\(([^()]*)\)\s*$", t)
    return (m.group(1).strip(), m.group(2).strip()) if m else (t, "")


def clave_variedad(texto) -> str:
    """Clave para comparar variedades sin importar el nombre técnico ni el prefijo 'Semilla de':
    'Semilla de maíz rosa (H-392 A)' y 'Semilla de maíz rosa' dan la misma clave."""
    base = separar_nombre(texto)[0].lower()
    return re.sub(r"^semilla\s+de\s+", "", base).strip()


def leer_nombres_tecnicos(ruta: str) -> dict:
    """Catálogo {nombre común: nombre técnico/científico}. Los archivos guardados antes de que hubiera
    nombres científicos por defecto no borran esos valores con campos vacíos."""
    import json
    datos = dict(NOMBRES_TECNICOS_DEFECTO)
    try:
        with open(ruta, encoding="utf-8") as f:
            leido = json.load(f)
        nuevo = int(leido.get("_version", 1)) >= _VERSION_NOMBRES
        for k, v in leido.items():
            if str(k).startswith("_"):
                continue
            if str(v).strip() or nuevo or not datos.get(str(k)):
                datos[str(k)] = str(v)
    except (OSError, ValueError, AttributeError):
        pass
    return datos


def guardar_nombres_tecnicos(ruta: str, datos: dict) -> None:
    import json
    salida = {"_version": _VERSION_NOMBRES}
    salida.update({k: v for k, v in datos.items() if not str(k).startswith("_")})
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=2)


def aplicar_nombres_tecnicos(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """'Maíz blanco' -> 'Maíz blanco (PBB-DH-183)'. Un nombre que ya trae su nombre técnico entre
    paréntesis se agrupa con el mismo nombre sin él (los lotes viejos y los nuevos quedan juntos)."""
    if df.empty or "Variedad" not in df:
        return df
    mapa = {k: v for k, v in (mapa or {}).items() if not str(k).startswith("_")}
    d = df.copy()
    def nombre(v):
        base, propio = separar_nombre(v)
        t = str(mapa.get(base, "") or "").strip() or propio
        return f"{base} ({t})" if t else base
    d["Variedad"] = d["Variedad"].map(nombre)
    return d


def resumen_proyecto(df: pd.DataFrame, fuente: str = "") -> pd.DataFrame:
    """Tabla corta con la información del proyecto para los exportables."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Dato", "Valor"])
    def valores(col, quitar=("Sin dato", "No registrada", "Sin anotar", "Control", "nan", "Sin definir")):
        if col not in df:
            return "—"
        v = [str(x) for x in df[col].dropna().astype(str).unique() if str(x) not in quitar]
        def clave(x):
            m = re.search(r"\d+(?:\.\d+)?", x)
            return (float(m.group()) if m else 1e9, x)
        return ", ".join(sorted(v, key=clave)) or "—"
    por_lote = df.groupby("Corrida").size()
    f = pd.to_datetime(df.get("Fecha etiqueta"), errors="coerce", format="mixed").dropna()
    filas = [
        ("Fuente de datos", fuente or "—"),
        ("Variedades", valores("Variedad")),
        ("Lotes (cajas)", str(df["Corrida"].nunique())),
        ("Semillas", f"{len(df):,}"),
        ("Semillas por lote", f"{por_lote.min()}" if por_lote.min() == por_lote.max()
         else f"{por_lote.min()} a {por_lote.max()}"),
        ("Tipos de tratamiento", valores("Tipo de tratamiento")),
        ("Tiempos de tratamiento", ", ".join(M._etiqueta_tratamiento(x) for x in
                                             sorted(df["Tratamiento (s)"].dropna().unique())) or "—"),
        ("Temperaturas", valores("Temperatura")),
        ("Tipos de gas", valores("Gas")),
        ("Flujos", valores("Flujo")),
        ("Orientaciones", valores("Orientación")),
        ("Morfologías", valores("Morfología")),
        ("Combinaciones de tratamiento", str(df["Combinación"].nunique()) if "Combinación" in df else "—"),
        ("Fechas de registro", f"{f.min():%Y-%m-%d} a {f.max():%Y-%m-%d}" if len(f) else "—"),
    ]
    return pd.DataFrame(filas, columns=["Dato", "Valor"])


def unidad_por_defecto(df: pd.DataFrame) -> str:
    unidades = set(df["Unidad original"].dropna().astype(str)) if not df.empty else {"mm"}
    return "mm" if unidades <= {"mm"} else "cm"


def variedades(df: pd.DataFrame) -> List[str]:
    return sorted(df["Variedad"].dropna().astype(str).unique()) if not df.empty else []


def tratamientos(df: pd.DataFrame) -> List[str]:
    if df.empty:
        return []
    def orden(x):
        m = re.search(r"\d+(?:\.\d+)?", str(x))
        return (0 if "control" in str(x).lower() else 1, float(m.group()) if m else 0.0, str(x))
    return sorted(df["Tratamiento"].dropna().astype(str).unique(), key=orden)


def fechas(df: pd.DataFrame) -> List[str]:
    if df.empty or "Fecha etiqueta" not in df:
        return []
    f = pd.to_datetime(df["Fecha etiqueta"], errors="coerce", format="mixed").dropna()
    return sorted(f.dt.strftime("%Y-%m-%d").unique())


# ---------------------------------------------------------------------------
# Análisis (mismo flujo que main() de dashboardst.py)
# ---------------------------------------------------------------------------

def dims_automaticas(df: pd.DataFrame) -> List[str]:
    tipos = [t for t in M.CONFIG_TIPO if t in set(df["Tipo de experimento"])]
    if len(tipos) == 1:
        dims = list(M.CONFIG_TIPO[tipos[0]]["default"])
    else:
        dims = ["Variedad", "Tratamiento"]
    dims = [d for d in dims if d in df.columns]
    if df["Variedad"].nunique() > 1 and "Variedad" not in dims:
        dims = ["Variedad"] + dims
    return dims or ["Tratamiento"]


def filtrar(df: pd.DataFrame, variedades_sel=None, tratamiento=None, desde=None, hasta=None,
            gas=None) -> pd.DataFrame:
    d = df
    if variedades_sel:
        d = d[d["Variedad"].astype(str).isin(variedades_sel)]
    if gas:
        d = d[d["Gas"].astype(str) == gas]
    if tratamiento:
        d = d[d["Tratamiento"].astype(str) == tratamiento]
    if desde or hasta:
        f = pd.to_datetime(d["Fecha etiqueta"], errors="coerce", format="mixed")
        m = f.notna()
        if desde:
            m &= f >= pd.to_datetime(desde)
        if hasta:
            m &= f <= pd.to_datetime(hasta) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        d = d[m]
    return d.copy()


def analizar(df: pd.DataFrame, variable: str, comparar: str, unidad_dato: str, alfa: float,
             unidad_long: str, op: dict) -> dict:
    """Devuelve un dict con todos los resultados que muestran las pestañas 2 y 3.
    df: datos ya filtrados, con longitudes en cm (como los lee el motor)."""
    M.LONG = unidad_long
    salida = {"ok": False, "mensaje": "", "avisos": [], "variable": variable, "alfa": alfa}
    if df is None or df.empty:
        salida["mensaje"] = "No hay datos: registra lotes en la primera pestaña o carga archivos Excel."
        return salida

    d = df.copy()
    # Hojas 'Raíz' / 'Tallo' con un tratamiento por columna: cada fila mide una sola variable.
    es_col = d["Tipo de experimento"] == M.TIPO_COLUMNAS
    if es_col.any():
        if variable == "Germinó":
            d = d[~es_col]
            salida["avisos"].append("Los archivos con un tratamiento por columna solo traen semillas "
                                    "germinadas medidas; no entran al % de germinación.")
        else:
            d = d[~es_col | d[variable].notna()]
    if d.empty:
        salida["mensaje"] = f"Los datos elegidos no tienen {variable.lower()}."
        return salida
    if unidad_long == "mm":
        d[["Raíz", "Tallo"]] = (d[["Raíz", "Tallo"]] * 10).round(4)

    dims = OPCIONES_COMPARAR.get(comparar) or dims_automaticas(d)
    dims = [x for x in dims if x in d.columns] or ["Tratamiento"]
    agrupacion = " + ".join(ETIQUETAS_DIM.get(x, x) for x in dims)
    d["Grupo"] = M.construir_grupo(d, dims)

    # Unidad experimental: si ningún grupo tiene 2 o más cajas, el dato es la semilla.
    solo_col = bool((d["Tipo de experimento"] == M.TIPO_COLUMNAS).all())
    cajas_grupo = d.groupby("Grupo")["Corrida"].nunique()
    una_caja = bool(len(cajas_grupo)) and int(cajas_grupo.max()) < 2 and not solo_col
    forzar_semilla = solo_col or una_caja
    if forzar_semilla:
        por_caja = False
    else:
        por_caja = True if variable == "Germinó" else unidad_dato.startswith("Por caja")
    unidad_txt = "cajas" if por_caja else "semillas"
    u, u2, nombre_var = M._unidades(variable)

    op = dict(op)
    if variable in ("Raíz", "Tallo") and not bool((pd.to_numeric(d["Germinó"], errors="coerce") == 0).any()):
        op["no_germinadas"] = "excluir"
    df_semilla, registro = M.aplicar_descartes(d, variable, op)
    if por_caja:
        df_an = M.agregar_por_unidad(df_semilla, variable, M.COLUMNAS_CAJA + dims)
    else:
        df_an = df_semilla.dropna(subset=[variable]).copy()
        if variable == "Germinó":
            df_an[variable] = pd.to_numeric(df_an[variable], errors="coerce") * 100

    # Raíz y tallo con los descartes elegidos (sin atípicos ni rango), como en 'Más gráficas'
    ambos = d.copy()
    for v_ in ("Raíz", "Tallo"):
        ambos = M.aplicar_descartes(ambos, v_, {**op, "atipicos": False, "rango": None})[0]
    salida["ambos"] = ambos
    germ_col = M.columna_germinacion(df_semilla) if not df_semilla.empty else "Germinó"
    salida.update({
        "dims": dims, "agrupacion": agrupacion, "por_caja": por_caja, "una_caja": una_caja,
        "solo_columnas": solo_col, "forzar_semilla": forzar_semilla, "unidad_txt": unidad_txt,
        "u": u, "u2": u2, "nombre_var": nombre_var, "unidad_long": unidad_long,
        "df_filtrado": d, "df_semilla": df_semilla, "registro": registro, "df_an": df_an,
        "metricas": {
            "Semillas analizadas": len(df_semilla),
            "Cajas": int(df_semilla["Corrida"].nunique()) if not df_semilla.empty else 0,
            "Grupos comparados": int(df_an["Grupo"].nunique()) if not df_an.empty else 0,
            f"Datos ({unidad_txt})": len(df_an),
            "% germinación global": (float(df_semilla[germ_col].mean() * 100)
                                     if not df_semilla.empty and df_semilla["Germinó"].notna().any() else np.nan),
            "Registros descartados": len(registro),
        },
    })
    if df_an.empty:
        salida["mensaje"] = "No quedan datos con los filtros y descartes actuales."
        return salida

    res = M.pruebas_requeridas(df_an, variable, alfa)
    anova, tk, ls = res["anova"], res["tukey"], res["lsd"]
    desc = M.tabla_descriptivos(df_an, "Grupo", variable, alfa)
    datos_q = None if variable == "Germinó" else df_semilla.dropna(subset=[variable])
    factor_q = op.get("factor_iqr", 1.5)
    tc = M.tabla_cuartiles(datos_q if datos_q is not None else df_an, variable, factor_q)
    salida.update(res)
    salida.update({"ok": True, "desc": desc, "cuartiles": tc, "factor_iqr": factor_q,
                   "datos_cuartiles": datos_q if datos_q is not None else df_an,
                   "diag_cuartiles": M.diagnostico_cuartiles(tc, factor_q) if not tc.empty else None,
                   "ranking": M.tabla_ranking(desc, res["letras_tukey"], res["letras_lsd"], u,
                                              M.germinacion_por_grupo(df_semilla))})

    # Avisos (mismos criterios que el dashboard)
    av = salida["avisos"]
    pocos = df_an.groupby("Grupo").size()
    pocos = pocos[pocos < 3]
    if len(pocos):
        av.append(f"Grupos con menos de 3 {unidad_txt}: " + ", ".join(f"{g} (n = {n})" for g, n in pocos.items())
                  + ". Un grupo con 1 dato no entra al ANOVA.")
    if solo_col:
        av.append("Cada dato es una semilla (el archivo no tiene cajas).")
    elif una_caja:
        av.append("Cada dato es una semilla porque cada grupo tiene una sola caja (lote); con 2 o más "
                  "cajas por grupo el análisis sería más sólido."
                  + (" En germinación cada semilla vale 0 % o 100 %." if variable == "Germinó" else ""))
    elif not por_caja:
        av.append("Cada dato es una semilla: las pruebas tienden a encontrar diferencias que no son "
                  "reales. Para el reporte usa 'Por caja'.")
    if d["Variedad"].nunique() > 1 and "Variedad" not in dims:
        av.append("Estás juntando variedades distintas en un mismo grupo; agrega 'Variedad' en "
                  "'Comparar por' o filtra una variedad.")
    if anova and not anova["significativo"] and not ls.empty and (ls["Resultado"] == M.SIG).any():
        av.append("El ANOVA no es significativo: no reportes las diferencias que marque LSD.")

    if anova:
        salida["conclusion"] = M.texto_reporte(variable, u, agrupacion, unidad_txt, anova, tk, alfa, desc)
        salida["frases"] = M.pares_significativos(tk, u)
        salida["tabla_anova"] = M.tabla_anova_clara(anova, u2)
    else:
        salida["conclusion"] = ("Se necesitan al menos 2 grupos con 2 o más datos para el ANOVA. "
                                "Cambia 'Comparar por' o 'Cada dato es'.")
        salida["frases"] = []
        salida["tabla_anova"] = pd.DataFrame()
    salida["tabla_tukey"] = M.tabla_tukey_clara(tk, u)
    salida["tabla_lsd"] = M.tabla_lsd_clara(ls, u)

    dd = desc.copy()
    if not dd.empty:
        dd["GL (n − 1)"] = dd["n"] - 1
        dd = dd.rename(columns={"Media": f"Promedio ({u})", "Desviación estándar": f"Desviación estándar ({u})",
                                "Varianza": f"Varianza ({u2})", "SEM": f"Error estándar ({u})",
                                "CV %": "CV (%)", "Mínimo": f"Mínimo ({u})", "Mediana": f"Mediana ({u})",
                                "Máximo": f"Máximo ({u})", "n": f"n ({unidad_txt})"})
        cols = ["Grupo", f"n ({unidad_txt})", f"Promedio ({u})", f"Desviación estándar ({u})",
                f"Varianza ({u2})", "GL (n − 1)", f"Error estándar ({u})", "CV (%)",
                f"Mínimo ({u})", f"Mediana ({u})", f"Máximo ({u})"]
        if por_caja and "Semillas con dato" in df_an:
            sem_g = df_an.groupby("Grupo")["Semillas con dato"].sum().rename("Semillas usadas")
            dd = dd.merge(sem_g, left_on="Grupo", right_index=True, how="left")
            cols.insert(2, "Semillas usadas")
        dd = dd[[c for c in cols if c in dd.columns]]
    salida["tabla_descriptivos"] = dd
    return salida


# ---------------------------------------------------------------------------
# Caja por caja (misma lógica que la pestaña del dashboard)
# ---------------------------------------------------------------------------

def _apartado(f) -> str:
    if str(f.get("Clase", "")) == "S" or M._es_control(f.get("Grupo", "")):
        return "Patrón (sin tratamiento)"
    parte = str(f.get("Clase nombre", ""))
    return f"Corridas · {parte}" if parte in ("Embrión", "Endospermo") else "Corridas (tratadas)"


_ORDEN_AP = {"Patrón (sin tratamiento)": 0, "Corridas · Embrión": 1, "Corridas · Endospermo": 2,
             "Corridas (tratadas)": 3}


def tabla_cajas(res: dict) -> pd.DataFrame:
    crudo, analizado, variable = res["df_filtrado"], res["df_semilla"], res["variable"]
    es_germ = variable == "Germinó"
    u = res["u"]
    info = crudo.drop_duplicates("Corrida").set_index("Corrida")
    filas = []
    for caja, c in crudo.groupby("Corrida", sort=False):
        a = analizado[analizado["Corrida"] == caja]
        v = pd.to_numeric(a[variable], errors="coerce").dropna()
        germ = pd.to_numeric(c["Germinó"], errors="coerce")
        f = info.loc[caja]
        fila = {
            "Apartado": _apartado(f),
            "Caja": f"{f['Bloque']} · rep {f['Réplica']} · {f['Hoja']}",
            "Grupo": f["Grupo"], "Semillas": len(c), "Germinaron": int((germ == 1).sum()),
            "% germinación": germ.mean() * 100 if len(germ) else np.nan,
            "Semillas en el análisis": len(a) if es_germ else int(v.notna().sum()),
            "_id": caja,
        }
        if not es_germ:
            fila[f"Promedio ({u})"] = v.mean() if len(v) else np.nan
            fila[f"Desviación estándar ({u})"] = v.std(ddof=1) if len(v) > 1 else np.nan
            fila[f"Varianza ({res['u2']})"] = v.var(ddof=1) if len(v) > 1 else np.nan
        filas.append(fila)
    t = pd.DataFrame(filas)
    if t.empty:
        return t
    rep = t["Caja"].duplicated(keep=False)
    t.loc[rep, "Caja"] = t.loc[rep, "Caja"] + " · " + t.loc[rep, "_id"].map(info["Variedad"]).astype(str)
    t["_o"] = list(zip(t["Apartado"].map(_ORDEN_AP), t["Grupo"].astype(str), range(len(t))))
    return t.sort_values("_o").drop(columns="_o").reset_index(drop=True)


def semillas_de_cajas(res: dict, ids: List[str], etiquetas: dict) -> pd.DataFrame:
    crudo, analizado, variable, u = res["df_filtrado"], res["df_semilla"], res["variable"], res["u"]
    usados = analizado.set_index("ID semilla")[variable] if variable in analizado else pd.Series(dtype=float)
    partes = []
    for caja in ids:
        c = crudo[crudo["Corrida"] == caja].sort_values("No").copy()
        if variable == "Germinó":
            entra = c["ID semilla"].isin(usados.index)
        else:
            entra = c["ID semilla"].map(usados).notna()
        c["Entra al análisis"] = np.where(entra, "Sí", "No (descartada)")
        c.insert(0, "Caja", etiquetas.get(caja, caja))
        partes.append(c)
    if not partes:
        return pd.DataFrame()
    d = pd.concat(partes, ignore_index=True)
    d = d[["Caja", "No", "Germinó", "Raíz", "Tallo", "Entra al análisis"]]
    return d.rename(columns={"No": "Semilla n.º", "Germinó": "Germinó (1 = sí)",
                             "Raíz": f"Raíz ({res['unidad_long']})", "Tallo": f"Tallo ({res['unidad_long']})"})


def formatear_tabla(df: pd.DataFrame, dec: int = 3) -> Tuple[List[str], List[List[str]]]:
    """Columnas y filas como texto (3 decimales, 4 en p-valores, vacíos en blanco)."""
    if df is None or df.empty:
        return [], []
    cols = [c for c in df.columns if not str(c).startswith("_")]
    filas = []
    for _, r in df[cols].iterrows():
        fila = []
        for c in cols:
            v = r[c]
            if v is None or (isinstance(v, float) and np.isnan(v)):
                fila.append("")
            elif isinstance(v, (bool, np.bool_)):
                fila.append("Sí" if v else "No")
            elif isinstance(v, (int, np.integer)):
                fila.append(str(int(v)))
            elif isinstance(v, (float, np.floating)):
                fila.append(f"{v:.4f}" if M._es_columna_p(c) else f"{v:.{dec}f}")
            else:
                fila.append(str(v))
        filas.append(fila)
    return [str(c) for c in cols], filas


# ---------------------------------------------------------------------------
# Importar Excel a la base de datos (pestaña 1)
# ---------------------------------------------------------------------------

_BLANCO, _ROSA, _CACAHUAZINTLE, _CRIOLLO, _LENTEJA = TIPOS_SEMILLA
_VARIEDAD_BD = [("blanc", _BLANCO), ("pbb", _BLANCO), ("rosa", _ROSA), ("h-392", _ROSA),
                ("cacahuazintle", _CACAHUAZINTLE), ("amylacea", _CACAHUAZINTLE), ("criollo", _CRIOLLO),
                ("lenteja", _LENTEJA), ("lens culinaris", _LENTEJA)]
_COLOR_BD = {"B": _BLANCO, "R": _ROSA, "C": _CACAHUAZINTLE, "CR": _CRIOLLO, "L": _LENTEJA}


def _variedad_bd(color, variedad) -> str:
    if str(color).upper() in _COLOR_BD:
        return _COLOR_BD[str(color).upper()]
    t = str(variedad or "").lower()
    for clave, nombre in _VARIEDAD_BD:
        if clave in t:
            return nombre
    v = str(variedad or "").strip()
    return v if v and v != "?" else "Sin especificar"


def _tiempo_bd(f) -> str:
    if str(f.get("Clase", "")) == "S" or f.get("Tratamiento (s)") == 0:
        return "Control"
    seg = f.get("Tratamiento (s)")
    if seg is not None and not pd.isna(seg):
        return f"{float(seg):g}s"
    return "Sin definir"


def _temperatura_bd(f) -> str:
    t = f.get("Temperatura registrada")
    if t is None or pd.isna(t):
        t = f.get("Temperatura")
    t = str(t or "").strip()
    return "Ambiente" if t in ("", "Control", "No registrada", "Sin anotar", "nan") else t


def _fecha_bd(f, hoy: str) -> str:
    fecha = pd.to_datetime(f.get("Fecha etiqueta"), errors="coerce", format="mixed")
    return fecha.strftime("%Y-%m-%d %H:%M:%S") if not pd.isna(fecha) else hoy


def preparar_importacion(df_xl: pd.DataFrame, existentes: set) -> Tuple[list, dict]:
    """Convierte las cajas de los Excel (formato del motor, en cm) en lotes para la base (en mm).
    Devuelve (lotes, resumen). No importa cajas vacías, lotes ya registrados ni hojas con un
    tratamiento por columna (ahí raíz y tallo no son de las mismas semillas)."""
    from datetime import datetime
    resumen = {"lotes": 0, "semillas": 0, "repetidos": [], "vacias": 0, "sin_pareja": 0,
               "fecha_hoy": 0, "archivos": set()}
    if df_xl is None or df_xl.empty:
        return [], resumen
    existentes = {(n, clave_variedad(v)) for n, v in existentes}  # lotes viejos sin nombre científico
    d = M.variedad_por_color(df_xl)
    es_col = d["Tipo de experimento"] == M.TIPO_COLUMNAS
    resumen["sin_pareja"] = int(d.loc[es_col, "Corrida"].nunique())
    d = d[~es_col]
    vacias = M.cajas_vacias(d) if not d.empty else set()
    resumen["vacias"] = len(vacias)
    d = d[~d["Corrida"].isin(vacias)]
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lotes = []
    for caja, c in d.groupby("Corrida", sort=False):
        f = c.iloc[0].to_dict()
        if f["Tipo de experimento"] == M.TIPO_LARGO:
            nombre = str(f["Bloque"])
        else:
            nombre = f"{f['Bloque']} · rep {f['Réplica']} · {f['Hoja']}"
        variedad = _variedad_bd(f.get("Color"), f.get("Variedad"))
        if (nombre, clave_variedad(variedad)) in existentes:
            resumen["repetidos"].append(nombre)
            continue
        fecha = _fecha_bd(f, hoy)
        resumen["fecha_hoy"] += int(fecha == hoy)
        mediciones = []
        for _, s in c.sort_values("No").iterrows():
            raiz = 0.0 if pd.isna(s["Raíz"]) else round(float(s["Raíz"]) * 10, 2)
            tallo = 0.0 if pd.isna(s["Tallo"]) else round(float(s["Tallo"]) * 10, 2)
            germ = s["Germinó"]
            germ = (1 if (raiz > 0 or tallo > 0) else 0) if pd.isna(germ) else int(germ)
            hojas = s.get("Hojas")
            mediciones.append({"num_muestra": int(s["No"]), "longitud_raiz": raiz, "longitud_tallo": tallo,
                               "cantidad_hojas": None if hojas is None or pd.isna(hojas) else int(hojas),
                               "germino": germ})
        # números de semilla únicos dentro del lote (la base no admite repetidos)
        vistos = set()
        for k, med in enumerate(mediciones, start=1):
            if med["num_muestra"] in vistos:
                med["num_muestra"] = max(vistos) + 1
            vistos.add(med["num_muestra"])
        def _fac(col):
            v = f.get(col)
            return None if v is None or (isinstance(v, float) and pd.isna(v)) or str(v) in ("Sin dato", "nan") else v
        flujo = f.get("Flujo (L/min)")
        lotes.append({"nombre_lote": nombre, "tipo_semilla": variedad, "fecha_registro": fecha,
                      "gas": _fac("Gas"), "orientacion": _fac("Orientación"), "morfologia": _fac("Morfología"),
                      "tipo_tratamiento": _fac("Tipo de tratamiento") or "Plasma",
                      "flujo": None if flujo is None or pd.isna(flujo) else float(flujo),
                      "cantidad_semillas": len(mediciones),
                      "tiempo_tratamiento": _tiempo_bd(f), "temperatura": _temperatura_bd(f),
                      "mediciones": mediciones, "archivo": f.get("Archivo")})
        existentes.add((nombre, clave_variedad(variedad)))
        resumen["lotes"] += 1
        resumen["semillas"] += len(mediciones)
        resumen["archivos"].add(f.get("Archivo"))
    return lotes, resumen

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
import re
from typing import Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

import motor_biogrow as M

NOMBRE_BD = "Base de datos BioGrow.xlsx"
ENCABEZADOS_BD = ["Lote Origen", "Variedad de Semilla", "No. Muestra", "Tratamiento",
                  "Temperatura (°C)", "Cant. Hojas", "Raíz (mm)", "Tallo (mm)", "Estado",
                  "Fecha de Registro"]

# "Comparar por": etiqueta visible -> columnas del motor (None = automático según el experimento)
OPCIONES_COMPARAR = {
    "Automático (según el experimento)": None,
    "Variedad + Tratamiento": ["Variedad", "Tratamiento"],
    "Tratamiento": ["Tratamiento"],
    "Variedad": ["Variedad"],
    "Temperatura": ["Temperatura"],
    "Variedad + Temperatura": ["Variedad", "Temperatura"],
    "Parte expuesta (Embrión / Endospermo / Patrón)": ["Clase nombre"],
    "Fecha / hoja": ["Hoja"],
    "Archivo / origen": ["Archivo"],
}
ETIQUETAS_DIM = {
    "Temperatura": "Temperatura", "Tratamiento": "Tratamiento", "Clase nombre": "Parte expuesta",
    "Variedad": "Variedad", "Condición": "Condición", "Tirada": "Tirada", "Bloque": "Bloque",
    "Hoja": "Fecha / hoja", "Archivo": "Archivo", "Fecha etiqueta": "Fecha", "Corrida": "Caja",
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
        hojas = r.get("cantidad_hojas")
        ws.append([
            str(r.get("nombre_lote") or "").strip(),
            str(r.get("tipo_semilla") or "").strip(),
            int(r.get("num_muestra") or 0),
            str(r.get("tiempo_tratamiento") or "").strip(),
            str(r.get("temperatura") or "Ambiente").strip(),
            int(hojas) if hojas not in (None, "") else None,
            raiz, tallo,
            "Germinó" if int(germino) == 1 else "No germinó",
            fecha,
        ])
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


def leer_excels(rutas: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    if not rutas:
        return pd.DataFrame(), []
    df, avisos = M.cargar_varios(list(rutas))
    if not df.empty:
        df["Origen"] = "Excel"
    return df, avisos


def combinar(df_bd: pd.DataFrame, df_xl: pd.DataFrame, modo: str) -> pd.DataFrame:
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
    return M.variedad_por_color(df)


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


def filtrar(df: pd.DataFrame, variedades_sel=None, tratamiento=None, desde=None, hasta=None) -> pd.DataFrame:
    d = df
    if variedades_sel:
        d = d[d["Variedad"].astype(str).isin(variedades_sel)]
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

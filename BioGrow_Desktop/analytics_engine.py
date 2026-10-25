"""
analytics_engine.py — Motor Estadístico y Lector Híbrido (MySQL + Excel) de BioGrow
Integra la lógica completa de dashboardst.py con la estructura MySQL de BioGrow Desktop.
"""

from __future__ import annotations

import io
import math
import re
import zipfile
from itertools import combinations
from pathlib import Path
from typing import List, Tuple
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd
from scipy import stats


# ===========================================================================
# 1. CONSTANTES Y LECTOR DE EXCEL (PORTADO DE dashboardst.py)
# ===========================================================================

PATRON_TITULO = re.compile(r"^\s*T\s*[-–]\s*(0*)(\d+)\s*([MNmn]?)\s*$")
PATRON_FECHA_HOJA = re.compile(
    r"(\d{1,2})\s*(?:de\s+)?("
    r"enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
    r"septiembre|octubre|noviembre|diciembre)",
    re.IGNORECASE,
)
MES_A_NUMERO = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}
TRATAMIENTOS_POR_FECHA = {
    (28, 2): {"ambos": 120},
    (7, 3): {"par": 60, "impar": 30},
    (14, 3): {"par": 60, "impar": 30},
    (21, 3): {"par": 75, "impar": 45},
    (25, 4): {"par": 75, "impar": 45},
    (2, 5): {"par": 45, "impar": 30},
    (16, 5): {"par": 45, "impar": 30},
}
NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

PATRON_LEYENDA = re.compile(r"^\s*(\d+)\s*[-–]\s*(\d+(?:[.,]\d+)?)\s*(s|seg|min)\b", re.IGNORECASE)
PATRON_TIEMPO = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*(s|seg|segundos|min|minutos)\s*$", re.IGNORECASE)
PATRON_MUESTRA = re.compile(
    r"^\s*M\s*0*(\d+)\s*[-–]\s*T?\s*(amb(?:iente)?|\d+(?:[.,]\d+)?)\s*(?:°\s*C)?\s*$",
    re.IGNORECASE,
)
PATRON_FECHA_ARCHIVO = re.compile(r"(\d{1,2})[-_ ]([A-Za-zé]+)[-_ ](\d{4})")
PATRON_T_NOTA = re.compile(r"\bt\s*=\s*(\d+(?:[.,]\d+)?)\s*s\b", re.IGNORECASE)
PATRON_T_ARCHIVO = re.compile(r"(?:^|[_\-\s]|t)(\d+(?:[.,]\d+)?)\s*s(?:\b|_|\.|$)", re.IGNORECASE)

TIPO_TIEMPO = "Tiempo de tratamiento por fecha (Blanco / Rosa)"
TIPO_TIEMPO30 = "Tiempo 30 s por variedad (Cacahuazintle / Criollo)"
TIPO_TEMPERATURA = "Temperatura del tratamiento (Criollo)"
TIPO_MYSQL = "Registros en Base de Datos (MySQL)"

ENCABEZADO_RAIZ = "raíz [mm]"
TEXTO_SIN_GERMINAR = "sin germinar"
SEGUNDOS_TIEMPO30 = 30
_OFFSETS = {"No": 0, "Raíz": 1, "Tallo": 2, "Germinó": 3}

SIG = "Diferencia significativa"
NO_SIG = "Sin diferencia significativa"


def _col_a_indice(letras: str) -> int:
    indice = 0
    for caracter in letras:
        indice = indice * 26 + (ord(caracter.upper()) - ord("A") + 1)
    return indice


def _partir_celda(referencia: str):
    coincidencia = re.match(r"([A-Z]+)(\d+)", referencia)
    letras, numero = coincidencia.groups()
    return letras, int(numero)


def _leer_cadenas_compartidas(zf: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    raiz = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter("{%s}t" % NS_MAIN["m"])) for si in raiz.findall("m:si", NS_MAIN)]


def _mapa_hojas(zf: zipfile.ZipFile):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_a_destino = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS_REL)}
    hojas = []
    for hoja in workbook.find("m:sheets", NS_MAIN).findall("m:sheet", NS_MAIN):
        nombre = hoja.get("name")
        rid = hoja.get("{%s}id" % NS_REL["r"])
        destino = rid_a_destino.get(rid, "")
        destino = "xl/" + destino if not destino.startswith("/") else destino.lstrip("/")
        hojas.append((nombre, destino))
    return hojas


def _leer_celdas(zf: zipfile.ZipFile, ruta_hoja: str, cadenas):
    raiz = ET.fromstring(zf.read(ruta_hoja))
    datos = raiz.find("m:sheetData", NS_MAIN)
    celdas = {}
    if datos is None:
        return celdas
    for fila in datos.findall("m:row", NS_MAIN):
        for celda in fila.findall("m:c", NS_MAIN):
            referencia = celda.get("r")
            if not referencia:
                continue
            letras, numero_fila = _partir_celda(referencia)
            col = _col_a_indice(letras)
            tipo = celda.get("t")
            v = celda.find("m:v", NS_MAIN)
            if v is None or v.text is None:
                is_ = celda.find("m:is", NS_MAIN)
                if is_ is not None:
                    celdas[(numero_fila, col)] = "".join(t.text or "" for t in is_.iter("{%s}t" % NS_MAIN["m"]))
                continue
            texto = v.text
            if tipo == "s":
                celdas[(numero_fila, col)] = cadenas[int(texto)]
            elif tipo == "b":
                celdas[(numero_fila, col)] = int(texto)
            else:
                try:
                    numero = float(texto)
                    celdas[(numero_fila, col)] = int(numero) if numero.is_integer() else numero
                except ValueError:
                    celdas[(numero_fila, col)] = texto
    return celdas


def _fecha_de_hoja(nombre_hoja: str):
    coincidencia = PATRON_FECHA_HOJA.search(nombre_hoja.lower())
    if not coincidencia:
        return None, None, None
    dia = int(coincidencia.group(1))
    mes_texto = coincidencia.group(2).lower()
    return dia, MES_A_NUMERO[mes_texto], f"{dia:02d} {mes_texto}"


def _color_de_hoja(nombre_hoja: str) -> str:
    texto = nombre_hoja.lower()
    if "cacahuazintle" in texto:
        return "C"
    if "criollo" in texto:
        return "CR"
    if "rosa" in texto or re.search(r"datos-?r\b", texto) or "-r " in texto or "(r" in texto:
        return "R"
    if "blanc" in texto or re.search(r"datos-?b\b", texto):
        return "B"
    return "?"


def _a_segundos(numero_txt, unidad):
    valor = float(str(numero_txt).replace(",", "."))
    return valor * 60 if unidad.lower().startswith("min") else valor


def _leyenda_paridad(celdas):
    regla = {}
    for (fila, _col), valor in celdas.items():
        if fila > 2 or not isinstance(valor, str):
            continue
        m = PATRON_LEYENDA.match(valor.strip())
        if m:
            clave = "par" if int(m.group(1)) % 2 == 0 else "impar"
            regla[clave] = _a_segundos(m.group(2), m.group(3))
    return regla if {"par", "impar"} <= set(regla) else None


def _anotacion_bloque(celdas, fila_titulo, col_bloque):
    for col in (col_bloque, col_bloque + 1):
        valor = celdas.get((fila_titulo - 1, col))
        if not isinstance(valor, str) or not valor.strip():
            continue
        texto = valor.strip()
        if PATRON_TITULO.match(texto) or PATRON_LEYENDA.match(texto):
            continue
        m = PATRON_TIEMPO.match(texto.replace(" ", ""))
        if m:
            return _a_segundos(m.group(1), m.group(2)), None
        return None, texto
    return None, None


def _segundos_tratamiento(dia, mes, tirada, regla_hoja=None):
    regla = regla_hoja or TRATAMIENTOS_POR_FECHA.get((dia, mes))
    if regla is None:
        return None
    if "ambos" in regla:
        return regla["ambos"]
    return regla["par"] if tirada % 2 == 0 else regla["impar"]


def _datos_clase(letra: str):
    letra = (letra or "").upper()
    if letra == "M":
        return "M", "Embrión"
    if letra == "N":
        return "N", "Endospermo"
    return "S", "Patrón"


def _extraer_bloques_de_hoja(nombre_hoja, celdas, advertencias):
    dia, mes, etiqueta_fecha = _fecha_de_hoja(nombre_hoja)
    color = _color_de_hoja(nombre_hoja)
    regla_hoja = _leyenda_paridad(celdas)
    titulos = []
    for (fila, col), valor in celdas.items():
        if isinstance(valor, str):
            coincidencia = PATRON_TITULO.match(valor.strip())
            if coincidencia:
                titulos.append((
                    fila, col,
                    int(coincidencia.group(2)),
                    coincidencia.group(3),
                    coincidencia.group(1),
                    valor.strip()
                ))
    titulos.sort()
    hoja_con_mn = any(t[3] for t in titulos)
    registros = []
    contador_replica = {}

    for fila_titulo, col_bloque, tirada, letra, cero, titulo_texto in titulos:
        if letra:
            codigo_clase, nombre_clase = _datos_clase(letra)
        elif cero or hoja_con_mn:
            codigo_clase, nombre_clase = "S", "Patrón"
        else:
            codigo_clase, nombre_clase = "N", "Endospermo"
        seg_anot, condicion = _anotacion_bloque(celdas, fila_titulo, col_bloque)

        if codigo_clase == "S":
            segundos = 0
        elif seg_anot is not None:
            segundos = seg_anot
            condicion = condicion or f"Tiempo anotado ({seg_anot:g} s)"
        else:
            segundos = _segundos_tratamiento(dia, mes, tirada, regla_hoja)
        condicion = condicion or "Estándar"

        clave_replica = (tirada, codigo_clase)
        contador_replica[clave_replica] = contador_replica.get(clave_replica, 0) + 1
        numero_replica = contador_replica[clave_replica]

        fila_datos = fila_titulo + 2
        while True:
            celda_no = celdas.get((fila_datos, col_bloque + _OFFSETS["No"]))
            if not isinstance(celda_no, (int, float)):
                break
            raiz = celdas.get((fila_datos, col_bloque + _OFFSETS["Raíz"]))
            tallo = celdas.get((fila_datos, col_bloque + _OFFSETS["Tallo"]))
            germino = celdas.get((fila_datos, col_bloque + _OFFSETS["Germinó"]))

            registros.append({
                "Hoja": nombre_hoja,
                "Color": color,
                "Tirada": tirada,
                "Clase": codigo_clase,
                "Clase nombre": nombre_clase,
                "Bloque": titulo_texto,
                "Réplica": numero_replica,
                "No": int(celda_no),
                "Raíz": float(raiz) if isinstance(raiz, (int, float)) else None,
                "Tallo": float(tallo) if isinstance(tallo, (int, float)) else None,
                "Hojas": np.nan,
                "Germinó": int(germino) if isinstance(germino, (int, float)) else None,
                "Tratamiento (s)": segundos,
                "Condición": condicion,
                "Temperatura": "No registrada",
                "Unidad original": "mm",
            })
            fila_datos += 1
    return registros


def _es_formato_temperatura(celdas):
    return any(isinstance(v, str) and PATRON_MUESTRA.match(v.strip()) for v in celdas.values())


def _es_formato_tiempo(celdas):
    return any(isinstance(v, str) and v.strip().lower() == ENCABEZADO_RAIZ for v in celdas.values())


def _unidad_longitud(celdas):
    for (fila, _c), valor in celdas.items():
        if fila <= 3 and isinstance(valor, str) and re.search(r"\bcm\b", valor, re.IGNORECASE):
            return 10.0, "cm"
    return 1.0, "mm"


def _tiempo_temperatura(celdas, nombre_archivo):
    en_hoja = None
    for (fila, _c), valor in celdas.items():
        if fila <= 2 and isinstance(valor, str):
            m = PATRON_T_NOTA.search(valor)
            if m:
                en_hoja = float(m.group(1).replace(",", "."))
    m = PATRON_T_ARCHIVO.search(Path(nombre_archivo or "").stem)
    en_nombre = float(m.group(1).replace(",", ".")) if m else None
    if en_hoja is not None:
        return en_hoja
    return en_nombre


def _etiqueta_desde_archivo(nombre_archivo, respaldo):
    m = PATRON_FECHA_ARCHIVO.search(Path(nombre_archivo or "").stem)
    if not m:
        return respaldo
    variedad = "Criollo" if "criollo" in (nombre_archivo or "").lower() else respaldo
    return f"{variedad} {int(m.group(1)):02d}-{m.group(2).capitalize()}-{m.group(3)}"


def _extraer_hoja_temperatura(nombre_hoja, celdas, advertencias, nombre_archivo=""):
    factor, unidad = _unidad_longitud(celdas)
    segundos = _tiempo_temperatura(celdas, nombre_archivo)
    color = _color_de_hoja(nombre_hoja)
    etiqueta_hoja = _etiqueta_desde_archivo(nombre_archivo, nombre_hoja)

    titulos = []
    for (fila, col), valor in celdas.items():
        if not isinstance(valor, str):
            continue
        texto = valor.strip()
        m = PATRON_MUESTRA.match(texto)
        if m:
            titulos.append((fila, col, "M", int(m.group(1)), m.group(2), texto))
            continue
        m = PATRON_TITULO.match(texto)
        if m:
            extra = m.group(3) if (m.group(3) or not m.group(1)) else ""
            titulos.append((fila, col, "T", int(m.group(2)), extra, texto))
    titulos.sort()

    registros, contador = [], {}
    for fila_t, col_t, tipo, numero, extra, texto in titulos:
        filas_bloque, fila = [], fila_t + 2
        while isinstance(celdas.get((fila, col_t)), (int, float)):
            raiz = celdas.get((fila, col_t + 1))
            tallo = celdas.get((fila, col_t + 2))
            germ = celdas.get((fila, col_t + 3))
            filas_bloque.append((
                int(celdas[(fila, col_t)]),
                raiz * factor if isinstance(raiz, (int, float)) else None,
                tallo * factor if isinstance(tallo, (int, float)) else None,
                int(germ) if isinstance(germ, (int, float)) else None,
            ))
            fila += 1
        if not filas_bloque or all(not (r or t or g) for _, r, t, g in filas_bloque):
            continue

        if tipo == "M":
            temp_txt = extra.lower()
            temp_lbl = "Ambiente" if temp_txt.startswith("amb") else f"{float(temp_txt.replace(',', '.')):g} °C"
            codigo, nombre_clase = "T", "Tratada"
            seg = segundos
            tirada = numero
        else:
            codigo, nombre_clase = _datos_clase(extra)
            temp_lbl = "Control" if codigo == "S" else "Sin anotar"
            seg = 0 if codigo == "S" else segundos
            tirada = numero if codigo != "S" else 0

        clave = (tipo, texto.upper().replace(" ", ""))
        contador[clave] = contador.get(clave, 0) + 1
        for no, raiz, tallo, germ in filas_bloque:
            registros.append({
                "Hoja": etiqueta_hoja,
                "Color": color,
                "Tirada": tirada,
                "Clase": codigo,
                "Clase nombre": nombre_clase,
                "Bloque": re.sub(r"\s+", " ", texto),
                "Réplica": contador[clave],
                "No": no,
                "Raíz": raiz,
                "Tallo": tallo,
                "Hojas": np.nan,
                "Germinó": germ,
                "Tratamiento (s)": seg,
                "Condición": "Estándar",
                "Temperatura": temp_lbl,
                "Unidad original": unidad,
            })
    return registros


def _clase_desde_texto(texto):
    t = (texto or "").strip().lower().rstrip(":")
    if t.startswith("embri"):
        return "M", "Embrión"
    if t.startswith("endosperm"):
        return "N", "Endospermo"
    return "S", "Patrón"


def _leer_bloques_tiempo30(celdas, col_datos, variedad, etiqueta_hoja):
    registros = []
    encabezados = sorted(
        fila for (fila, col), valor in celdas.items()
        if col == col_datos and isinstance(valor, str) and valor.strip().lower() == ENCABEZADO_RAIZ
    )
    contador_replica = {}
    for fila_enc in encabezados:
        etiqueta_izq = str(celdas.get((fila_enc, col_datos - 1)) or "").strip()
        m = re.search(r"experimento\s*(\d+)", etiqueta_izq, re.IGNORECASE)
        if m:
            tirada = int(m.group(1))
            texto_clase = celdas.get((fila_enc - 1, col_datos)) or ""
            codigo, nombre_clase = _clase_desde_texto(str(texto_clase))
            nombre_bloque = f"Exp{tirada}-{codigo}"
        else:
            tirada = 0
            codigo, nombre_clase = "S", "Patrón"
            nombre_bloque = "Control"
        segundos_bloque = 0 if codigo == "S" else SEGUNDOS_TIEMPO30
        clave = (tirada, codigo)
        contador_replica[clave] = contador_replica.get(clave, 0) + 1
        replica = contador_replica[clave]

        fila = fila_enc + 1
        valores = []
        sin_germinar = 0
        limite = max(f for (f, _c) in celdas) + 2
        while fila < limite:
            celda = celdas.get((fila, col_datos))
            if isinstance(celda, str):
                if TEXTO_SIN_GERMINAR in celda.strip().lower():
                    conteo = celdas.get((fila, col_datos + 1))
                    sin_germinar = int(conteo) if isinstance(conteo, (int, float)) else 0
                break
            if isinstance(celda, (int, float)):
                tallo = celdas.get((fila, col_datos + 1))
                valores.append((float(celda), float(tallo) if isinstance(tallo, (int, float)) else None))
            fila += 1

        if not valores and sin_germinar == 0:
            continue
        numero = 0
        for raiz, tallo in valores:
            numero += 1
            registros.append({
                "Hoja": etiqueta_hoja, "Color": variedad, "Tirada": tirada,
                "Clase": codigo, "Clase nombre": nombre_clase, "Bloque": nombre_bloque,
                "Réplica": replica, "No": numero, "Raíz": raiz, "Tallo": tallo,
                "Hojas": np.nan, "Germinó": 1, "Tratamiento (s)": segundos_bloque,
                "Condición": "Estándar", "Temperatura": "No registrada", "Unidad original": "mm",
            })
        for _ in range(sin_germinar):
            numero += 1
            registros.append({
                "Hoja": etiqueta_hoja, "Color": variedad, "Tirada": tirada,
                "Clase": codigo, "Clase nombre": nombre_clase, "Bloque": nombre_bloque,
                "Réplica": replica, "No": numero, "Raíz": 0.0, "Tallo": 0.0,
                "Hojas": 0.0, "Germinó": 0, "Tratamiento (s)": segundos_bloque,
                "Condición": "Estándar", "Temperatura": "No registrada", "Unidad original": "mm",
            })
    return registros


def _extraer_hoja_tiempo(nombre_hoja, celdas):
    etiqueta = "Tiempo 30 s"
    for (_f, _c), valor in celdas.items():
        if isinstance(valor, str) and re.match(r"^\s*t\s*=", valor.strip(), re.IGNORECASE):
            etiqueta = valor.strip().replace("t =", "Tiempo").strip()
            break
    variedades = [
        (col, valor.strip())
        for (fila, col), valor in sorted(celdas.items())
        if fila <= 3 and isinstance(valor, str) and valor.strip().isupper() and not valor.strip().lower().startswith("t =")
    ]
    registros = []
    for col_variedad, nombre_variedad in variedades:
        encontrados = _leer_bloques_tiempo30(celdas, col_variedad + 1, nombre_variedad, etiqueta)
        if not encontrados:
            encontrados = _leer_bloques_tiempo30(celdas, col_variedad, nombre_variedad, etiqueta)
        registros.extend(encontrados)
    return registros


def _etiqueta_tratamiento(seg):
    if pd.isna(seg):
        return "Sin definir"
    if seg == 0:
        return "Control (0 s)"
    return f"{int(seg)} s" if float(seg).is_integer() else f"{seg:g} s"


def _leer_excel_estandar(ruta_archivo: str) -> pd.DataFrame:
    """Si el usuario carga un Excel exportado por BioGrow Desktop, lo lee directamente."""
    try:
        df_raw = pd.read_excel(ruta_archivo, header=3)
        cols = {str(c).strip().lower(): c for c in df_raw.columns}
        if "lote origen" in cols and "raíz (mm)" in cols:
            registros = []
            for _, r in df_raw.iterrows():
                lote = str(r[cols["lote origen"]])
                var = str(r[cols.get("variedad de semilla", "lote origen")])
                no = int(r[cols.get("no. muestra", "lote origen")]) if pd.notna(r.get(cols.get("no. muestra"))) else 1
                trat = str(r[cols.get("tratamiento", "lote origen")])
                hojas = float(r[cols["cant. hojas"]]) if "cant. hojas" in cols and pd.notna(r[cols["cant. hojas"]]) else np.nan
                raiz = float(r[cols["raíz (mm)"]]) if pd.notna(r[cols["raíz (mm)"]]) else 0.0
                tallo = float(r[cols["tallo (mm)"]]) if pd.notna(r[cols["tallo (mm)"]]) else 0.0
                germ = 1 if (raiz > 0 or tallo > 0) else 0
                es_ctrl = "control" in trat.lower() or "patrón" in trat.lower() or trat.strip() == "0"
                registros.append({
                    "Archivo": Path(ruta_archivo).name,
                    "Origen": "Excel",
                    "Hoja": lote,
                    "Color": var,
                    "Variedad": var,
                    "Tirada": 0 if es_ctrl else 1,
                    "Clase": "S" if es_ctrl else "T",
                    "Clase nombre": "Patrón" if es_ctrl else "Tratada",
                    "Bloque": trat,
                    "Réplica": 1,
                    "No": no,
                    "Raíz": raiz,
                    "Tallo": tallo,
                    "Hojas": hojas,
                    "Germinó": germ,
                    "Tratamiento": "Control (0 s)" if es_ctrl else trat,
                    "Tratamiento (s)": 0 if es_ctrl else np.nan,
                    "Condición": "Estándar",
                    "Temperatura": "No registrada",
                    "Unidad original": "mm",
                    "Tipo de experimento": TIPO_MYSQL,
                })
            return pd.DataFrame(registros)
    except Exception:
        pass
    return pd.DataFrame()


def cargar_archivos_excel(rutas: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    """Lee múltiples archivos Excel (tanto de experimentos BioGrow como exportados)."""
    marcos, advertencias = [], []
    mapa_var = {
        "B": "Semilla de maíz blanco (PBB-DH-183)",
        "R": "Semilla de maíz rosa (H-392 A)",
        "C": "Semilla de maíz cacahuazintle",
        "CR": "Semilla de maíz criollo",
        "CACAHUAZINTLE": "Semilla de maíz cacahuazintle",
        "CRIOLLO": "Semilla de maíz criollo",
    }

    for ruta in rutas:
        nombre = Path(ruta).name
        try:
            df_exp = _leer_excel_estandar(ruta)
            if not df_exp.empty:
                marcos.append(df_exp)
                continue

            todos = []
            with zipfile.ZipFile(ruta) as zf:
                cadenas = _leer_cadenas_compartidas(zf)
                for nombre_hoja, ruta_hoja in _mapa_hojas(zf):
                    celdas = _leer_celdas(zf, ruta_hoja, cadenas)
                    if not celdas:
                        continue
                    if _es_formato_temperatura(celdas):
                        tipo = TIPO_TEMPERATURA
                        regs = _extraer_hoja_temperatura(nombre_hoja, celdas, advertencias, nombre)
                    elif _es_formato_tiempo(celdas):
                        tipo = TIPO_TIEMPO30
                        regs = _extraer_hoja_tiempo(nombre_hoja, celdas)
                    else:
                        tipo = TIPO_TIEMPO
                        regs = _extraer_bloques_de_hoja(nombre_hoja, celdas, advertencias)
                    for r in regs:
                        r["Tipo de experimento"] = tipo
                        r["Archivo"] = nombre
                        r["Origen"] = "Excel"
                    todos.extend(regs)

            if todos:
                df_p = pd.DataFrame(todos)
                for col in ["Raíz", "Tallo", "Hojas", "Germinó", "Tratamiento (s)"]:
                    df_p[col] = pd.to_numeric(df_p[col], errors="coerce")
                crecio = (df_p["Raíz"].fillna(0) > 0) | (df_p["Tallo"].fillna(0) > 0)
                df_p.loc[df_p["Germinó"].isna() & crecio, "Germinó"] = 1
                df_p["Tratamiento"] = df_p["Tratamiento (s)"].apply(_etiqueta_tratamiento)
                df_p["Variedad"] = df_p["Color"].apply(lambda v: mapa_var.get(str(v).strip().upper(), str(v)))
                marcos.append(df_p)
            else:
                advertencias.append(f"{nombre}: no se reconoció ningún bloque de datos.")
        except Exception as e:
            advertencias.append(f"Error en {nombre}: {e}")

    if not marcos:
        return pd.DataFrame(), advertencias

    df = pd.concat(marcos, ignore_index=True)
    return _normalizar_columnas_finales(df), advertencias


# ===========================================================================
# 2. ADAPTADOR DE DATOS MYSQL -> DATAFRAME ESTANDARIZADO
# ===========================================================================

def convertir_mysql_a_df(registros_mysql: List[dict]) -> pd.DataFrame:
    """Convierte la salida de db.obtener_historial_con_id() al DataFrame de análisis."""
    if not registros_mysql:
        return pd.DataFrame()

    filas = []
    for r in registros_mysql:
        trat_raw = str(r.get("tiempo_tratamiento") or "").strip()
        lote_raw = str(r.get("nombre_lote") or "").strip()
        var_raw = str(r.get("tipo_semilla") or "").strip()
        raiz = float(r["longitud_raiz"]) if r.get("longitud_raiz") is not None else 0.0
        tallo = float(r["longitud_tallo"]) if r.get("longitud_tallo") is not None else 0.0
        hojas = float(r["cantidad_hojas"]) if r.get("cantidad_hojas") is not None else np.nan
        germino = int(r.get("germino", 1 if (raiz > 0 or tallo > 0) else 0))
        temp_raw = str(r.get("temperatura") or "Ambiente").strip()

        # Detectar si el tratamiento es control/patrón o extraer segundos/clase si usa nomenclatura T-1M
        es_control = (
            "control" in trat_raw.lower()
            or "patrón" in trat_raw.lower()
            or "patron" in trat_raw.lower()
            or trat_raw in ("0", "0s", "0 s")
            or bool(re.match(r"^\s*T\s*[-–]\s*0+\d+\s*$", trat_raw, re.IGNORECASE))
        )

        m_tit = PATRON_TITULO.match(trat_raw)
        if m_tit:
            cero, num_tirada, letra = m_tit.group(1), int(m_tit.group(2)), m_tit.group(3)
            if letra:
                cod_clase, nom_clase = _datos_clase(letra)
            elif cero:
                cod_clase, nom_clase = "S", "Patrón"
            else:
                cod_clase, nom_clase = "N", "Endospermo"
        else:
            num_tirada = int(r.get("experimento_id") or 1)
            cod_clase = "S" if es_control else "T"
            nom_clase = "Patrón" if es_control else "Tratada"

        # Intentar extraer segundos si dice "45s", "30 seg", etc.
        m_seg = PATRON_TIEMPO.match(trat_raw.replace(" ", ""))
        if es_control:
            seg = 0.0
            trat_etiqueta = "Control (0 s)"
        elif m_seg:
            seg = _a_segundos(m_seg.group(1), m_seg.group(2))
            trat_etiqueta = _etiqueta_tratamiento(seg)
        else:
            seg = np.nan
            trat_etiqueta = trat_raw

        fecha_reg = r.get("fecha_registro")
        fecha_str = fecha_reg.strftime("%Y-%m-%d") if hasattr(fecha_reg, "strftime") else str(fecha_reg or "")[:10]

        filas.append({
            "Archivo": "Base de Datos MySQL",
            "Origen": "MySQL",
            "Hoja": f"{lote_raw} ({fecha_str})" if fecha_str else lote_raw,
            "Lote": lote_raw,
            "Color": var_raw,
            "Variedad": var_raw,
            "Tirada": 0 if es_control else num_tirada,
            "Clase": cod_clase,
            "Clase nombre": nom_clase,
            "Bloque": trat_raw,
            "Réplica": int(r.get("experimento_id") or 1),
            "No": int(r.get("num_muestra") or 1),
            "Raíz": raiz,
            "Tallo": tallo,
            "Hojas": hojas,
            "Germinó": germino,
            "Tratamiento (s)": seg,
            "Tratamiento": trat_etiqueta,
            "Condición": "Estándar",
            "Temperatura": temp_raw,
            "Unidad original": "mm",
            "Tipo de experimento": TIPO_MYSQL,
        })

    df = pd.DataFrame(filas)
    return _normalizar_columnas_finales(df)


def _normalizar_columnas_finales(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    salida = df.copy()
    if "Lote" not in salida.columns:
        salida["Lote"] = salida["Hoja"].astype(str)
    salida["Estado germinación"] = salida["Germinó"].map({1: "Germinó", 0: "No germinó"})
    salida["Corrida"] = (
        salida["Origen"].astype(str) + " | "
        + salida["Hoja"].astype(str) + " | "
        + salida["Bloque"].astype(str) + " | rep "
        + salida["Réplica"].astype(str)
    )
    salida["ID semilla"] = salida["Corrida"] + " | No " + salida["No"].astype(str)
    return salida


def combinar_fuentes(df_mysql: pd.DataFrame, df_excel: pd.DataFrame, modo: str) -> pd.DataFrame:
    """modo: 'Solo MySQL', 'Solo Excel', 'MySQL + Excel (Combinados)'."""
    if modo == "Solo MySQL":
        return df_mysql.copy()
    if modo == "Solo Excel":
        return df_excel.copy()
    marcos = [d for d in (df_mysql, df_excel) if d is not None and not d.empty]
    if not marcos:
        return pd.DataFrame()
    return _normalizar_columnas_finales(pd.concat(marcos, ignore_index=True))


# ===========================================================================
# 3. MOTOR ESTADÍSTICO (ANOVA, TUKEY, LSD, CUARTILES, LETRAS CLD)
# ===========================================================================

def construir_grupo(df: pd.DataFrame, dimension: str) -> pd.Series:
    mapa_col = {
        "Tratamiento": "Tratamiento",
        "Temperatura (°C)": "Temperatura",
        "Variedad de Semilla": "Variedad",
        "Lote / Hoja": "Hoja",
        "Parte Expuesta (Clase)": "Clase nombre",
        "Origen (MySQL vs Excel)": "Origen",
        "Tratamiento + Variedad": ("Tratamiento", "Variedad"),
        "Tratamiento + Temperatura": ("Tratamiento", "Temperatura"),
    }


def preparar_datos_analisis(
    df: pd.DataFrame,
    variable: str,
    dimension_grupo: str,
    excluir_no_germinadas: bool = True,
    por_caja: bool = True,
    excluir_atipicos: bool = False,
    factor_iqr: float = 1.5,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve (df_semillas_limpio, df_unidades_analisis).
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    d = df.copy()
    d["Grupo"] = construir_grupo(d, dimension_grupo)
    d[variable] = pd.to_numeric(d[variable], errors="coerce")

    if variable in ("Raíz", "Tallo", "Hojas"):
        if excluir_no_germinadas:
            d.loc[d["Germinó"] == 0, variable] = np.nan
        else:
            d.loc[d["Germinó"] == 0, variable] = d.loc[d["Germinó"] == 0, variable].fillna(0.0)

        # Cálculo robusto de atípicos por IQR sin conflictos de longitudes
    if excluir_atipicos:
        # Convertir a valores numéricos limpios y extraer valores sin NaN
        vals_num = pd.to_numeric(d[variable], errors="coerce")
        s_clean = vals_num.dropna()
        
        if len(s_clean) >= 4:
            q1 = s_clean.quantile(0.25)
            q3 = s_clean.quantile(0.75)
            iqr = q3 - q1
            lim_inf = q1 - factor_iqr * iqr
            lim_sup = q3 + factor_iqr * iqr
            
            # Crear máscara booleana alineada exactamente con el índice de d
            fuera = vals_num.notna() & ((vals_num < lim_inf) | (vals_num > lim_sup))
            d = d[~fuera].copy()

    if variable == "Germinó" or por_caja:
        agregado = (
            d.groupby(["Corrida", "Grupo", "Clase", "Hoja", "Variedad"], dropna=False)
            .agg(**{
                variable: (variable, "mean"),
                "Semillas en la caja": (variable, "size"),
                "Semillas con dato": (variable, "count"),
            })
            .reset_index()
            .dropna(subset=[variable])
        )
        if variable == "Germinó":
            agregado[variable] = agregado[variable] * 100.0
        return d, agregado

    return d, d.dropna(subset=[variable]).copy()


def descriptivos_serie(serie, alfa=0.05):
    valores = pd.to_numeric(serie, errors="coerce").dropna()
    n = len(valores)
    base = {
        "n": n, "Media": np.nan, "Desviación estándar": np.nan,
        "Varianza": np.nan, "SEM": np.nan, "IC inf": np.nan, "IC sup": np.nan,
        "CV %": np.nan, "Q1": np.nan, "Mediana": np.nan,
        "Q3": np.nan, "IQR": np.nan, "Mínimo": np.nan, "Máximo": np.nan,
    }
    if n == 0:
        return base
    media = valores.mean()
    sd = valores.std(ddof=1) if n > 1 else 0.0
    var = valores.var(ddof=1) if n > 1 else 0.0
    sem = sd / math.sqrt(n) if n > 1 else 0.0
    if n > 1:
        t = stats.t.ppf(1 - alfa / 2, n - 1)
        ic_inf, ic_sup = media - t * sem, media + t * sem
    else:
        ic_inf = ic_sup = media
    q1, q3 = valores.quantile(0.25), valores.quantile(0.75)
    base.update({
        "Media": media, "Desviación estándar": sd, "Varianza": var,
        "SEM": sem, "IC inf": ic_inf, "IC sup": ic_sup,
        "CV %": (sd / media * 100) if (n > 1 and media != 0) else np.nan,
        "Q1": q1, "Mediana": valores.median(), "Q3": q3,
        "IQR": q3 - q1, "Mínimo": valores.min(), "Máximo": valores.max(),
    })
    return base


def tabla_descriptivos(df, columna_grupo, variable, alfa=0.05):
    if df.empty:
        return pd.DataFrame()
    filas = []
    for nombre, sub in df.groupby(columna_grupo, dropna=False):
        fila = {"Grupo": str(nombre)}
        fila.update(descriptivos_serie(sub[variable], alfa))
        filas.append(fila)
    salida = pd.DataFrame(filas)
    return salida.sort_values("Media", ascending=False).reset_index(drop=True) if not salida.empty else salida


def _grupos_validos(df, variable, columna_grupo="Grupo", minimo=2):
    grupos = []
    if df.empty:
        return grupos

    # Determinar dinámicamente la columna de agrupación correcta
    col_real = columna_grupo
    if col_real not in df.columns:
        if "Grupo" in df.columns:
            col_real = "Grupo"
        elif "Tratamiento" in df.columns:
            col_real = "Tratamiento"
        elif "Lote Origen" in df.columns:
            col_real = "Lote Origen"
        else:
            col_real = df.columns[0] if len(df.columns) > 0 else "Grupo"

    # Si por alguna razón col_real sigue sin estar en el DataFrame, usar la primera columna disponible
    if col_real not in df.columns:
        col_real = df.columns[0]

    for nombre, sub in df.groupby(col_real, dropna=False):
        valores = pd.to_numeric(sub[variable], errors="coerce").dropna()
        if len(valores) >= minimo:
            grupos.append((str(nombre), valores.astype(float)))
            
    return grupos


def anova_una_via(df, variable, columna_grupo="Grupo", alfa=0.05):
    grupos = _grupos_validos(df, variable, columna_grupo)
    k = len(grupos)
    if k < 2:
        return None
    n_total = sum(len(v) for _, v in grupos)
    gran_media = np.concatenate([v.values for _, v in grupos]).mean()

    sc_entre = sum(len(v) * (v.mean() - gran_media) ** 2 for _, v in grupos)
    sc_dentro = sum(((v - v.mean()) ** 2).sum() for _, v in grupos)
    sc_total = sc_entre + sc_dentro

    gl_entre, gl_dentro, gl_total = k - 1, n_total - k, n_total - 1
    if gl_dentro <= 0:
        return None

    cm_entre = sc_entre / gl_entre
    cm_dentro = sc_dentro / gl_dentro
    f_calc = cm_entre / cm_dentro if cm_dentro > 0 else np.inf
    f_crit = stats.f.ppf(1 - alfa, gl_entre, gl_dentro)
    p_valor = stats.f.sf(f_calc, gl_entre, gl_dentro)
    significativo = p_valor < alfa
    eta2 = sc_entre / sc_total if sc_total > 0 else np.nan

    signo = ">" if f_calc > f_crit else "≤"
    tabla = pd.DataFrame([
        {
            "Fuente": "Entre grupos (tratamiento)",
            "GL": gl_entre,
            "SC": sc_entre,
            "CM": cm_entre,
            "F calc": f_calc,
            "F crít": f_crit,
            "Comparación": f"{f_calc:.3f} {signo} {f_crit:.3f}",
            "p-valor": p_valor,
            "Conclusión": "✔ Sí difieren" if significativo else "✘ No difieren",
        },
        {
            "Fuente": "Dentro de grupos (error)",
            "GL": gl_dentro,
            "SC": sc_dentro,
            "CM": cm_dentro,
            "F calc": np.nan, "F crít": np.nan, "Comparación": "", "p-valor": np.nan, "Conclusión": "",
        },
        {
            "Fuente": "Total",
            "GL": gl_total,
            "SC": sc_total,
            "CM": np.nan, "F calc": np.nan, "F crít": np.nan, "Comparación": "", "p-valor": np.nan, "Conclusión": "",
        },
    ])

    medias = pd.DataFrame([
        {"Grupo": nombre, "n": len(v), "Media": v.mean(), "Varianza": v.var(ddof=1)}
        for nombre, v in grupos
    ])

    return {
        "tabla": tabla, "medias": medias,
        "CM error": cm_dentro, "GL error": gl_dentro, "GL entre": gl_entre,
        "k": k, "n total": n_total, "F calculada": f_calc, "F crítica": f_crit,
        "p-valor": p_valor, "significativo": significativo, "alfa": alfa,
        "GL total": gl_total, "SC entre": sc_entre, "SC error": sc_dentro,
        "CM entre": cm_entre, "eta2": eta2,
    }


def tukey_hsd(anova):
    if anova is None:
        return pd.DataFrame()
    medias, cm_error = anova["medias"], anova["CM error"]
    gl_error, k, alfa = anova["GL error"], anova["k"], anova["alfa"]
    if len(medias) < 2 or cm_error <= 0 or gl_error <= 0:
        return pd.DataFrame()

    q_crit = stats.studentized_range.ppf(1 - alfa, k, gl_error)
    filas = []
    for (_, g1), (_, g2) in combinations(medias.iterrows(), 2):
        dif = g1["Media"] - g2["Media"]
        error = math.sqrt((cm_error / 2) * (1 / g1["n"] + 1 / g2["n"]))
        q_calc = abs(dif) / error if error > 0 else np.inf
        dhs = q_crit * error
        p_aj = stats.studentized_range.sf(q_calc, k, gl_error)
        signif = abs(dif) > dhs
        filas.append({
            "Grupo 1": g1["Grupo"], "Grupo 2": g2["Grupo"],
            "n1": int(g1["n"]), "n2": int(g2["n"]),
            "Media 1": g1["Media"], "Media 2": g2["Media"],
            "|Diferencia|": abs(dif), "DHS": dhs,
            "Comparación": f"{abs(dif):.2f} {'>' if signif else '≤'} {dhs:.2f}",
            "q calc": q_calc, "q crít": q_crit, "p ajustada": p_aj,
            "Resultado": SIG if signif else NO_SIG,
            "Conclusión": "✔ Sí difieren" if signif else "✘ No difieren",
        })
    return pd.DataFrame(filas).sort_values("|Diferencia|", ascending=False).reset_index(drop=True)


def lsd_fisher(anova):
    if anova is None:
        return pd.DataFrame()
    medias, cm_error = anova["medias"], anova["CM error"]
    gl_error, alfa = anova["GL error"], anova["alfa"]
    if len(medias) < 2 or cm_error <= 0 or gl_error <= 0:
        return pd.DataFrame()

    t_crit = stats.t.ppf(1 - alfa / 2, gl_error)
    filas = []
    for (_, g1), (_, g2) in combinations(medias.iterrows(), 2):
        dif = abs(g1["Media"] - g2["Media"])
        error = math.sqrt(cm_error * (1 / g1["n"] + 1 / g2["n"]))
        t_calc = dif / error if error > 0 else np.inf
        valor_lsd = t_crit * error
        p = 2 * stats.t.sf(t_calc, gl_error)
        signif = dif > valor_lsd
        filas.append({
            "Grupo 1": g1["Grupo"], "Grupo 2": g2["Grupo"],
            "n1": int(g1["n"]), "n2": int(g2["n"]),
            "Media 1": g1["Media"], "Media 2": g2["Media"],
            "|Diferencia|": dif, "LSD": valor_lsd,
            "Comparación": f"{dif:.2f} {'>' if signif else '≤'} {valor_lsd:.2f}",
            "t calc": t_calc, "t crít": t_crit, "p-valor": p,
            "Resultado": SIG if signif else NO_SIG,
            "Conclusión": "✔ Sí difieren" if signif else "✘ No difieren",
        })
    return pd.DataFrame(filas).sort_values("|Diferencia|", ascending=False).reset_index(drop=True)


def letras_cld(medias, comparaciones_df):
    """Algoritmo inserción-absorción de Piepho (2004) para letras de significancia."""
    if medias is None or medias.empty or comparaciones_df is None or comparaciones_df.empty:
        return pd.DataFrame()

    medias = medias.sort_values("Media", ascending=False).reset_index(drop=True)
    grupos = medias["Grupo"].astype(str).tolist()

    distintos = set()
    for _, f in comparaciones_df.iterrows():
        if f["Resultado"] == SIG:
            distintos.add(frozenset((str(f["Grupo 1"]), str(f["Grupo 2"]))))

    columnas = [set(grupos)]
    for par in distintos:
        a, b = tuple(par)
        nuevas = []
        for col in columnas:
            if a in col and b in col:
                nuevas.append(col - {a})
                nuevas.append(col - {b})
            else:
                nuevas.append(col)
        nuevas = [c for c in nuevas if c]
        unicas = []
        for c in sorted(nuevas, key=len, reverse=True):
            if not any(c <= u for u in unicas):
                unicas.append(c)
        columnas = unicas

    posicion = {g: i for i, g in enumerate(grupos)}
    columnas.sort(key=lambda c: min(posicion[g] for g in c))

    abecedario = "abcdefghijklmnopqrstuvwxyz"
    letras = {g: "" for g in grupos}
    for i, col in enumerate(columnas):
        letra = abecedario[i] if i < 26 else abecedario[i // 26 - 1] + abecedario[i % 26]
        for g in grupos:
            if g in col:
                letras[g] += letra

    salida = medias.copy()
    salida["Letras"] = salida["Grupo"].astype(str).map(letras)
    return salida[["Grupo", "n", "Media", "Letras"]]


def tabla_cuartiles(datos, variable, factor=1.5):
    if datos.empty:
        return pd.DataFrame()
    filas = []
    for g, sub in datos.groupby("Grupo"):
        x = pd.to_numeric(sub[variable], errors="coerce").dropna().astype(float)
        n = len(x)
        if n == 0:
            continue
        q1, med, q3 = x.quantile([0.25, 0.5, 0.75]).tolist()
        iqr = q3 - q1
        lim_inf, lim_sup = q1 - factor * iqr, q3 + factor * iqr
        bowley = (q3 + q1 - 2 * med) / iqr if iqr > 0 else np.nan
        if n < 4 or pd.isna(bowley):
            forma = "Pocos datos (n<4)"
        elif abs(bowley) < 0.1:
            forma = "Simétrica"
        else:
            forma = f"{'Muy' if abs(bowley) >= 0.3 else 'Lig.'} sesgada ({'altos' if bowley > 0 else 'bajos'})"

        filas.append({
            "Grupo": str(g), "n": n, "Mínimo": x.min(), "Q1": q1, "Mediana": med,
            "Q3": q3, "Máximo": x.max(), "IQR": iqr,
            "Lím. Inf": lim_inf, "Lím. Sup": lim_sup,
            "Atípicos": int(((x < lim_inf) | (x > lim_sup)).sum()),
            "Bowley": bowley, "Forma": forma,
        })
    t = pd.DataFrame(filas)
    return t.sort_values("Mediana", ascending=False).reset_index(drop=True) if not t.empty else t


def ejecutar_suite_completa(df_semillas: pd.DataFrame, df_an: pd.DataFrame, variable: str, alfa: float = 0.05):
    """Ejecuta Descriptivos, ANOVA, Tukey, LSD, Letras CLD y Cuartiles en un solo llamado."""
    
    # 1. Traducción automática y segura del nombre de la variable
    col_var = variable
    if not df_semillas.empty and col_var not in df_semillas.columns:
        candidatos = []
        v_lower = str(variable).lower()
        if "raíz" in v_lower or "raiz" in v_lower:
            candidatos = ["longitud_raiz", "Raíz (mm)", "Raiz (mm)", "Raíz", "raiz"]
        elif "tallo" in v_lower:
            candidatos = ["longitud_tallo", "Tallo (mm)", "Tallo", "tallo"]
        elif "hoja" in v_lower:
            candidatos = ["cantidad_hojas", "Cant. Hojas", "Hojas", "hojas"]
        elif "germin" in v_lower:
            candidatos = ["germino", "Estado", "Germinó", "germinacion"]
        
        for c in candidatos:
            if c in df_semillas.columns:
                col_var = c
                break

    # Si en df_an tampoco existe con ese nombre, usar col_var o la primera numérica
    if not df_an.empty and "Grupo" not in df_an.columns:
        if "Tratamiento" in df_an.columns:
            df_an = df_an.rename(columns={"Tratamiento": "Grupo"})
        elif "Lote Origen" in df_an.columns:
            df_an = df_an.rename(columns={"Lote Origen": "Grupo"})
        elif len(df_an.columns) > 0:
            df_an = df_an.rename(columns={df_an.columns[0]: "Grupo"})

    desc = tabla_descriptivos(df_an, "Grupo", col_var, alfa)
    anova = anova_una_via(df_an, col_var, "Grupo", alfa)
    tk = tukey_hsd(anova) if anova else pd.DataFrame()
    ls = lsd_fisher(anova) if anova else pd.DataFrame()
    let_tk = letras_cld(anova["medias"], tk) if anova else pd.DataFrame()
    let_lsd = letras_cld(anova["medias"], ls) if anova else pd.DataFrame()

    # Extracción segura para cuartiles
    col_var_sem = col_var if (not df_semillas.empty and col_var in df_semillas.columns) else variable
    base_q = df_semillas.dropna(subset=[col_var_sem]) if variable != "Germinó" and not df_semillas.empty else df_an
    cuartiles = tabla_cuartiles(base_q, col_var_sem if col_var_sem in base_q.columns else variable)

    # Ranking consolidado
    ranking = pd.DataFrame()
    if not desc.empty:
        ranking = desc[["Grupo", "n", "Media", "Desviación estándar", "Varianza", "CV %"]].copy()
        ranking["Promedio ± DE"] = [
            f"{m:.2f} ± {s:.2f}" if pd.notna(s) else f"{m:.2f}"
            for m, s in zip(ranking["Media"], ranking["Desviación estándar"])
        ]
        mapa_tk = dict(zip(let_tk["Grupo"], let_tk["Letras"])) if not let_tk.empty and "Grupo" in let_tk.columns else {}
        mapa_ls = dict(zip(let_lsd["Grupo"], let_lsd["Letras"])) if not let_lsd.empty and "Grupo" in let_lsd.columns else {}
        ranking["Tukey"] = ranking["Grupo"].map(mapa_tk).fillna("-")
        ranking["LSD"] = ranking["Grupo"].map(mapa_ls).fillna("-")

        if not df_semillas.empty and "Germinó" in df_semillas.columns and "Grupo" in df_semillas.columns:
            g_grp = df_semillas.dropna(subset=["Germinó"]).groupby("Grupo")["Germinó"].agg(["sum", "count"])
            ranking["% Germ."] = [
                f"{g_grp.loc[g, 'sum'] / g_grp.loc[g, 'count'] * 100:.1f}% ({int(g_grp.loc[g, 'sum'])}/{int(g_grp.loc[g, 'count'])})"
                if g in g_grp.index and g_grp.loc[g, "count"] > 0 else "—"
                for g in ranking["Grupo"]
            ]
        else:
            ranking["% Germ."] = "—"

        ranking.insert(0, "Lugar", range(1, len(ranking) + 1))

    return {
        "descriptivos": desc,
        "ranking": ranking,
        "anova": anova,
        "tukey": tk,
        "lsd": ls,
        "letras_tukey": let_tk,
        "letras_lsd": let_lsd,
        "cuartiles": cuartiles,
    }


def generar_texto_conclusion(variable: str, agrupacion: str, unidad_txt: str, res: dict, alfa: float) -> str:
    anova = res.get("anova")
    desc = res.get("descriptivos")
    tk = res.get("tukey")
    if anova is None or desc is None or desc.empty:
        return "No hay suficientes grupos con al menos 2 réplicas para ejecutar el ANOVA."

    u = "%" if variable == "Germinó" else ("hojas" if variable == "Hojas" else "mm")
    signo_txt = "SÍ indicó diferencias estadísticamente significativas" if anova["significativo"] else "NO indicó diferencias significativas"
    txt = (
        f"RESUMEN EJECUTIVO ({variable} por {agrupacion}):\n"
        f"• Se analizaron {anova['n total']} {unidad_txt} distribuidas en {anova['k']} grupos.\n"
        f"• El ANOVA de una vía {signo_txt} "
        f"(F calc = {anova['F calculada']:.3f}, F crít = {anova['F crítica']:.3f}; "
        f"GL = {anova['GL entre']}, {anova['GL error']}; p = {anova['p-valor']:.4f}; α = {alfa}; η² = {anova['eta2']:.3f}).\n"
    )
    mejor = desc.iloc[0]
    txt += f"• El desempeño más alto lo obtuvo el grupo '{mejor['Grupo']}' con una media de {mejor['Media']:.2f} {u} (± {mejor['Desviación estándar']:.2f}).\n"

    if tk is not None and not tk.empty:
        sig_tk = tk[tk["Resultado"] == SIG]
        if not sig_tk.empty:
            pares = [f"{r['Grupo 1']} vs {r['Grupo 2']} (|Δ|={r['|Diferencia|']:.2f} > DHS={r['DHS']:.2f})" for _, r in sig_tk.head(4).iterrows()]
            txt += f"• Pares que difieren según Tukey HSD ({len(sig_tk)}): " + " | ".join(pares)
        else:
            txt += "• Según la prueba conservadora de Tukey HSD, ningún par individual superó la DHS."
    return txt
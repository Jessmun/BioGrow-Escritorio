"""
utils.py — Funciones auxiliares de BioGrow Desktop
Validaciones, formateo de datos y utilidades compartidas.
Sin dependencias de UI ni de base de datos.
"""

import math
from datetime import datetime
from typing import Optional, Tuple


# ---------------------------------------------------------------------------
# Regla de negocio: germinación
# ---------------------------------------------------------------------------

def calcular_germino(raiz: float, tallo: float) -> int:
    """
    Determina el estado de germinación según la regla de negocio del sistema.

    Regla:
    germino = 1 si raíz > 0.00 OR tallo > 0.00.
    Si ambas mediciones son 0.00, no germinó.

    Args:
        raiz:  Longitud de raíz en mm
        tallo: Longitud de tallo en mm

    Returns:
        1 si germinó, 0 si no germinó
    """
    return 1 if (raiz > 0.00 or tallo > 0.00) else 0


def estado_germinacion_texto(germino: int) -> str:
    """
    Convierte el valor entero de germinación a texto en español.

    Args:
        germino: 0 o 1

    Returns:
        "Germinó" o "No germinó"
    """
    return "Germinó" if germino == 1 else "No germinó"


# ---------------------------------------------------------------------------
# Validaciones de formulario
# ---------------------------------------------------------------------------

def validar_campos_lote(
    nombre_lote: str,
    tipo_semilla: str,
    fecha_registro: str,
    tiempo_tratamiento: str,
) -> Tuple[bool, str]:
    """
    Valida los campos del formulario principal antes de guardar.
    Replica las validaciones del JavaScript original.

    Args:
        nombre_lote:        Identificador del lote
        tipo_semilla:       Variedad de semilla seleccionada
        fecha_registro:     Fecha de registro (cadena)
        tiempo_tratamiento: Condición de tratamiento

    Returns:
        (True, "") si todo es válido
        (False, "mensaje de error") si hay un campo inválido
    """
    if not nombre_lote or not nombre_lote.strip():
        return False, "El campo 'Identificador del Lote' es obligatorio."

    if not tipo_semilla or not tipo_semilla.strip():
        return False, "Debe seleccionar una 'Variedad de Semilla'."

    if not fecha_registro or not fecha_registro.strip():
        return False, "El campo 'Fecha de Registro' es obligatorio."

    if not tiempo_tratamiento or not tiempo_tratamiento.strip():
        return False, "El campo 'Tiempo de Tratamiento' es obligatorio."

    return True, ""


def validar_valor_numerico(
    valor_str: str,
    nombre_campo: str = "valor"
) -> Tuple[bool, Optional[float], str]:
    """
    Valida y convierte una cadena para los campos de medición (raíz/tallo).

    Distingue un campo vacío (no capturado) de un 0.00 válido.

    Args:
        valor_str:    Cadena a convertir
        nombre_campo: Nombre del campo para el mensaje de error

    Returns:
        (True, None, "")              si el campo está vacío
        (True, valor_float, "")       si es un número >= 0
        (False, None, "mensaje")      si no es numérico o es negativo
    """
    try:
        valor = parsear_float(valor_str)
    except ValueError:
        return False, None, (
            f"'{valor_str}' no es un número válido para {nombre_campo}."
        )

    if valor is None:
        return True, None, ""   # Vacío: no capturado (distinto de 0.00)

    if valor < 0:
        return False, None, f"El {nombre_campo} no puede ser negativo."

    return True, valor, ""


def parsear_cantidad_hojas(texto: str) -> Optional[int]:
    """
    Convierte la cantidad de hojas capturada por el usuario.

    La cantidad de hojas es opcional.
    Si el campo está vacío, devuelve None para almacenarlo como NULL
    en la base de datos.

    Args:
        texto: Cantidad de hojas capturada.

    Returns:
        None si está vacío.
        Entero si contiene una cantidad válida.

    Raises:
        ValueError: Si el valor no es un entero válido o es negativo.
    """
    if texto is None or not texto.strip():
        return None

    texto = texto.strip()

    try:
        valor = int(texto)
    except ValueError:
        raise ValueError(
            "La cantidad de hojas debe ser un número entero."
        )

    if valor < 0:
        raise ValueError(
            "La cantidad de hojas no puede ser negativa."
        )

    return valor


# ---------------------------------------------------------------------------
# Formateo de datos
# ---------------------------------------------------------------------------

def formatear_fecha_actual() -> str:
    """
    Devuelve la fecha y hora actual formateada para el campo de registro.

    Formato: YYYY-MM-DD HH:MM

    Returns:
        Cadena de fecha/hora actual
    """
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def formatear_numero(valor: float, decimales: int = 2) -> str:
    """
    Formatea un número flotante con el número especificado de decimales.

    Args:
        valor:     Valor a formatear
        decimales: Número de decimales (por defecto 2)

    Returns:
        Cadena con el número formateado.
    """
    return f"{valor:.{decimales}f}"


def parsear_float(
    texto: str,
    default: Optional[float] = None
) -> Optional[float]:
    """
    Convierte una cadena a float.
    Maneja tanto punto (.) como coma (,) como separador decimal.

    Un campo vacío devuelve `default` (None por defecto), de modo que
    se distinga de un 0.00 capturado. Un texto no numérico NO se
    convierte en 0.0: lanza ValueError.

    Args:
        texto:   Cadena a convertir
        default: Valor de retorno si el campo está vacío

    Returns:
        Valor flotante parseado, o default si el campo está vacío

    Raises:
        ValueError: Si el texto no es un número finito válido.
    """
    if texto is None or not texto.strip():
        return default

    try:
        valor = float(texto.strip().replace(",", "."))
    except ValueError:
        raise ValueError(f"'{texto}' no es un número válido.")

    # Rechaza "nan", "inf", etc., que float() acepta
    if not math.isfinite(valor):
        raise ValueError(f"'{texto}' no es un número válido.")

    return valor


# ---------------------------------------------------------------------------
# Utilidades de interfaz
# ---------------------------------------------------------------------------

def truncar_texto(texto: str, max_chars: int = 25) -> str:
    """
    Trunca un texto largo y añade puntos suspensivos si excede el límite.

    Args:
        texto:     Texto a truncar
        max_chars: Número máximo de caracteres permitidos

    Returns:
        Texto truncado con "..." si excede max_chars,
        o texto original.
    """
    if len(texto) <= max_chars:
        return texto

    return texto[:max_chars - 3] + "..."


def centrar_ventana(ventana, ancho: int, alto: int) -> None:
    """
    Centra una ventana de Tkinter/CTk en la pantalla.

    Args:
        ventana: Instancia de Tk/CTk
        ancho:   Ancho de la ventana en píxeles
        alto:    Alto de la ventana en píxeles
    """
    ventana.update_idletasks()

    screen_w = ventana.winfo_screenwidth()
    screen_h = ventana.winfo_screenheight()

    x = (screen_w - ancho) // 2
    y = (screen_h - alto) // 2

    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

# ---------------------------------------------------------------------------

def validar_temperatura(temperatura_str: str) -> Tuple[bool, str, str]:
    """
    Valida el campo de temperatura en grados centígrados (°C).
    Acepta valores numéricos (ej. '35', '35.5', '25°C') o 'Ambiente'.
    Si se deja vacío, asigna 'Ambiente' por defecto.

    Returns:
        (True, temperatura_formateada, "") si es válido
        (False, "", mensaje_error) si no es válido
    """
    if not temperatura_str or not temperatura_str.strip():
        return True, "Ambiente", ""

    t = temperatura_str.strip()
    if t.lower() in ("amb", "ambiente", "t. amb", "tamb", "control"):
        return True, "Ambiente", ""

    # Limpiar sufijos como °C, c, °
    limpio = t.lower().replace("°c", "").replace("ºc", "").replace("°", "").replace("c", "").strip()
    try:
        val = float(limpio.replace(",", "."))
    except ValueError:
        return False, "", "La temperatura debe ser un número en °C (ej: 35) o 'Ambiente'."

    if val < -20 or val > 150:
        return False, "", "La temperatura en °C está fuera de un rango válido (-20 a 150 °C)."

    txt_num = f"{int(val)}" if val.is_integer() else f"{val:g}"
    return True, f"{txt_num} °C", ""
"""
models.py — Modelos de datos de BioGrow Desktop
Clases de datos puras (sin dependencias de UI ni de base de datos).
Modelan la estructura actual de MySQL.
"""

from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# Modelo de experimento
# ---------------------------------------------------------------------------

@dataclass
class Experimento:
    nombre_lote: str
    tipo_semilla: str
    fecha_registro: Optional[str] = None
    tiempo_tratamiento: str = ""
    temperatura: str = "Ambiente"
    id: Optional[int] = None
    

# ---------------------------------------------------------------------------
# Modelo de medición individual
# ---------------------------------------------------------------------------

@dataclass
class Medicion:
    """
    Representa una semilla dentro de un experimento.

    Corresponde a la tabla `mediciones_semillas` de MySQL.
    """

    experimento_id: int
    num_muestra: int
    longitud_raiz: float
    longitud_tallo: float
    cantidad_hojas: Optional[int]
    germino: int
    id: Optional[int] = None


# ---------------------------------------------------------------------------
# Vista de historial
# ---------------------------------------------------------------------------

@dataclass
class RegistroHistorial:
    nombre_lote: str
    tipo_semilla: str
    num_muestra: int
    tiempo_tratamiento: str
    longitud_raiz: float
    longitud_tallo: float
    cantidad_hojas: Optional[int]
    germino: int
    temperatura: str = "Ambiente"

    @property
    def resultado_texto(self) -> str:
        """Devuelve la etiqueta de germinación en español."""
        return "Germinó" if self.germino == 1 else "No germinó"

    @property
    def raiz_fmt(self) -> str:
        """Longitud de raíz con 2 decimales."""
        return f"{self.longitud_raiz:.2f}"

    @property
    def tallo_fmt(self) -> str:
        """Longitud de tallo con 2 decimales."""
        return f"{self.longitud_tallo:.2f}"

    @property
    def hojas_fmt(self) -> str:
        """Cantidad de hojas. Vacío cuando no fue registrada."""
        if self.cantidad_hojas is None:
            return ""
        return str(self.cantidad_hojas)


# ---------------------------------------------------------------------------
# Contenedor de datos del formulario
# ---------------------------------------------------------------------------

@dataclass
class DatosLote:
    nombre_lote: str
    tipo_semilla: str
    fecha_registro: Optional[str]
    tiempo_tratamiento: str
    temperatura: str = "Ambiente"
    mediciones: list = field(default_factory=list)
    gas: str = "Sin gas"
    flujo: Optional[float] = None
    orientacion: str = "Aleatorio"
    morfologia: str = "Plana"
    cantidad_semillas: int = 27
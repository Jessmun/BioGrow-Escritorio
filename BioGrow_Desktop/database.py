"""
database.py — Capa de acceso a datos de BioGrow Desktop (Versión SQLite Local)
Gestión de base de datos local autocontenida (biogrow.db).
"""

import os
import sqlite3
from typing import List, Optional, Tuple

from models import (
    Experimento,
    Medicion,
    RegistroHistorial,
    DatosLote,
)
from utils import calcular_germino

DB_PATH = os.getenv("BIOGROW_DB_PATH", "biogrow.db")


class DatabaseManager:
    """Gestiona todas las operaciones de base de datos local SQLite de BioGrow."""

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or DB_PATH
        self._conn = None

    def _connect(self):
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            return conn
        except Exception as e:
            raise ConnectionError(f"No fue posible conectar con la base de datos SQLite: {e}") from e

    def __enter__(self):
        self._conn = self._connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._conn is not None:
            if exc_type:
                self._conn.rollback()
            else:
                self._conn.commit()
            self._conn.close()
        return False

    def init_db(self) -> None:
        conn = self._connect()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS experimentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre_lote TEXT NOT NULL,
                    tipo_semilla TEXT NOT NULL,
                    fecha_registro TEXT NOT NULL,
                    tiempo_tratamiento TEXT NOT NULL,
                    temperatura TEXT DEFAULT 'Ambiente'
                )
            """)

            try:
                cursor.execute("ALTER TABLE experimentos ADD COLUMN temperatura TEXT DEFAULT 'Ambiente'")
            except sqlite3.OperationalError:
                pass

            # Factores nuevos del tratamiento. Se agregan sin tocar los lotes ya registrados
            # (en ellos quedan vacíos y el análisis los muestra como 'Sin dato').
            for columna, tipo in [("gas", "TEXT"), ("flujo", "REAL"), ("orientacion", "TEXT"),
                                  ("morfologia", "TEXT"), ("cantidad_semillas", "INTEGER")]:
                try:
                    cursor.execute(f"ALTER TABLE experimentos ADD COLUMN {columna} {tipo}")
                except sqlite3.OperationalError:
                    pass

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS mediciones_semillas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    experimento_id INTEGER NOT NULL,
                    num_muestra INTEGER NOT NULL,
                    longitud_raiz REAL NOT NULL,
                    longitud_tallo REAL NOT NULL,
                    cantidad_hojas INTEGER NULL,
                    germino INTEGER NOT NULL,
                    CONSTRAINT fk_mediciones_experimento
                        FOREIGN KEY (experimento_id)
                        REFERENCES experimentos(id)
                        ON DELETE CASCADE
                        ON UPDATE CASCADE,
                    CONSTRAINT uq_muestra_experimento
                        UNIQUE (experimento_id, num_muestra)
                )
            """)
            conn.commit()
        finally:
            cursor.close()
            conn.close()

    def guardar_lote(self, datos: DatosLote) -> Tuple[bool, str]:
        if not datos.nombre_lote or not datos.tipo_semilla or not datos.tiempo_tratamiento:
            return False, "Faltan campos obligatorios: Lote, Variedad y Tiempo de Tratamiento."

        if not datos.mediciones:
            return False, "No se proporcionaron mediciones de semillas."

        conn = None
        cursor = None
        try:
            conn = self._connect()
            conn.execute("BEGIN TRANSACTION;")
            cursor = conn.cursor()

            fecha_limpia = datos.fecha_registro
            if fecha_limpia:
                fecha_limpia = fecha_limpia.replace("T", " ")
                if len(fecha_limpia) == 16:
                    fecha_limpia += ":00"

            cursor.execute(
                """
                INSERT INTO experimentos (
                    nombre_lote, tipo_semilla, fecha_registro, tiempo_tratamiento, temperatura,
                    gas, flujo, orientacion, morfologia, cantidad_semillas
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    datos.nombre_lote,
                    datos.tipo_semilla,
                    fecha_limpia,
                    datos.tiempo_tratamiento,
                    getattr(datos, "temperatura", None) or "Ambiente",
                    datos.gas, datos.flujo, datos.orientacion, datos.morfologia, datos.cantidad_semillas,
                ),
            )
            experimento_id = cursor.lastrowid

            for med in datos.mediciones:
                num_muestra = med.get("num_muestra")
                raiz = med.get("longitud_raiz")
                tallo = med.get("longitud_tallo")
                cantidad_hojas = med.get("cantidad_hojas")

                if num_muestra is None or raiz is None or tallo is None or cantidad_hojas is None:
                    raise ValueError(f"Medición incompleta (muestra {num_muestra}): se requieren raíz, tallo y hojas.")

                num_muestra = int(num_muestra)
                raiz = float(raiz)
                tallo = float(tallo)
                germino = calcular_germino(raiz, tallo)

                cursor.execute(
                    """
                    INSERT INTO mediciones_semillas (
                        experimento_id, num_muestra, longitud_raiz,
                        longitud_tallo, cantidad_hojas, germino
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        experimento_id,
                        num_muestra,
                        raiz,
                        tallo,
                        cantidad_hojas,
                        germino,
                    ),
                )

            conn.commit()
            return True, f"Lote '{datos.nombre_lote}' guardado correctamente con {len(datos.mediciones)} mediciones."
        except Exception as e:
            if conn is not None:
                conn.rollback()
            return False, f"Error al guardar el lote: {str(e)}"
        finally:
            if cursor is not None:
                cursor.close()
            if conn is not None:
                conn.close()

    def obtener_historial(
        self,
        texto_busqueda: str = "",
        filtro_semilla: str = "",
        filtro_germino: str = "",
    ) -> List[RegistroHistorial]:
        rows = self.obtener_historial_con_id(texto_busqueda, filtro_semilla, filtro_germino)
        return [
            RegistroHistorial(
                nombre_lote=row["nombre_lote"],
                tipo_semilla=row["tipo_semilla"],
                num_muestra=row["num_muestra"],
                tiempo_tratamiento=row["tiempo_tratamiento"],
                longitud_raiz=float(row["longitud_raiz"]),
                longitud_tallo=float(row["longitud_tallo"]),
                cantidad_hojas=(
                    int(row["cantidad_hojas"])
                    if row["cantidad_hojas"] is not None
                    else None
                ),
                germino=int(row["germino"]),
                temperatura=row.get("temperatura") or "Ambiente",
            )
            for row in rows
        ]

    def obtener_historial_con_id(
        self,
        texto_busqueda: str = "",
        filtro_semilla: str = "",
        filtro_germino: str = "",
    ) -> List[dict]:
        query = """
            SELECT
                m.id AS medicion_id,
                e.id AS experimento_id,
                e.nombre_lote AS nombre_lote,
                e.tipo_semilla AS tipo_semilla,
                e.fecha_registro AS fecha_registro,
                m.num_muestra AS num_muestra,
                e.tiempo_tratamiento AS tiempo_tratamiento,
                COALESCE(e.temperatura, 'Ambiente') AS temperatura,
                e.gas AS gas,
                e.flujo AS flujo,
                e.orientacion AS orientacion,
                e.morfologia AS morfologia,
                e.cantidad_semillas AS cantidad_semillas,
                m.longitud_raiz AS longitud_raiz,
                m.longitud_tallo AS longitud_tallo,
                m.cantidad_hojas AS cantidad_hojas,
                m.germino AS germino
            FROM mediciones_semillas m
            INNER JOIN experimentos e
                ON m.experimento_id = e.id
            WHERE 1 = 1
        """
        params = []
        if texto_busqueda:
            like = f"%{texto_busqueda}%"
            query += """
                AND (
                    e.nombre_lote LIKE ?
                    OR e.tiempo_tratamiento LIKE ?
                    OR e.tipo_semilla LIKE ?
                    OR e.temperatura LIKE ?
                    OR e.gas LIKE ?
                    OR e.orientacion LIKE ?
                    OR e.morfologia LIKE ?
                )
            """
            params.extend([like] * 7)

        if filtro_semilla:
            query += " AND e.tipo_semilla = ?"
            params.append(filtro_semilla)

        if filtro_germino in ("0", "1"):
            query += " AND m.germino = ?"
            params.append(int(filtro_germino))

        query += " ORDER BY m.id DESC"

        conn = None
        cursor = None
        try:
            conn = self._connect()
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
        finally:
            if cursor is not None:
                cursor.close()
            if conn is not None:
                conn.close()

    def actualizar_medicion(
        self,
        medicion_id: int,
        experimento_id: int,
        nombre_lote: str,
        tipo_semilla: str,
        tiempo_tratamiento: str,
        longitud_raiz: float,
        longitud_tallo: float,
        cantidad_hojas: int,
        temperatura: str = "Ambiente",
        gas: Optional[str] = None,
        flujo: Optional[float] = None,
        orientacion: Optional[str] = None,
        morfologia: Optional[str] = None,
    ) -> Tuple[bool, str]:
        conn = None
        cursor = None
        try:
            conn = self._connect()
            conn.execute("BEGIN TRANSACTION;")
            cursor = conn.cursor()

            germino = calcular_germino(longitud_raiz, longitud_tallo)

            cursor.execute(
                """
                UPDATE experimentos
                SET nombre_lote = ?, tipo_semilla = ?, tiempo_tratamiento = ?, temperatura = ?,
                    gas = ?, flujo = ?, orientacion = ?, morfologia = ?
                WHERE id = ?
                """,
                (nombre_lote, tipo_semilla, tiempo_tratamiento, temperatura, gas, flujo, orientacion,
                 morfologia, int(experimento_id)),
            )

            cursor.execute(
                """
                UPDATE mediciones_semillas
                SET longitud_raiz = ?, longitud_tallo = ?, cantidad_hojas = ?, germino = ?
                WHERE id = ?
                """,
                (float(longitud_raiz), float(longitud_tallo), int(cantidad_hojas), int(germino), int(medicion_id)),
            )

            conn.commit()
            return True, "Registro actualizado correctamente."
        except Exception as e:
            if conn:
                conn.rollback()
            return False, f"Error al actualizar el registro: {e}"
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def eliminar_medicion(self, medicion_id: int, num_muestra: Optional[int] = None) -> Tuple[bool, str]:
        conn = None
        cursor = None
        try:
            conn = self._connect()
            cursor = conn.cursor()
            if num_muestra is not None:
                cursor.execute(
                    """
                    DELETE FROM mediciones_semillas
                    WHERE id IN (
                        SELECT m.id FROM mediciones_semillas m
                        INNER JOIN experimentos e ON m.experimento_id = e.id
                        WHERE e.nombre_lote = ? AND m.num_muestra = ?
                    )
                    """,
                    (medicion_id, num_muestra),
                )
            else:
                cursor.execute(
                    "DELETE FROM mediciones_semillas WHERE id = ?",
                    (int(medicion_id),)
                )
            conn.commit()
            if cursor.rowcount > 0:
                return True, "Medición eliminada correctamente."
            return False, "No se encontró el registro seleccionado."
        except Exception as e:
            if conn:
                conn.rollback()
            return False, f"Error al eliminar la medición: {e}"
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def eliminar_lote_por_nombre(self, nombre_lote: str) -> Tuple[bool, str]:
        conn = None
        cursor = None
        try:
            conn = self._connect()
            cursor = conn.cursor()
            cursor.execute("PRAGMA foreign_keys = ON;")
            cursor.execute(
                "DELETE FROM experimentos WHERE nombre_lote = ?",
                (nombre_lote,)
            )
            conn.commit()
            if cursor.rowcount > 0:
                return True, f"Lote '{nombre_lote}' y todas sus muestras fueron eliminados."
            return False, f"No se encontró el lote '{nombre_lote}'."
        except Exception as e:
            if conn:
                conn.rollback()
            return False, f"Error al eliminar el lote: {e}"
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def nombres_de_lotes(self) -> set:
        """Pares (nombre_lote, tipo_semilla) ya registrados, para no importar dos veces el mismo lote."""
        conn = self._connect()
        try:
            return {(r[0], r[1]) for r in conn.execute("SELECT nombre_lote, tipo_semilla FROM experimentos")}
        finally:
            conn.close()

    def guardar_lotes_importados(self, lotes: list) -> Tuple[bool, str]:
        """Guarda varios lotes leídos de Excel en una sola transacción (todo o nada).

        Cada lote es un dict con nombre_lote, tipo_semilla, fecha_registro, tiempo_tratamiento,
        temperatura y mediciones [{num_muestra, longitud_raiz, longitud_tallo, cantidad_hojas, germino}].
        A diferencia de la captura manual, la cantidad de hojas puede venir vacía y el estado de
        germinación se toma del Excel.
        """
        conn = None
        try:
            conn = self._connect()
            conn.execute("BEGIN TRANSACTION;")
            n_sem = 0
            for lote in lotes:
                cur = conn.execute(
                    "INSERT INTO experimentos (nombre_lote, tipo_semilla, fecha_registro, tiempo_tratamiento, "
                    "temperatura, gas, flujo, orientacion, morfologia, cantidad_semillas) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (lote["nombre_lote"], lote["tipo_semilla"], lote["fecha_registro"],
                     lote["tiempo_tratamiento"], lote.get("temperatura") or "Ambiente",
                     lote.get("gas"), lote.get("flujo"), lote.get("orientacion"), lote.get("morfologia"),
                     lote.get("cantidad_semillas") or len(lote["mediciones"])),
                )
                exp_id = cur.lastrowid
                for med in lote["mediciones"]:
                    conn.execute(
                        "INSERT INTO mediciones_semillas (experimento_id, num_muestra, longitud_raiz, "
                        "longitud_tallo, cantidad_hojas, germino) VALUES (?, ?, ?, ?, ?, ?)",
                        (exp_id, int(med["num_muestra"]), float(med["longitud_raiz"]),
                         float(med["longitud_tallo"]), med.get("cantidad_hojas"), int(med["germino"])),
                    )
                    n_sem += 1
            conn.commit()
            return True, f"Se importaron {len(lotes)} lotes con {n_sem} semillas a la base de datos."
        except Exception as e:
            if conn is not None:
                conn.rollback()
            return False, f"No se importó nada (error al guardar): {e}"
        finally:
            if conn is not None:
                conn.close()

    def contar_registros(self) -> dict:
        conn = None
        cursor = None
        try:
            conn = self._connect()
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) AS total FROM mediciones_semillas")
            total = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) AS germinadas FROM mediciones_semillas WHERE germino = 1")
            germinadas = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(DISTINCT id) AS lotes FROM experimentos")
            lotes = cursor.fetchone()[0]

            return {
                "total": total,
                "germinadas": germinadas,
                "no_germinadas": total - germinadas,
                "lotes": lotes,
            }
        finally:
            if cursor is not None:
                cursor.close()
            if conn is not None:
                conn.close()
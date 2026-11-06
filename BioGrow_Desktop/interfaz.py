"""
interfaz.py — Interfaz gráfica de BioGrow Desktop
- Colores duales adaptativos (Modo Claro / Modo Oscuro sin zonas oscuras residuales).
- Selector de páginas en la esquina superior izquierda.
- Botón de cambio de tema (reemplaza al switch).
- Modo Administrador con Edición y Eliminación de registros y lotes.
- Exportación a Excel (.xlsx) y PDF (.pdf) con diseño profesional.
"""
import json
import re
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import customtkinter as ctk
from datetime import datetime
from typing import Optional

from database import DatabaseManager
from models import DatosLote
from utils import (
    calcular_germino,
    validar_campos_lote,
    validar_valor_numerico,
    parsear_cantidad_hojas,
    formatear_numero,
    formatear_fecha_actual,
    centrar_ventana,
    truncar_texto,
    validar_temperatura,
    validar_flujo,
    validar_cantidad_semillas,
    TIPOS_GAS,
    TIPOS_TRATAMIENTO,
    ORIENTACIONES,
    MORFOLOGIAS,
)

import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# Análisis y gráficas: mismos cálculos que el dashboard (dashboardst.py)
import analisis_biogrow as AN
import graficas_biogrow as GR
import reporte_biogrow as REP
import exportar_biogrow as EX
import fondos_biogrow as FD
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Tuplas reactivas (Modo Claro, Modo Oscuro) para CustomTkinter
UI_COLORS = {
    "bg_primary":     ("#f1f5f9", "#0b0f19"),
    "bg_card":        ("#ffffff", "#161c2d"),
    "bg_secondary":   ("#e2e8f0", "#111827"),
    "bg_card_light":  ("#f8fafc", "#1e2a3a"),
    "input_bg":       ("#f8fafc", "#0f172a"),
    "text_primary":   ("#0f172a", "#f8fafc"),
    "text_secondary": ("#475569", "#94a3b8"),
    "text_muted":     ("#64748b", "#64748b"),
    "border":         ("#cbd5e1", "#1e293b"),
    "accent_green":   ("#059669", "#10b981"),
    "accent_hover":   ("#047857", "#34d399"),
    "accent_indigo":  ("#4f46e5", "#6366f1"),
    "accent_red":     ("#dc2626", "#ef4444"),
    "accent_amber":   ("#d97706", "#f59e0b"),
}

# Colores planos para ttk.Treeview y tk.Entry (que no aceptan tuplas)
TTK_PALETTES = {
    "light": {
        "bg_primary":    "#f1f5f9",
        "bg_card":       "#ffffff",
        "bg_secondary":  "#e2e8f0",
        "bg_card_light": "#e2e8f0",
        "input_bg":      "#ffffff",
        "text_primary":  "#0f172a",
        "text_secondary":"#334155",
        "text_muted":    "#64748b",
        "border":        "#cbd5e1",
        "row_alt":       "#f1f5f9",
        "accent_green":  "#059669",
        "accent_indigo": "#4f46e5",
        "accent_red":    "#dc2626",
    },
    "dark": {
        "bg_primary":    "#0b0f19",
        "bg_card":       "#161c2d",
        "bg_secondary":  "#111827",
        "bg_card_light": "#1e2a3a",
        "input_bg":      "#0f172a",
        "text_primary":  "#f8fafc",
        "text_secondary":"#94a3b8",
        "text_muted":    "#64748b",
        "border":        "#1e293b",
        "row_alt":       "#12192a",
        "accent_green":  "#10b981",
        "accent_indigo": "#6366f1",
        "accent_red":    "#ef4444",
    },
}

NUM_MUESTRAS = 27
# Variedades con su nombre técnico (híbrido) o científico entre paréntesis (catálogo en analisis_biogrow)
TIPOS_SEMILLA = list(AN.TIPOS_SEMILLA)
TRATAMIENTOS_BASE = ["Control", "30s", "45s", "60s"]
ADMIN_PIN = "admin123"


class BioGrowApp(ctk.CTk):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self._modo_actual = "dark"
        self._admin_activo = False

        self._filas_captura: list = []
        self._num_muestras: int = NUM_MUESTRAS
        self._datos_historial_actual: list = []
        self._entry_activo: Optional[tk.Widget] = None
        self._celda_activa: Optional[tuple] = None
        self._mostrando_error: bool = False

        self._archivos_excel_cargados: list = []
        self._df_excel_cache: pd.DataFrame = pd.DataFrame()
        self._advertencias_excel: list = []
        self._ultimo_resultado_stats: Optional[dict] = None
        self._fig_actual: Optional[Figure] = None

        self._configurar_ventana()
        self._aplicar_estilos_ttk()
        self._construir_ui()
        self._inicializar_tabla_captura()
        self._cargar_historial()

    def _configurar_ventana(self) -> None:
        self.title("BioGrow Analytics — Monitor de Germinación y Biometría de Semillas")
        self.configure(fg_color=UI_COLORS["bg_primary"])
        centrar_ventana(self, 1540, 940)
        self.minsize(1280, 780)

    def _aplicar_estilos_ttk(self) -> None:
        p = TTK_PALETTES[self._modo_actual]
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "BioGrow.Treeview",
            background=p["bg_card"],
            foreground=p["text_primary"],
            fieldbackground=p["bg_card"],
            borderwidth=0,
            rowheight=32,
            font=("Segoe UI", 10),
        )
        style.configure(
            "BioGrow.Treeview.Heading",
            background=p["bg_secondary"],
            foreground=p["text_secondary"],
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
        )
        style.map(
            "BioGrow.Treeview",
            background=[("selected", p["accent_indigo"])],
            foreground=[("selected", "#ffffff")],
        )
        style.map(
            "BioGrow.Treeview.Heading",
            background=[("active", p["bg_card_light"])],
        )

        style.configure(
            "Captura.Treeview",
            background=p["bg_card"],
            foreground=p["text_primary"],
            fieldbackground=p["bg_card"],
            borderwidth=0,
            rowheight=28,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Captura.Treeview.Heading",
            background=p["bg_secondary"],
            foreground=p["text_secondary"],
            relief="flat",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Captura.Treeview",
            background=[("selected", p["accent_indigo"])],
            foreground=[("selected", "#ffffff")],
        )
        style.map(
            "Captura.Treeview.Heading",
            background=[("active", p["bg_card_light"])],
        )

        style.configure(
            "Dark.Vertical.TScrollbar",
            background=p["bg_secondary"],
            troughcolor=p["bg_primary"],
            arrowcolor=p["text_muted"],
            borderwidth=0,
        )
        style.configure(
            "Dark.Horizontal.TScrollbar",
            background=p["bg_secondary"],
            troughcolor=p["bg_primary"],
            arrowcolor=p["text_muted"],
            borderwidth=0,
        )

        # Actualizar colores de tags en tablas si ya fueron creadas
        if hasattr(self, "_tree_captura"):
            self._tree_captura.tag_configure("germino_si", foreground=p["accent_green"], background=p["bg_card"])
            self._tree_captura.tag_configure("germino_no", foreground=p["text_muted"], background=p["bg_card"])
            self._tree_captura.tag_configure("row_alt", background=p["row_alt"], foreground=p["text_primary"])
            self._tree_captura.tag_configure("row_norm", background=p["bg_card"], foreground=p["text_primary"])
            self._tree_captura.tag_configure("alt_si", background=p["row_alt"], foreground=p["accent_green"])
            self._tree_captura.tag_configure("alt_no", background=p["row_alt"], foreground=p["text_muted"])

        if hasattr(self, "_tree_historial"):
            t = self._tree_historial
            t.tag_configure("germino_si", foreground=p["accent_green"], background=self._bg_fila(t, p["bg_card"]))
            t.tag_configure("germino_no", foreground=p["accent_red"], background=self._bg_fila(t, p["bg_card"]))
            t.tag_configure("alt_si", background=self._bg_fila(t, p["row_alt"]), foreground=p["accent_green"])
            t.tag_configure("alt_no", background=self._bg_fila(t, p["row_alt"]), foreground=p["accent_red"])
        self._refrescar_fondos_tablas()

    def _construir_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._construir_header()
        self._construir_paginas()

    def _construir_header(self) -> None:
        self._header = ctk.CTkFrame(
            self,
            fg_color=UI_COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        self._header.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
        self._header.grid_columnconfigure(2, weight=1)

        # 1. Logo y Título (Izquierda)
        brand_frame = ctk.CTkFrame(self._header, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=(16, 12), pady=10, sticky="w")

        ctk.CTkLabel(brand_frame, text="🌱", font=ctk.CTkFont(size=28)).pack(side="left", padx=(0, 10))

        titles = ctk.CTkFrame(brand_frame, fg_color="transparent")
        titles.pack(side="left")
        ctk.CTkLabel(
            titles,
            text="BioGrow Analytics",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=UI_COLORS["text_primary"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            titles,
            text="Monitor de Germinación y Biometría",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=UI_COLORS["text_secondary"],
        ).pack(anchor="w")

        # Separador vertical sutil
        ctk.CTkFrame(self._header, width=1, height=34, fg_color=UI_COLORS["border"]).grid(
            row=0, column=1, padx=8, pady=10
        )

        # 2. Navegación Superior Izquierda (Compacta y atractiva)
        self._nav_bar = ctk.CTkSegmentedButton(
            self._header,
            values=["🌱 Registro & Historial", "📐 Cálculos y Modelos", "📊 Tablero & Gráficas"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=UI_COLORS["bg_secondary"],
            selected_color=UI_COLORS["accent_indigo"],
            selected_hover_color=UI_COLORS["accent_indigo"],
            unselected_color=UI_COLORS["bg_secondary"],
            unselected_hover_color=UI_COLORS["bg_card_light"],
            text_color=UI_COLORS["text_primary"],
            height=34,
            corner_radius=8,
            command=self._cambiar_pagina,
        )
        self._nav_bar.set("🌱 Registro & Historial")
        self._nav_bar.grid(row=0, column=2, padx=12, pady=10, sticky="w")

        # 3. Controles Derecha (Estado BD sutil, Admin, Botón Tema)
        right_box = ctk.CTkFrame(self._header, fg_color="transparent")
        right_box.grid(row=0, column=3, padx=16, pady=10, sticky="e")

        db_box = ctk.CTkFrame(right_box, fg_color="transparent")
        db_box.pack(side="left", padx=(0, 14))
        ctk.CTkLabel(
            db_box, text="●", font=ctk.CTkFont(size=9), text_color=UI_COLORS["accent_green"]
        ).pack(side="left", padx=(0, 4))
        ctk.CTkLabel(
            db_box,
            text="BD conectada",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left")

        self._btn_admin = ctk.CTkButton(
            right_box,
            text="🔒 Admin",
            width=85,
            height=32,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"],
            hover_color=UI_COLORS["bg_card_light"],
            border_width=1,
            border_color=UI_COLORS["border"],
            command=self._toggle_admin_mode,
        )
        self._btn_admin.pack(side="left", padx=(0, 10))

        self._btn_tema = ctk.CTkButton(
            right_box,
            text="☀️ Modo Claro",
            width=115,
            height=32,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"],
            hover_color=UI_COLORS["bg_card_light"],
            border_width=1,
            border_color=UI_COLORS["border"],
            command=self._cambiar_tema,
        )
        self._btn_tema.pack(side="left")

    def _construir_paginas(self) -> None:
        self._contenedor_paginas = ctk.CTkFrame(self, fg_color="transparent")
        self._contenedor_paginas.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        self._contenedor_paginas.grid_columnconfigure(0, weight=1)
        self._contenedor_paginas.grid_rowconfigure(0, weight=1)

        self._pag_registro = ctk.CTkFrame(self._contenedor_paginas, fg_color="transparent")
        self._pag_calculos = ctk.CTkFrame(self._contenedor_paginas, fg_color="transparent")
        self._pag_tablero  = ctk.CTkFrame(self._contenedor_paginas, fg_color="transparent")

        self._construir_cuerpo_registro(self._pag_registro)
        self._construir_seccion_calculos(self._pag_calculos)
        self._construir_seccion_tablero(self._pag_tablero)

        self._pag_registro.grid(row=0, column=0, sticky="nsew")

    # --- Ilustraciones difuminadas detrás de las tablas -----------------------

    def _activar_fondo_tabla(self, tree: ttk.Treeview, clave: str, estilo_base: str) -> None:
        """La tabla muestra la ilustración de su pestaña detrás de las filas (filas sin fondo propio)."""
        if not FD.disponible(clave):
            return
        if not hasattr(self, "_fondos_tabla"):
            self._fondos_tabla = {}
        n = len(self._fondos_tabla) + 1
        nombre, elemento = f"Fondo{n}.{estilo_base}", f"Fondo{n}.field"
        foto = tk.PhotoImage(master=self, width=1, height=1)
        style = ttk.Style(self)
        try:
            style.element_create(elemento, "image", foto, sticky="nsew")
            style.layout(nombre, [(elemento, {"sticky": "nswe", "children": [
                ("Treeview.padding", {"sticky": "nswe", "children": [("Treeview.treearea", {"sticky": "nswe"})]})]})])
        except tk.TclError:
            return
        tree.configure(style=nombre)
        self._fondos_tabla[tree] = {"clave": clave, "foto": foto, "estilo": nombre, "tam": None, "job": None}
        self._estilo_fondo_tabla(nombre)
        tree.bind("<Configure>", lambda e, t=tree: self._programar_fondo_tabla(t), add="+")

    def _estilo_fondo_tabla(self, nombre: str) -> None:
        p = TTK_PALETTES[self._modo_actual]
        style = ttk.Style(self)
        style.configure(nombre, background="")  # filas transparentes: se ve la ilustración
        style.map(nombre, background=[("selected", p["accent_indigo"])], foreground=[("selected", "#ffffff")])

    def _programar_fondo_tabla(self, tree) -> None:
        info = getattr(self, "_fondos_tabla", {}).get(tree)
        if not info:
            return
        if info["job"] is not None:
            self.after_cancel(info["job"])
        info["job"] = self.after(120, lambda: self._dibujar_fondo_tabla(tree))

    def _dibujar_fondo_tabla(self, tree, forzar: bool = False) -> None:
        info = self._fondos_tabla.get(tree)
        if not info:
            return
        info["job"] = None
        try:
            w, h = tree.winfo_width(), tree.winfo_height()
            tam = (w, h, self._modo_actual)
            if w < 40 or h < 40 or (tam == info["tam"] and not forzar):
                return
            p = TTK_PALETTES[self._modo_actual]
            img = FD.componer_tabla(info["clave"], w, h, p["bg_card"], self._modo_actual)
            if img is None:
                return
            from PIL import ImageTk
            nueva = ImageTk.PhotoImage(img, master=self)
            info["foto"].configure(width=w, height=h)
            info["foto"].tk.call(info["foto"], "copy", nueva)
            info["tam"] = tam
        except Exception:
            pass  # decorativo: nunca debe interrumpir la app

    def _refrescar_fondos_tablas(self) -> None:
        for tree, info in getattr(self, "_fondos_tabla", {}).items():
            self._estilo_fondo_tabla(info["estilo"])
            self._dibujar_fondo_tabla(tree, forzar=True)

    def _bg_fila(self, tree, color: str) -> str:
        """Color de fondo de las filas: transparente en las tablas con ilustración."""
        return "" if tree in getattr(self, "_fondos_tabla", {}) else color

    def _cambiar_pagina(self, valor: str) -> None:
        self._pag_registro.grid_forget()
        self._pag_calculos.grid_forget()
        self._pag_tablero.grid_forget()

        if "Registro" in valor:
            self._pag_registro.grid(row=0, column=0, sticky="nsew")
        elif "Cálculos" in valor:
            self._pag_calculos.grid(row=0, column=0, sticky="nsew")
            self._ejecutar_analisis_ui()
        elif "Tablero" in valor:
            self._pag_tablero.grid(row=0, column=0, sticky="nsew")
            self._ejecutar_analisis_ui()

    def _cambiar_tema(self) -> None:
        self._cerrar_entry_activo(guardar=True)
        if self._modo_actual == "dark":
            self._modo_actual = "light"
            ctk.set_appearance_mode("light")
            self._btn_tema.configure(text="🌙 Modo Oscuro")
        else:
            self._modo_actual = "dark"
            ctk.set_appearance_mode("dark")
            self._btn_tema.configure(text="☀️ Modo Claro")

        self._aplicar_estilos_ttk()
        self._refrescar_colores_captura()
        self._cargar_historial()
        self._ejecutar_analisis_ui()

    def _toggle_admin_mode(self) -> None:
        if not self._admin_activo:
            pin = simpledialog.askstring(
                "Modo Administrador",
                "Introduce el PIN de administrador:",
                show="*",
                parent=self,
            )
            if pin == ADMIN_PIN:
                self._admin_activo = True
                self._btn_admin.configure(
                    text="🔓 Admin Activo",
                    fg_color=UI_COLORS["accent_amber"],
                    text_color="#ffffff",
                )
                self._btn_editar.configure(state="normal")
                self._btn_eliminar.configure(state="normal")
                messagebox.showinfo(
                    "Modo Administrador",
                    "Modo administrador activado.\n"
                    "Ahora puedes editar o eliminar registros seleccionándolos o con clic derecho.",
                    parent=self,
                )
            elif pin is not None:
                messagebox.showerror("Acceso Denegado", "PIN de administrador incorrecto.", parent=self)
        else:
            self._admin_activo = False
            self._btn_admin.configure(
                text="🔒 Admin",
                fg_color=UI_COLORS["bg_secondary"],
                text_color=UI_COLORS["text_primary"],
            )
            self._btn_editar.configure(state="disabled")
            self._btn_eliminar.configure(state="disabled")
            messagebox.showinfo("Modo Administrador", "Modo administrador desactivado.", parent=self)

    # -----------------------------------------------------------------------
    # Página 1: Registro e Historial
    # -----------------------------------------------------------------------

    def _construir_cuerpo_registro(self, padre: ctk.CTkFrame) -> None:
        # Damos el mismo peso (y minsize amplio) al panel de registro para que se vea ancho y completo
        padre.grid_columnconfigure(0, weight=6, minsize=640)
        padre.grid_columnconfigure(1, weight=6, minsize=640)
        padre.grid_rowconfigure(0, weight=1)

        self._construir_panel_registro(padre)
        self._construir_panel_historial(padre)

    def _construir_panel_registro(self, padre: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(
            padre,
            fg_color=UI_COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        panel.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        # La fila 3 (tabla de captura) se expande dinámicamente para ver toda la sección sin cortar
        panel.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            panel,
            text="📥  Registro de Muestra por Lote",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=UI_COLORS["text_primary"],
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(14, 0), sticky="w")

        ctk.CTkFrame(panel, height=1, fg_color=UI_COLORS["border"]).grid(
            row=1, column=0, padx=18, pady=(8, 0), sticky="ew"
        )

        # Layout directo (sin ScrollableFrame intermedio que robe ancho), aprovechando el 100% del panel
        self._construir_campos_generales(panel)
        self._construir_tabla_captura(panel)
        self._construir_boton_guardar(panel)

    def _construir_campos_generales(self, padre) -> None:
        frame = ctk.CTkFrame(
            padre,
            fg_color=UI_COLORS["bg_card_light"],
            corner_radius=8,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        frame.grid(row=2, column=0, padx=16, pady=(10, 6), sticky="ew")
        frame.grid_columnconfigure((0, 1, 2), weight=1)

        # --- FILA 0 y 1: Lote (abarca 2 columnas) y Variedad (1 columna) ---
        ctk.CTkLabel(
            frame,
            text="Identificador del Lote / Excel",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, padx=(14, 6), pady=(10, 2), sticky="w")

        self._entry_lote = ctk.CTkEntry(
            frame,
            placeholder_text="Ej: Lote 16-Mayo / Criollo M1",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_lote.grid(row=1, column=0, padx=(14, 6), pady=(0, 8), sticky="ew")

        ctk.CTkLabel(
            frame,
            text="Variedad de Semilla",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=1, columnspan=2, padx=(6, 14), pady=(10, 2), sticky="w")

        # Ocupa dos columnas para que se lea el nombre técnico o científico completo
        self._combo_semilla = ctk.CTkComboBox(
            frame,
            values=TIPOS_SEMILLA,
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            button_color=UI_COLORS["border"],
            button_hover_color=UI_COLORS["accent_green"],
            dropdown_fg_color=UI_COLORS["bg_card"],
            dropdown_hover_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"],
            dropdown_text_color=UI_COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
            state="readonly",
        )
        self._combo_semilla.set(TIPOS_SEMILLA[0])
        self._combo_semilla.grid(row=1, column=1, columnspan=2, padx=(6, 14), pady=(0, 8), sticky="ew")

        # --- FILA 2 y 3: Fecha | Tipo de tratamiento | Tiempo de tratamiento ---
        ctk.CTkLabel(
            frame,
            text="Fecha de Registro",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=2, column=0, padx=(14, 6), pady=(4, 2), sticky="w")

        self._entry_fecha = ctk.CTkEntry(
            frame,
            placeholder_text="YYYY-MM-DD HH:MM",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_fecha.insert(0, formatear_fecha_actual())
        self._entry_fecha.grid(row=3, column=0, padx=(14, 6), pady=(0, 12), sticky="ew")

        ctk.CTkLabel(
            frame,
            text="Tiempo de Tratamiento",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=2, column=2, padx=(6, 14), pady=(4, 2), sticky="w")

        self._entry_tratamiento = ctk.CTkEntry(
            frame,
            placeholder_text="Ej: 45s, Control, T-1M",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_tratamiento.grid(row=3, column=2, padx=(6, 14), pady=(0, 12), sticky="ew")

        ctk.CTkLabel(
            frame,
            text="🌡️ Temperatura",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=4, column=0, padx=(14, 6), pady=(0, 2), sticky="w")

        # Campo abierto: número (35 -> 35 °C), rango o cualquier texto; vacío = Ambiente
        self._entry_temperatura = ctk.CTkEntry(
            frame,
            placeholder_text="Libre: 35, 35-40 °C… (vacío = Ambiente)",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_temperatura.grid(row=5, column=0, padx=(14, 6), pady=(0, 12), sticky="ew")

        # --- FILA 4 y 5: (Temperatura) | Tipo de gas | Flujo (L/min) ---
        # --- FILA 6 y 7: Orientación | Morfología | Cantidad de semillas ---
        def etiqueta(texto, fila, col):
            ctk.CTkLabel(frame, text=texto, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                         text_color=UI_COLORS["text_secondary"], anchor="w").grid(
                row=fila, column=col, padx=(14 if col == 0 else 6, 14 if col == 2 else 6), pady=(0, 2), sticky="w")

        def combo(valores, fila, col):
            c = ctk.CTkComboBox(
                frame, values=valores, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"],
                button_color=UI_COLORS["border"], button_hover_color=UI_COLORS["accent_green"],
                dropdown_fg_color=UI_COLORS["bg_card"], dropdown_hover_color=UI_COLORS["bg_secondary"],
                text_color=UI_COLORS["text_primary"], dropdown_text_color=UI_COLORS["text_primary"],
                font=ctk.CTkFont(family="Segoe UI", size=12), state="readonly")
            c.set(valores[0])
            c.grid(row=fila, column=col, padx=(14 if col == 0 else 6, 14 if col == 2 else 6), pady=(0, 12), sticky="ew")
            return c

        def entrada(placeholder, fila, col, valor=""):
            e = ctk.CTkEntry(frame, placeholder_text=placeholder, fg_color=UI_COLORS["input_bg"],
                             border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"],
                             placeholder_text_color=UI_COLORS["text_muted"], font=ctk.CTkFont(family="Segoe UI", size=12))
            if valor:
                e.insert(0, valor)
            e.grid(row=fila, column=col, padx=(14 if col == 0 else 6, 14 if col == 2 else 6), pady=(0, 12), sticky="ew")
            return e

        etiqueta("⚡ Tipo de tratamiento", 2, 1)
        self._combo_tipo_trat = combo(TIPOS_TRATAMIENTO, 3, 1)
        etiqueta("💨 Tipo de gas", 4, 1)
        self._combo_gas = combo(TIPOS_GAS, 5, 1)
        etiqueta("Flujo (L/min)", 4, 2)
        self._entry_flujo = entrada("Ej: 2 (vacío = sin flujo)", 5, 2)
        etiqueta("Orientación de la semilla", 6, 0)
        self._combo_orientacion = combo(ORIENTACIONES, 7, 0)
        etiqueta("Morfología de la semilla", 6, 1)
        self._combo_morfologia = combo(MORFOLOGIAS, 7, 1)
        etiqueta("Cantidad de semillas en el lote", 6, 2)
        caja = ctk.CTkFrame(frame, fg_color="transparent")
        caja.grid(row=7, column=2, padx=(6, 14), pady=(0, 12), sticky="ew")
        caja.grid_columnconfigure(0, weight=1)
        self._entry_cantidad = ctk.CTkEntry(caja, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"],
                                            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(family="Segoe UI", size=12))
        self._entry_cantidad.insert(0, str(NUM_MUESTRAS))
        self._entry_cantidad.grid(row=0, column=0, sticky="ew")
        self._entry_cantidad.bind("<Return>", lambda e: self._aplicar_cantidad_semillas())
        ctk.CTkButton(caja, text="Aplicar", width=64, fg_color=UI_COLORS["accent_indigo"], text_color="#ffffff",
                      font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._aplicar_cantidad_semillas).grid(row=0, column=1, padx=(6, 0))

    def _aplicar_cantidad_semillas(self) -> None:
        """Cambia el número de filas de la tabla de captura conservando lo ya capturado."""
        if not self._cerrar_entry_activo(guardar=True):
            return
        ok, n, msg = validar_cantidad_semillas(self._entry_cantidad.get())
        if not ok:
            messagebox.showerror("Cantidad inválida", msg, parent=self)
            return
        if n == self._num_muestras:
            return
        perdidas = [i + 1 for i, f in enumerate(self._filas_captura[n:], start=n)
                    if any(f[k] is not None for k in ("raiz", "tallo", "hojas"))]
        if perdidas and not messagebox.askyesno(
                "Reducir la tabla",
                f"Las muestras {perdidas[0]} a {perdidas[-1]} ya tienen datos y se borrarán. ¿Continuar?", parent=self):
            self._entry_cantidad.delete(0, tk.END)
            self._entry_cantidad.insert(0, str(self._num_muestras))
            return
        previas = self._filas_captura[:n]
        self._num_muestras = n
        self._inicializar_tabla_captura()
        for i, f in enumerate(previas):
            self._filas_captura[i] = f
            self._actualizar_fila_estado(i)

    def _construir_campo(self, padre, etiqueta: str, fila: int, col: int, placeholder: str = "", valor_inicial: str = "") -> None:
        ctk.CTkLabel(
            padre,
            text=etiqueta,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=fila, column=col, padx=(14 if col == 0 else 8, 8 if col == 0 else 14), pady=(12, 4), sticky="w")

        entry = ctk.CTkEntry(
            padre,
            placeholder_text=placeholder,
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        if valor_inicial:
            entry.insert(0, valor_inicial)
        entry.grid(row=fila + 1, column=col, padx=(14 if col == 0 else 8, 8 if col == 0 else 14), pady=(0, 12), sticky="ew")
        self._ultimo_entry = entry

    def _construir_tabla_captura(self, padre) -> None:
        contenedor_tabla = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor_tabla.grid(row=3, column=0, padx=16, pady=(2, 6), sticky="nsew")
        contenedor_tabla.grid_columnconfigure(0, weight=1)
        contenedor_tabla.grid_rowconfigure(1, weight=1)

        self._lbl_tabla_captura = ctk.CTkLabel(
            contenedor_tabla,
            text=f"📊  Tabla de Captura — {self._num_muestras} Muestras (Doble clic o Enter para editar)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        )
        self._lbl_tabla_captura.grid(row=0, column=0, pady=(0, 4), sticky="w")

        tabla_frame = ctk.CTkFrame(
            contenedor_tabla,
            fg_color=UI_COLORS["bg_card"],
            corner_radius=8,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        tabla_frame.grid(row=1, column=0, sticky="nsew")
        tabla_frame.grid_columnconfigure(0, weight=1)
        tabla_frame.grid_rowconfigure(0, weight=1)

        columnas = ("no", "raiz", "tallo", "hojas", "estado")
        self._tree_captura = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings",
            style="Captura.Treeview",
            height=12,
            selectmode="browse",
        )

        self._tree_captura.heading("no",     text="No.",           anchor="center")
        self._tree_captura.heading("raiz",   text="Raíz (mm)",     anchor="center")
        self._tree_captura.heading("tallo",  text="Tallo (mm)",    anchor="center")
        self._tree_captura.heading("hojas",  text="Cant. Hojas",   anchor="center")
        self._tree_captura.heading("estado", text="Germinación",   anchor="center")

        self._tree_captura.column("no",     width=55,  minwidth=45, anchor="center", stretch=False)
        self._tree_captura.column("raiz",   width=130, minwidth=100, anchor="center", stretch=True)
        self._tree_captura.column("tallo",  width=130, minwidth=100, anchor="center", stretch=True)
        self._tree_captura.column("hojas",  width=115, minwidth=95, anchor="center", stretch=True)
        self._tree_captura.column("estado", width=145, minwidth=120, anchor="center", stretch=True)

        vsb_captura = ttk.Scrollbar(
            tabla_frame, orient="vertical",
            command=self._tree_captura.yview,
            style="Dark.Vertical.TScrollbar",
        )
        self._tree_captura.configure(yscrollcommand=vsb_captura.set)
        self._tree_captura.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=2)
        vsb_captura.grid(row=0, column=1, sticky="ns", pady=2)

        self._aplicar_estilos_ttk()

        self._tree_captura.bind("<Double-1>", self._on_doble_click_captura)
        self._tree_captura.bind("<Return>",   self._on_enter_captura)
        self._tree_captura.bind("<Up>",       self._on_arrow_captura)
        self._tree_captura.bind("<Down>",     self._on_arrow_captura)
        self._tree_captura.bind("<Left>",     self._on_arrow_captura)
        self._tree_captura.bind("<Right>",    self._on_arrow_captura)
        self._tree_captura.bind("<<TreeviewSelect>>", self._on_seleccion_captura)

    def _inicializar_tabla_captura(self) -> None:
        self._filas_captura = []
        for item in self._tree_captura.get_children():
            self._tree_captura.delete(item)

        if hasattr(self, "_lbl_tabla_captura"):
            self._lbl_tabla_captura.configure(
                text=f"📊  Tabla de Captura — {self._num_muestras} Muestras (Doble clic o Enter para editar)")
        for i in range(1, self._num_muestras + 1):
            self._filas_captura.append({"raiz": None, "tallo": None, "hojas": None})
            tag = "row_alt" if i % 2 == 0 else "row_norm"
            self._tree_captura.insert("", "end", iid=str(i), values=(i, "", "", "", ""), tags=(tag,))

    def _refrescar_colores_captura(self) -> None:
        for idx in range(self._num_muestras):
            self._actualizar_fila_estado(idx)

    def _actualizar_fila_estado(self, fila_idx: int) -> None:
        fila = self._filas_captura[fila_idx]
        raiz, tallo, hojas = fila["raiz"], fila["tallo"], fila["hojas"]
        num = fila_idx + 1
        iid = str(num)
        es_par = (num % 2 == 0)

        if raiz is None or tallo is None:
            estado_txt = ""
            tag = "row_alt" if es_par else "row_norm"
        elif calcular_germino(raiz, tallo) == 1:
            estado_txt = "✓  Germinó"
            tag = "alt_si" if es_par else "germino_si"
        else:
            estado_txt = "○  No germinó"
            tag = "alt_no" if es_par else "germino_no"

        self._tree_captura.item(
            iid,
            values=(
                num,
                "" if raiz is None else formatear_numero(raiz),
                "" if tallo is None else formatear_numero(tallo),
                "" if hojas is None else str(hojas),
                estado_txt,
            ),
            tags=(tag,),
        )

    def _on_seleccion_captura(self, event) -> None:
        if self._celda_activa:
            sel = self._tree_captura.selection()
            if sel and int(sel[0]) - 1 == self._celda_activa[0]:
                return
        self._cerrar_entry_activo(guardar=True)

    def _on_doble_click_captura(self, event) -> None:
        region = self._tree_captura.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self._tree_captura.identify_column(event.x)
        item_id = self._tree_captura.identify_row(event.y)
        if not item_id:
            return
        col_num = int(col_id.replace("#", ""))
        if col_num in (2, 3, 4):
            self._abrir_entry_inline(item_id, col_num)

    def _on_enter_captura(self, event) -> None:
        sel = self._tree_captura.selection()
        if not sel:
            return
        item_id = sel[0]
        if self._entry_activo and self._celda_activa:
            fila_idx, col_num = self._celda_activa
            if not self._cerrar_entry_activo(guardar=True):
                return
            siguiente = fila_idx + 1
            if siguiente < self._num_muestras:
                nuevo_iid = str(siguiente + 1)
                self._tree_captura.selection_set(nuevo_iid)
                self._tree_captura.see(nuevo_iid)
                self._abrir_entry_inline(nuevo_iid, col_num)
        else:
            self._abrir_entry_inline(item_id, 2)

    def _on_arrow_captura(self, event) -> None:
        sel = self._tree_captura.selection()
        if not sel:
            return
        item_id = sel[0]
        fila_idx = int(item_id) - 1
        col_actual = self._celda_activa[1] if self._celda_activa else 2

        nueva_fila = fila_idx
        nueva_col = col_actual

        if event.keysym == "Up" and fila_idx > 0:
            nueva_fila = fila_idx - 1
        elif event.keysym == "Down" and fila_idx < self._num_muestras - 1:
            nueva_fila = fila_idx + 1
        elif event.keysym == "Left" and col_actual > 2:
            nueva_col = col_actual - 1
        elif event.keysym == "Right" and col_actual < 4:
            nueva_col = col_actual + 1

        if not self._cerrar_entry_activo(guardar=True):
            return

        nuevo_iid = str(nueva_fila + 1)
        self._tree_captura.selection_set(nuevo_iid)
        self._tree_captura.see(nuevo_iid)
        self._abrir_entry_inline(nuevo_iid, nueva_col)

    def _abrir_entry_inline(self, item_id: str, col_num: int) -> None:
        if not self._cerrar_entry_activo(guardar=True):
            return

        p = TTK_PALETTES[self._modo_actual]
        fila_idx = int(item_id) - 1
        valores = self._tree_captura.item(item_id, "values")
        valor_actual = str(valores[col_num - 1])

        bbox = self._tree_captura.bbox(item_id, f"#{col_num}")
        if not bbox:
            return
        x, y, w, h = bbox

        entry = tk.Entry(
            self._tree_captura,
            font=("Segoe UI", 10),
            bg=p["input_bg"],
            fg=p["text_primary"],
            insertbackground=p["text_primary"],
            relief="flat",
            bd=1,
            highlightthickness=2,
            highlightbackground=p["accent_green"],
            highlightcolor=p["accent_green"],
            justify="center",
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, valor_actual.replace(",", "."))
        entry.select_range(0, tk.END)
        entry.focus_set()

        self._entry_activo = entry
        self._celda_activa = (fila_idx, col_num)

        entry.bind("<Return>",   lambda e: self._on_enter_captura(e))
        entry.bind("<Escape>",   lambda e: self._cerrar_entry_activo(guardar=False))
        entry.bind("<FocusOut>", lambda e: self._cerrar_entry_activo(guardar=True))
        entry.bind("<Tab>",      lambda e: self._tab_captura())

    def _cerrar_entry_activo(self, guardar: bool = True) -> bool:
        if self._entry_activo is None:
            return True
        if self._mostrando_error:
            return False

        if guardar:
            texto = self._entry_activo.get()
            fila_idx, col_num = self._celda_activa
            if col_num == 4:
                try:
                    valor = parsear_cantidad_hojas(texto)
                except ValueError:
                    self._avisar_valor_invalido(
                        fila_idx,
                        f"La cantidad de hojas de la muestra {fila_idx + 1} no es válida.\n"
                        "Debe ser un entero mayor o igual a 0.",
                    )
                    return False
                self._filas_captura[fila_idx]["hojas"] = valor
            else:
                campo = "raíz" if col_num == 2 else "tallo"
                valido, valor, _ = validar_valor_numerico(texto, campo)
                if not valido:
                    self._avisar_valor_invalido(
                        fila_idx,
                        f"La longitud de {campo} de la muestra {fila_idx + 1} no es válida.\n"
                        "Debe ser un número mayor o igual a 0.",
                    )
                    return False
                if col_num == 2:
                    self._filas_captura[fila_idx]["raiz"] = valor
                else:
                    self._filas_captura[fila_idx]["tallo"] = valor
            self._actualizar_fila_estado(fila_idx)

        try:
            self._entry_activo.destroy()
        except Exception:
            pass
        self._entry_activo = None
        self._celda_activa = None
        return True

    def _avisar_valor_invalido(self, fila_idx: int, mensaje: str) -> None:
        entry = self._entry_activo
        self._mostrando_error = True
        try:
            messagebox.showerror("Valor inválido", mensaje, parent=self)
        finally:
            self._mostrando_error = False

        if entry is not None and entry.winfo_exists():
            self._tree_captura.selection_set(str(fila_idx + 1))
            entry.focus_set()
            entry.select_range(0, tk.END)

    def _tab_captura(self) -> str:
        if not self._celda_activa:
            return "break"
        fila_idx, col_num = self._celda_activa
        if not self._cerrar_entry_activo(guardar=True):
            return "break"

        if col_num < 4:
            nueva_col = col_num + 1
            nueva_fila = fila_idx
        else:
            nueva_col = 2
            nueva_fila = fila_idx + 1

        if nueva_fila < self._num_muestras:
            nuevo_iid = str(nueva_fila + 1)
            self._tree_captura.selection_set(nuevo_iid)
            self._tree_captura.see(nuevo_iid)
            self._abrir_entry_inline(nuevo_iid, nueva_col)
        return "break"

    def _construir_boton_guardar(self, padre) -> None:
        self._btn_guardar = ctk.CTkButton(
            padre,
            text="💾   Guardar Registro",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=UI_COLORS["accent_green"],
            hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff",
            height=40,
            corner_radius=8,
            command=self._guardar_lote,
        )
        self._btn_guardar.grid(row=4, column=0, padx=16, pady=(4, 14), sticky="ew")

    def _guardar_lote(self) -> None:
        if not self._cerrar_entry_activo(guardar=True):
            return

        nombre_lote    = self._entry_lote.get().strip()
        tipo_semilla   = self._combo_semilla.get()
        fecha_registro = self._entry_fecha.get().strip()
        tiempo_trat    = self._entry_tratamiento.get().strip()
        temp_raw       = self._entry_temperatura.get().strip()

        valido, msg_error = validar_campos_lote(nombre_lote, tipo_semilla, fecha_registro, tiempo_trat)
        if not valido:
            messagebox.showerror("Campo obligatorio", msg_error, parent=self)
            return

        ok_temp, temp_fmt, msg_temp = validar_temperatura(temp_raw)
        if not ok_temp:
            messagebox.showerror("Temperatura inválida", msg_temp, parent=self)
            return

        ok_flujo, flujo, msg_flujo = validar_flujo(self._entry_flujo.get())
        if not ok_flujo:
            messagebox.showerror("Flujo inválido", msg_flujo, parent=self)
            return
        gas = self._combo_gas.get()
        if gas == "Sin gas" and flujo:
            if not messagebox.askyesno("Revisa el gas", "Pusiste un flujo pero el gas es 'Sin gas'. ¿Guardar así?",
                                       parent=self):
                return

        mediciones = []
        for i, fila in enumerate(self._filas_captura, start=1):
            raiz, tallo, hojas = fila["raiz"], fila["tallo"], fila["hojas"]
            if raiz is None and tallo is None and hojas is None:
                continue
            if raiz is None or tallo is None or hojas is None:
                messagebox.showerror(
                    "Muestra incompleta",
                    f"La muestra {i} está incompleta. Debes capturar raíz, tallo y cantidad de hojas.",
                    parent=self,
                )
                return
            mediciones.append({
                "num_muestra":    i,
                "longitud_raiz":  raiz,
                "longitud_tallo": tallo,
                "cantidad_hojas": hojas,
            })

        if not mediciones:
            messagebox.showerror("Sin mediciones", "Debes capturar al menos una medición.", parent=self)
            return

        datos = DatosLote(
            nombre_lote=nombre_lote,
            tipo_semilla=tipo_semilla,
            fecha_registro=fecha_registro,
            tiempo_tratamiento=tiempo_trat,
            temperatura=temp_fmt,
            mediciones=mediciones,
            gas=gas,
            flujo=flujo,
            orientacion=self._combo_orientacion.get(),
            morfologia=self._combo_morfologia.get(),
            cantidad_semillas=self._num_muestras,
            tipo_tratamiento=self._combo_tipo_trat.get(),
        )

        self._btn_guardar.configure(state="disabled", text="⏳   Guardando...")
        exito, mensaje = self.db.guardar_lote(datos)
        self._btn_guardar.configure(state="normal", text="💾   Guardar Registro")

        if exito:
            messagebox.showinfo("✅ Registro Guardado", mensaje, parent=self)
            self._limpiar_formulario()
            self._cargar_historial()
        else:
            messagebox.showerror("❌ Error al Guardar", mensaje, parent=self)

    def _limpiar_formulario(self) -> None:
        self._entry_lote.delete(0, tk.END)
        self._combo_semilla.set(TIPOS_SEMILLA[0])
        self._entry_fecha.delete(0, tk.END)
        self._entry_fecha.insert(0, formatear_fecha_actual())
        self._entry_tratamiento.delete(0, tk.END)
        self._entry_temperatura.delete(0, tk.END)
        # Tipo de tratamiento, gas, flujo, orientación, morfología y cantidad se conservan: suelen repetirse entre lotes.
        self._inicializar_tabla_captura()

    # -----------------------------------------------------------------------
    # Panel Derecho: Historial, Exportación y Herramientas Admin
    # -----------------------------------------------------------------------

    def _construir_panel_historial(self, padre: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(
            padre,
            fg_color=UI_COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        panel.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        
        # Configurar explícitamente el peso de cada fila para evitar colapsos
        panel.grid_rowconfigure(0, weight=0)  # Título y botones superiores
        panel.grid_rowconfigure(1, weight=0)  # Línea divisoria
        panel.grid_rowconfigure(2, weight=0)  # Filtros de fecha y botón de gráficos
        panel.grid_rowconfigure(3, weight=0)  # Filtros avanzados (Variedad, Tratamiento, Germinación)
        panel.grid_rowconfigure(4, weight=1)  # Tabla del historial (se expande)

        # Fila 0: Barra superior del historial (Título y Botones de acción)
        top_hist_frame = ctk.CTkFrame(panel, fg_color="transparent")
        top_hist_frame.grid(row=0, column=0, padx=16, pady=(14, 0), sticky="ew")
        top_hist_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_hist_frame,
            text="📋  Historial Consolidado",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=UI_COLORS["text_primary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        acciones_frame = ctk.CTkFrame(top_hist_frame, fg_color="transparent")
        acciones_frame.grid(row=0, column=1, sticky="e")

        # Botones de Exportación
        ctk.CTkButton(
            acciones_frame, text="📊 Exportar Excel", width=110, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["accent_green"], hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff", command=self._exportar_excel,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            acciones_frame, text="📄 Exportar PDF", width=105, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["accent_indigo"], hover_color="#4338ca",
            text_color="#ffffff", command=self._exportar_pdf,
        ).pack(side="left", padx=(0, 12))

        # Botones de Administración (Editar y Eliminar)
        self._btn_editar = ctk.CTkButton(
            acciones_frame, text="✏️ Editar", width=80, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["accent_amber"], hover_color="#b45309",
            text_color="#ffffff", state="disabled", command=self._abrir_modal_edicion,
        )
        self._btn_editar.pack(side="left", padx=(0, 6))

        self._btn_eliminar = ctk.CTkButton(
            acciones_frame, text="🗑️ Borrar", width=80, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=UI_COLORS["accent_red"], hover_color="#b91c1c",
            text_color="#ffffff", state="disabled", command=self._confirmar_borrado_seleccion,
        )
        self._btn_eliminar.pack(side="left")

        # Fila 1: Línea divisoria
        ctk.CTkFrame(panel, height=1, fg_color=UI_COLORS["border"]).grid(
            row=1, column=0, padx=16, pady=(10, 0), sticky="ew"
        )

        # Fila 2: Filtros de Rango de Fechas y Botón de Gráficos del Rango
        self._construir_filtros_fecha(panel)

        # Fila 3: Filtros de Búsqueda y Selección Múltiple (Variedades, Tratamientos, Germinación)
        self._construir_filtros(panel)

        # Fila 4: Tabla / Treeview Consolidado
        self._construir_treeview_historial(panel)

    def _construir_filtros_fecha(self, padre) -> None:
        """Fila 2: Construye los selectores de fecha con calendario emergente y el botón de gráficos."""
        fecha_frame = ctk.CTkFrame(padre, fg_color="transparent")
        fecha_frame.grid(row=2, column=0, padx=16, pady=(8, 0), sticky="ew")

        ctk.CTkLabel(
            fecha_frame, text="📅 Desde:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"]
        ).pack(side="left", padx=(0, 4))

        # Variables de fecha ("Todas" o formato YYYY-MM-DD)
        self._var_hist_desde = tk.StringVar(value="Todas")
        self._btn_desde = ctk.CTkButton(
            fecha_frame, textvariable=self._var_hist_desde, width=110, height=28,
            fg_color=UI_COLORS["input_bg"], text_color=UI_COLORS["text_primary"],
            border_width=1, border_color=UI_COLORS["border"],
            hover_color=UI_COLORS["border"],
            command=lambda: self._abrir_calendario_modal(self._var_hist_desde)
        )
        self._btn_desde.pack(side="left", padx=4)

        ctk.CTkLabel(
            fecha_frame, text="Hasta:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"]
        ).pack(side="left", padx=(10, 4))

        self._var_hist_hasta = tk.StringVar(value="Todas")
        self._btn_hasta = ctk.CTkButton(
            fecha_frame, textvariable=self._var_hist_hasta, width=110, height=28,
            fg_color=UI_COLORS["input_bg"], text_color=UI_COLORS["text_primary"],
            border_width=1, border_color=UI_COLORS["border"],
            hover_color=UI_COLORS["border"],
            command=lambda: self._abrir_calendario_modal(self._var_hist_hasta)
        )
        self._btn_hasta.pack(side="left", padx=4)

        # Botón para limpiar o restablecer fechas
        ctk.CTkButton(
            fecha_frame, text="🔄 Ver Todo", width=85, height=28,
            fg_color=UI_COLORS["bg_secondary"], text_color=UI_COLORS["text_primary"],
            hover_color=UI_COLORS["border"], font=ctk.CTkFont(size=10, weight="bold"),
            command=self._restablecer_fechas
        ).pack(side="left", padx=(8, 4))

        # Botón para abrir la ventana nueva con gráficos del rango
        ctk.CTkButton(
            fecha_frame, text="📊 Ver Gráficos del Rango",
            fg_color=UI_COLORS["accent_indigo"], hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
            command=self._abrir_ventana_graficos_historial,
            width=170
        ).pack(side="right", padx=4)

        ctk.CTkButton(
            fecha_frame, text="📥 Importar",
            fg_color=UI_COLORS["accent_green"], hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
            command=self._importar_excel_desde_historial,
            width=95
        ).pack(side="right", padx=4)

    def _abrir_calendario_modal(self, variable_destino: tk.StringVar) -> None:
        """Abre una ventana emergente nativa con selectores de Fecha (Año, Mes, Día)."""
        top_cal = ctk.CTkToplevel(self)
        top_cal.title("Seleccionar Fecha")
        top_cal.geometry("300x260")
        top_cal.grab_set()
        centrar_ventana(top_cal, 300, 260)

        ctk.CTkLabel(
            top_cal, text="📅 Selecciona la Fecha",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(15, 10))

        frame_selectors = ctk.CTkFrame(top_cal, fg_color="transparent")
        frame_selectors.pack(padx=20, pady=10, fill="x")
        frame_selectors.grid_columnconfigure((0, 1, 2), weight=1)

        # Listas de opciones (Años 2024-2027, Meses 01-12, Días 01-31)
        anos = [str(y) for y in range(2024, 2028)]
        meses = [f"{m:02d}" for m in range(1, 13)]
        dias = [f"{d:02d}" for d in range(1, 32)]

        # Valores por defecto actuales (octubre 2026 o el actual)
        val_actual = variable_destino.get()
        if val_actual and val_actual != "Todas" and len(val_actual) >= 10:
            def_anio, def_mes, def_dia = val_actual[:4], val_actual[5:7], val_actual[8:10]
        else:
            def_anio, def_mes, def_dia = "2026", "10", "01"

        # ComboBox Año
        ctk.CTkLabel(frame_selectors, text="Año", font=ctk.CTkFont(size=10)).grid(row=0, column=0, padx=2)
        combo_anio = ctk.CTkComboBox(frame_selectors, values=anos, width=80, state="readonly")
        combo_anio.set(def_anio if def_anio in anos else "2026")
        combo_anio.grid(row=1, column=0, padx=2, pady=4)

        # ComboBox Mes
        ctk.CTkLabel(frame_selectors, text="Mes", font=ctk.CTkFont(size=10)).grid(row=0, column=1, padx=2)
        combo_mes = ctk.CTkComboBox(frame_selectors, values=meses, width=70, state="readonly")
        combo_mes.set(def_mes if def_mes in meses else "10")
        combo_mes.grid(row=1, column=1, padx=2, pady=4)

        # ComboBox Día
        ctk.CTkLabel(frame_selectors, text="Día", font=ctk.CTkFont(size=10)).grid(row=0, column=2, padx=2)
        combo_dia = ctk.CTkComboBox(frame_selectors, values=dias, width=70, state="readonly")
        combo_dia.set(def_dia if def_dia in dias else "01")
        combo_dia.grid(row=1, column=2, padx=2, pady=4)

        def seleccionar():
            fecha_str = f"{combo_anio.get()}-{combo_mes.get()}-{combo_dia.get()}"
            variable_destino.set(fecha_str)
            top_cal.destroy()
            self._cargar_historial()

        btn_frame = ctk.CTkFrame(top_cal, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(15, 15))
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_frame, text="Limpiar / Todas", fg_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"],
            command=lambda: [variable_destino.set("Todas"), top_cal.destroy(), self._cargar_historial()]
        ).grid(row=0, column=0, padx=(0, 5), sticky="ew")

        ctk.CTkButton(
            btn_frame, text="Aceptar", fg_color=UI_COLORS["accent_green"],
            hover_color=UI_COLORS["accent_hover"], text_color="#ffffff",
            command=seleccionar
        ).grid(row=0, column=1, padx=(5, 0), sticky="ew")

    def _restablecer_fechas(self) -> None:
        """Restablece los filtros de fecha a 'Todas'."""
        if hasattr(self, "_var_hist_desde"):
            self._var_hist_desde.set("Todas")
        if hasattr(self, "_var_hist_hasta"):
            self._var_hist_hasta.set("Todas")
        self._cargar_historial()

    def _construir_filtros(self, padre) -> None:
        """Fila 3: Filtros avanzados con soporte para selección múltiple."""
        filtros_frame = ctk.CTkFrame(padre, fg_color="transparent")
        filtros_frame.grid(row=3, column=0, padx=16, pady=(6, 8), sticky="ew")
        filtros_frame.grid_columnconfigure(0, weight=2)  # Buscador
        filtros_frame.grid_columnconfigure(1, weight=1)  # Botón Variedades
        filtros_frame.grid_columnconfigure(2, weight=1)  # Botón Tratamientos
        filtros_frame.grid_columnconfigure(3, weight=1)  # Botón Germinación
        filtros_frame.grid_columnconfigure(4, weight=1)  # Contador

        # 1. Buscador de texto libre
        self._var_busqueda = tk.StringVar()
        self._var_busqueda.trace_add("write", lambda *_: self._cargar_historial())

        ctk.CTkEntry(
            filtros_frame,
            placeholder_text="🔍  Buscar lote...",
            textvariable=self._var_busqueda,
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        # Inicializar conjuntos de selección múltiple con todas las opciones por defecto
        # None = sin filtro (se muestran todos, también variedades o tiempos fuera de la lista)
        for atributo in ("_sel_variedades", "_sel_tratamientos", "_sel_germino"):
            if not hasattr(self, atributo):
                setattr(self, atributo, None)

        # 2. Botón desplegable para Variedades (Pasando TIPOS_SEMILLA completos)
        ctk.CTkButton(
            filtros_frame, text="🌱 Variedades...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Variedades", "_sel_variedades", self._opciones_variedades())
        ).grid(row=0, column=1, padx=(0, 6), sticky="ew")

        # 3. Botón desplegable para Tratamientos (Pasando la lista completa de tratamientos)
        ctk.CTkButton(
            filtros_frame, text="⏱️ Tratamientos...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Tratamientos", "_sel_tratamientos", self._opciones_tratamientos())
        ).grid(row=0, column=2, padx=(0, 6), sticky="ew")

        # 4. Botón desplegable para Estado de Germinación
        ctk.CTkButton(
            filtros_frame, text="✓ Germinación...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Germinación", "_sel_germino", [("1", "Germinó"), ("0", "No germinó")])
        ).grid(row=0, column=3, padx=(0, 8), sticky="ew")

        # 5. Contador de registros
        self._lbl_total = ctk.CTkLabel(
            filtros_frame,
            text="0 registros",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="e",
        )
        self._lbl_total.grid(row=0, column=4, sticky="e")

    @staticmethod
    def _tratamiento_normalizado(texto) -> str:
        """'45', '45 s', '45 seg' -> '45s'; '2 min' -> '120s'; 'control' -> 'Control'; lo demás tal cual."""
        t = str(texto or "").strip()
        if not t or "control" in t.lower() or t in ("0", "0s"):
            return "Control" if t else "Sin definir"
        m = re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*(s|seg|segs|segundos?|min|mins|minutos?)?", t, re.IGNORECASE)
        if not m:
            return t
        v = float(m.group(1).replace(",", "."))
        if (m.group(2) or "").lower().startswith("min"):
            v *= 60
        return f"{v:g}s"

    def _opciones_variedades(self) -> list:
        """Variedades del catálogo más las que aparezcan en la base con otro nombre."""
        opciones = {AN.clave_variedad(t): t for t in TIPOS_SEMILLA}
        for row in getattr(self, "_historial_crudo", []):
            opciones.setdefault(AN.clave_variedad(row.get("tipo_semilla")), str(row.get("tipo_semilla") or "Sin especificar"))
        return list(opciones.items())

    def _opciones_tratamientos(self) -> list:
        vals = set(TRATAMIENTOS_BASE) | {self._tratamiento_normalizado(r.get("tiempo_tratamiento"))
                                         for r in getattr(self, "_historial_crudo", [])}
        def orden(t):
            m = re.fullmatch(r"(\d+(?:\.\d+)?)s", t)
            return (0, 0, t) if t == "Control" else ((1, float(m.group(1)), t) if m else (2, 0, t))
        return sorted(vals, key=orden)

    def _abrir_selector_multiple(self, titulo: str, atributo: str, opciones_disponibles: list) -> None:
        """Ventana modal con Checkboxes para selección múltiple de filtros.
        La selección se guarda en el atributo indicado; si quedan todas marcadas, se guarda None (sin filtro)."""
        conjunto_ref = getattr(self, atributo, None)
        top = ctk.CTkToplevel(self)
        top.title(f"Seleccionar {titulo}")
        top.geometry("440x400")
        top.grab_set()

        ctk.CTkLabel(
            top, text=f"Seleccione {titulo} a mostrar:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(pady=(15, 10))

        frame_chk = ctk.CTkScrollableFrame(top, width=300, height=220)
        frame_chk.pack(padx=15, pady=5, fill="both", expand=True)

        vars_dict = {}
        for op in opciones_disponibles:
            val_key = op[0] if isinstance(op, tuple) else op
            val_lbl = op[1] if isinstance(op, tuple) else op
            
            var = tk.BooleanVar(value=(conjunto_ref is None or val_key in conjunto_ref))
            vars_dict[val_key] = var
            
            chk = ctk.CTkCheckBox(frame_chk, text=val_lbl, variable=var)
            chk.pack(anchor="w", pady=4, padx=5)

        def aplicar_seleccion():
            elegidos = {str(k) for k, v in vars_dict.items() if v.get()}
            setattr(self, atributo, None if len(elegidos) == len(vars_dict) else elegidos)
            top.destroy()
            self._cargar_historial()

        ctk.CTkButton(
            top, text="Aplicar Filtros", fg_color=UI_COLORS["accent_green"],
            hover_color=UI_COLORS["accent_hover"], text_color="#ffffff",
            font=ctk.CTkFont(size=11, weight="bold"), command=aplicar_seleccion
        ).pack(pady=15)

    def _construir_treeview_historial(self, padre) -> None:
        tree_frame = ctk.CTkFrame(
            padre,
            fg_color=UI_COLORS["bg_card"],
            corner_radius=8,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        # ⚠️ CORRECCIÓN: Cambiado de row=3 a row=4 para ubicarse debajo de los filtros avanzados
        tree_frame.grid(row=4, column=0, padx=16, pady=(0, 16), sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        columnas = ("fecha", "lote", "semilla", "muestra", "tipo_trat", "tratamiento", "temperatura", "gas", "flujo",
                    "orientacion", "morfologia", "hojas", "raiz", "tallo", "resultado")
        
        self._tree_historial = ttk.Treeview(
            tree_frame,
            columns=columnas,
            show="headings",
            style="BioGrow.Treeview",
            selectmode="browse",
        )

        # Incluimos "fecha" en el diccionario de encabezados
        encabezados = {
            "fecha":       ("Fecha",          110, "center"),
            "lote":        ("Lote Origen",    120, "w"),
            "semilla":     ("Tipo Semilla",   240, "w"),  # cabe el nombre técnico o científico
            "muestra":     ("No.",             42, "center"),
            "tipo_trat":   ("Tipo trat.",      80, "center"),
            "tratamiento": ("Tratamiento",     85, "center"),
            "temperatura": ("Temperatura",     90, "center"),
            "gas":         ("Gas",             72, "center"),
            "flujo":       ("Flujo (L/min)",   80, "center"),
            "orientacion": ("Orientación",     82, "center"),
            "morfologia":  ("Morfología",      72, "center"),
            "hojas":       ("Cant. Hojas",     75, "center"),
            "raiz":        ("Raíz (mm)",       72, "center"),
            "tallo":       ("Tallo (mm)",      72, "center"),
            "resultado":   ("Resultado",       95, "center"),
        }
        
        for col, (texto, ancho, anchor) in encabezados.items():
            self._tree_historial.heading(col, text=texto, anchor=anchor)
            self._tree_historial.column(col, width=ancho, minwidth=45, anchor=anchor, stretch=(col == "lote"))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree_historial.yview, style="Dark.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree_historial.xview, style="Dark.Horizontal.TScrollbar")
        self._tree_historial.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree_historial.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(2, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(2, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(2, 0))

        self._activar_fondo_tabla(self._tree_historial, "registro", "BioGrow.Treeview")
        self._aplicar_estilos_ttk()

        # Menú contextual de clic derecho (Edición y Eliminación)
        self._menu_contextual = tk.Menu(self, tearoff=0)
        self._menu_contextual.add_command(label="✏️ Modificar este registro", command=self._abrir_modal_edicion)
        self._menu_contextual.add_separator()
        self._menu_contextual.add_command(label="🗑️ Eliminar esta medición", command=self._confirmar_borrado_seleccion)
        self._menu_contextual.add_command(label="🗑️ Eliminar lote completo", command=self._eliminar_lote_menu)
        self._tree_historial.bind("<Button-2>", self._mostrar_menu_contextual)
        self._tree_historial.bind("<Button-3>", self._mostrar_menu_contextual)

    def _mostrar_menu_contextual(self, event) -> None:
        if not self._admin_activo:
            return
        item = self._tree_historial.identify_row(event.y)
        if item:
            self._tree_historial.selection_set(item)
            self._menu_contextual.post(event.x_root, event.y_root)

    def _obtener_fila_seleccionada(self) -> Optional[dict]:
        sel = self._tree_historial.selection()
        if not sel:
            messagebox.showwarning("Selección vacía", "Por favor selecciona una fila del historial.", parent=self)
            return None
        med_id = int(sel[0])
        for row in self._datos_historial_actual:
            if row["medicion_id"] == med_id:
                return row
        return None

    # -----------------------------------------------------------------------
    # Modal de Edición (Modo Administrador)
    # -----------------------------------------------------------------------

    def _abrir_modal_edicion(self) -> None:
        if not self._admin_activo:
            return
        reg = self._obtener_fila_seleccionada()
        if not reg:
            return

        modal = ctk.CTkToplevel(self)
        modal.title(f"Editar Muestra #{reg['num_muestra']} — {reg['nombre_lote']}")
        modal.configure(fg_color=UI_COLORS["bg_primary"])
        centrar_ventana(modal, 520, 720)
        modal.transient(self)
        modal.grab_set()

        card = ctk.CTkFrame(modal, fg_color=UI_COLORS["bg_card"], corner_radius=12, border_width=1, border_color=UI_COLORS["border"])
        card.pack(fill="both", expand=True, padx=16, pady=16)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card,
            text=f"✏️ Editar Muestra No. {reg['num_muestra']}",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=UI_COLORS["text_primary"],
        ).grid(row=0, column=0, padx=20, pady=(16, 10), sticky="w")

        # Lote
        ctk.CTkLabel(card, text="Nombre del Lote:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=1, column=0, padx=20, sticky="w")
        ent_lote = ctk.CTkEntry(card, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_lote.insert(0, reg["nombre_lote"])
        ent_lote.grid(row=2, column=0, padx=20, pady=(2, 10), sticky="ew")

        # Variedad
        ctk.CTkLabel(card, text="Variedad de Semilla:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=3, column=0, padx=20, sticky="w")
        actual_sem = next((t for t in TIPOS_SEMILLA if AN.clave_variedad(t) == AN.clave_variedad(reg["tipo_semilla"])),
                          str(reg["tipo_semilla"] or TIPOS_SEMILLA[0]))
        valores_sem = TIPOS_SEMILLA if actual_sem in TIPOS_SEMILLA else TIPOS_SEMILLA + [actual_sem]
        cmb_sem = ctk.CTkComboBox(
            card, values=valores_sem, state="readonly",
            fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"]
        )
        cmb_sem.set(actual_sem)
        cmb_sem.grid(row=4, column=0, padx=20, pady=(2, 10), sticky="ew")

        # Tratamiento
        ctk.CTkLabel(card, text="Tiempo de Tratamiento:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=5, column=0, padx=20, sticky="w")
        ent_trat = ctk.CTkEntry(card, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_trat.insert(0, reg["tiempo_tratamiento"])
        ent_trat.grid(row=6, column=0, padx=20, pady=(2, 10), sticky="ew")

        # Factores del tratamiento (se aplican a todo el lote)
        fac = ctk.CTkFrame(card, fg_color="transparent")
        fac.grid(row=9, column=0, padx=20, pady=(0, 6), sticky="ew")
        fac.grid_columnconfigure((0, 1), weight=1)
        def _lbl(t, r, c):
            ctk.CTkLabel(fac, text=t, font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=UI_COLORS["text_secondary"]).grid(row=r, column=c, sticky="w", padx=(0 if c == 0 else 6, 0))
        def _cmb(vals, actual, r, c):
            cb = ctk.CTkComboBox(fac, values=vals, state="readonly", fg_color=UI_COLORS["input_bg"],
                                 border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
            cb.set(actual if actual in vals else vals[0])
            cb.grid(row=r, column=c, padx=(0 if c == 0 else 6, 0), pady=(2, 8), sticky="ew")
            return cb
        _lbl("Tipo de gas", 0, 0)
        cmb_gas = _cmb(TIPOS_GAS, reg.get("gas"), 1, 0)
        _lbl("Flujo (L/min)", 0, 1)
        ent_flujo = ctk.CTkEntry(fac, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"],
                                 text_color=UI_COLORS["text_primary"])
        if reg.get("flujo") is not None:
            ent_flujo.insert(0, f"{float(reg['flujo']):g}")
        ent_flujo.grid(row=1, column=1, padx=(6, 0), pady=(2, 8), sticky="ew")
        _lbl("Orientación", 2, 0)
        cmb_ori = _cmb(ORIENTACIONES, reg.get("orientacion"), 3, 0)
        _lbl("Morfología", 2, 1)
        cmb_mor = _cmb(MORFOLOGIAS, reg.get("morfologia"), 3, 1)
        _lbl("Tipo de tratamiento", 4, 0)
        cmb_tipo = _cmb(TIPOS_TRATAMIENTO, reg.get("tipo_tratamiento") or "Plasma", 5, 0)
        _lbl("Temperatura (libre)", 4, 1)
        ent_temp = ctk.CTkEntry(fac, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"],
                                text_color=UI_COLORS["text_primary"])
        ent_temp.insert(0, str(reg.get("temperatura") or "Ambiente"))
        ent_temp.grid(row=5, column=1, padx=(6, 0), pady=(2, 8), sticky="ew")
        ctk.CTkLabel(card, text="Tipo de tratamiento, temperatura, gas, flujo, orientación y morfología cambian en todo el lote.",
                     font=ctk.CTkFont(size=10), text_color=UI_COLORS["text_muted"]).grid(row=8, column=0, padx=20, sticky="w")

        # Medidas biométricas en fila
        med_frame = ctk.CTkFrame(card, fg_color="transparent")
        med_frame.grid(row=7, column=0, padx=20, pady=(4, 14), sticky="ew")
        med_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(med_frame, text="Raíz (mm)", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=0, column=0, sticky="w")
        ent_raiz = ctk.CTkEntry(med_frame, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_raiz.insert(0, f"{float(reg['longitud_raiz']):.2f}")
        ent_raiz.grid(row=1, column=0, padx=(0, 6), pady=2, sticky="ew")

        ctk.CTkLabel(med_frame, text="Tallo (mm)", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=0, column=1, sticky="w")
        ent_tallo = ctk.CTkEntry(med_frame, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_tallo.insert(0, f"{float(reg['longitud_tallo']):.2f}")
        ent_tallo.grid(row=1, column=1, padx=6, pady=2, sticky="ew")

        ctk.CTkLabel(med_frame, text="Cant. Hojas", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=0, column=2, sticky="w")
        ent_hojas = ctk.CTkEntry(med_frame, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_hojas.insert(0, str(reg["cantidad_hojas"] if reg["cantidad_hojas"] is not None else 0))
        ent_hojas.grid(row=1, column=2, padx=(6, 0), pady=2, sticky="ew")

        def guardar_edicion():
            ok_r, val_r, msg_r = validar_valor_numerico(ent_raiz.get(), "raíz")
            ok_t, val_t, msg_t = validar_valor_numerico(ent_tallo.get(), "tallo")
            if not ok_r or val_r is None:
                messagebox.showerror("Error", msg_r or "Raíz inválida", parent=modal)
                return
            if not ok_t or val_t is None:
                messagebox.showerror("Error", msg_t or "Tallo inválido", parent=modal)
                return
            try:
                val_h = parsear_cantidad_hojas(ent_hojas.get())
                if val_h is None:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Error", "La cantidad de hojas debe ser un entero >= 0.", parent=modal)
                return

            ok_f, val_f, msg_f = validar_flujo(ent_flujo.get())
            if not ok_f:
                messagebox.showerror("Error", msg_f, parent=modal)
                return
            ok_tp, val_tp, msg_tp = validar_temperatura(ent_temp.get())
            if not ok_tp:
                messagebox.showerror("Error", msg_tp, parent=modal)
                return
            exito, msg = self.db.actualizar_medicion(
                gas=cmb_gas.get(), flujo=val_f, orientacion=cmb_ori.get(), morfologia=cmb_mor.get(),
                temperatura=val_tp, tipo_tratamiento=cmb_tipo.get(),
                medicion_id=reg["medicion_id"],
                experimento_id=reg["experimento_id"],
                nombre_lote=ent_lote.get().strip(),
                tipo_semilla=cmb_sem.get(),
                tiempo_tratamiento=ent_trat.get().strip(),
                longitud_raiz=val_r,
                longitud_tallo=val_t,
                cantidad_hojas=val_h,
            )
            if exito:
                messagebox.showinfo("Actualizado", msg, parent=modal)
                modal.destroy()
                self._cargar_historial()
            else:
                messagebox.showerror("Error", msg, parent=modal)

        btn_box = ctk.CTkFrame(card, fg_color="transparent")
        btn_box.grid(row=10, column=0, padx=20, pady=(10, 16), sticky="ew")
        btn_box.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_box, text="Cancelar", fg_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"],
            command=modal.destroy
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btn_box, text="💾 Guardar Cambios", fg_color=UI_COLORS["accent_green"],
            hover_color=UI_COLORS["accent_hover"], text_color="#ffffff",
            command=guardar_edicion
        ).grid(row=0, column=1, padx=(6, 0), sticky="ew")

    # -----------------------------------------------------------------------
    # Eliminación de Registros
    # -----------------------------------------------------------------------

    def _confirmar_borrado_seleccion(self) -> None:
        reg = self._obtener_fila_seleccionada()
        if not reg:
            return

        resp = messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Deseas eliminar la muestra #{reg['num_muestra']} del lote '{reg['nombre_lote']}'?",
            parent=self,
        )
        if resp:
            exito, msg = self.db.eliminar_medicion(reg["medicion_id"])
            if exito:
                messagebox.showinfo("Registro eliminado", msg, parent=self)
                self._cargar_historial()
            else:
                messagebox.showerror("Error", msg, parent=self)

    def _eliminar_lote_menu(self) -> None:
        reg = self._obtener_fila_seleccionada()
        if not reg:
            return
        nombre_lote = reg["nombre_lote"]
        resp = messagebox.askyesno(
            "Confirmar eliminación de lote",
            f"¿Estás seguro de que deseas eliminar COMPLETAMENTE el lote '{nombre_lote}' y todas sus muestras?\nEsta acción no se puede deshacer.",
            parent=self,
        )
        if resp:
            exito, msg = self.db.eliminar_lote_por_nombre(nombre_lote)
            if exito:
                messagebox.showinfo("Lote eliminado", msg, parent=self)
                self._cargar_historial()
            else:
                messagebox.showerror("Error", msg, parent=self)

    def _cargar_historial(self) -> None:
        # Obtener datos de la base de datos sin filtrar individualmente en SQL para manejar conjuntos múltiples
        datos = self.db.obtener_historial_con_id(
            texto_busqueda=self._var_busqueda.get().strip() if hasattr(self, "_var_busqueda") else "",
            filtro_semilla="",
            filtro_germino="",
        )
        self._historial_crudo = datos

        # Filtrar por rango de fechas (Desde / Hasta)
        f_desde = self._var_hist_desde.get() if hasattr(self, "_var_hist_desde") else "Todas"
        f_hasta = self._var_hist_hasta.get() if hasattr(self, "_var_hist_hasta") else "Todas"

        self._datos_historial_actual = []
        for row in datos:
            fecha_reg = str(row.get("fecha_registro", ""))[:10]
            if f_desde != "Todas" and fecha_reg < f_desde:
                continue
            if f_hasta != "Todas" and fecha_reg > f_hasta:
                continue

            # Filtrar por Variedad múltiple (sin importar el nombre técnico entre paréntesis)
            sel = getattr(self, "_sel_variedades", None)
            if sel is not None and AN.clave_variedad(row["tipo_semilla"]) not in sel:
                continue

            # Filtrar por Tratamiento múltiple (45, 45 s y 45s cuentan como el mismo)
            sel = getattr(self, "_sel_tratamientos", None)
            if sel is not None and self._tratamiento_normalizado(row["tiempo_tratamiento"]) not in sel:
                continue

            # Filtrar por Germinación múltiple
            sel = getattr(self, "_sel_germino", None)
            if sel is not None and str(row["germino"]) not in sel:
                continue

            self._datos_historial_actual.append(row)

        for item in self._tree_historial.get_children():
            self._tree_historial.delete(item)

        for idx, row in enumerate(self._datos_historial_actual):
            es_par = (idx % 2 == 1)
            if int(row["germino"]) == 1:
                estado_txt = "✓  Germinó"
                tag = "alt_si" if es_par else "germino_si"
            else:
                estado_txt = "✗  No germinó"
                tag = "alt_no" if es_par else "germino_no"

            hojas_txt = "" if row["cantidad_hojas"] is None else str(row["cantidad_hojas"])
            fecha_txt = str(row.get("fecha_registro", ""))[:10]

            self._tree_historial.insert(
                "", "end",
                iid=str(row["medicion_id"]),
                values=(
                    fecha_txt,
                    truncar_texto(row["nombre_lote"], 22),
                    row["tipo_semilla"],
                    row["num_muestra"],
                    row.get("tipo_tratamiento") or "Plasma",
                    row["tiempo_tratamiento"],
                    row.get("temperatura", "Ambiente"),
                    row.get("gas") or "Sin gas",  # lotes anteriores a estos campos: valores por defecto
                    "" if row.get("flujo") is None else f"{float(row['flujo']):g}",
                    row.get("orientacion") or "Aleatorio",
                    row.get("morfologia") or "Plana",
                    hojas_txt,
                    f"{float(row['longitud_raiz']):.2f}",
                    f"{float(row['longitud_tallo']):.2f}",
                    estado_txt,
                ),
                tags=(tag,),
            )

        total = len(self._datos_historial_actual)
        self._lbl_total.configure(text=f"{total} registro{'s' if total != 1 else ''}")
    # -----------------------------------------------------------------------
    # Exportación a Excel (.xlsx) y PDF (.pdf)
    # -----------------------------------------------------------------------

    def _exportar_excel(self) -> None:
        if not self._datos_historial_actual:
            messagebox.showwarning("Sin datos", "No hay registros en la vista actual para exportar.", parent=self)
            return

        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar Reporte en Excel",
            defaultextension=".xlsx",
            initialfile=f"Reporte_BioGrow_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            # Mismo formato que la plantilla (machote): el archivo se puede volver a cargar o importar.
            contenido = AN.excel_machote(self._datos_historial_actual,
                                         titulo="BIOGROW ANALYTICS — REPORTE DE GERMINACIÓN Y BIOMETRÍA DE SEMILLAS")
            with open(ruta, "wb") as f:
                f.write(contenido)
            messagebox.showinfo("✅ Exportación Exitosa", f"Archivo Excel guardado correctamente en:\n{ruta}", parent=self)
        except ImportError:
            messagebox.showerror("Librería faltante", "Instala openpyxl ejecutando en tu terminal:\npip install openpyxl", parent=self)
        except Exception as e:
            messagebox.showerror("Error al exportar", f"No se pudo generar el archivo Excel:\n{e}", parent=self)

    def _exportar_pdf(self) -> None:
        if not self._datos_historial_actual:
            messagebox.showwarning("Sin datos", "No hay registros en la vista actual para exportar.", parent=self)
            return

        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar Reporte en PDF",
            defaultextension=".pdf",
            initialfile=f"Reporte_BioGrow_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
        )
        if not ruta:
            return

        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

            doc = SimpleDocTemplate(
                ruta,
                pagesize=landscape(letter),
                rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
            )
            elements = []
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                "BioTitle",
                parent=styles["Heading1"],
                fontSize=16,
                textColor=colors.HexColor("#0f172a"),
                spaceAfter=4,
            )
            sub_style = ParagraphStyle(
                "BioSub",
                parent=styles["Normal"],
                fontSize=10,
                textColor=colors.HexColor("#475569"),
                spaceAfter=14,
            )

            total = len(self._datos_historial_actual)
            germ = sum(1 for r in self._datos_historial_actual if int(r["germino"]) == 1)
            pct = (germ / total * 100) if total > 0 else 0

            elements.append(Paragraph("BioGrow Analytics — Reporte Consolidado de Semillas", title_style))
            elements.append(Paragraph(
                f"Fecha de exportación: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "
                f"Total de muestras: <b>{total}</b>  |  "
                f"Germinadas: <b>{germ} ({pct:.1f}%)</b>",
                sub_style
            ))

            data = [["Lote Origen", "Variedad de Semilla", "No.", "Tratamiento", "Hojas", "Raíz (mm)", "Tallo (mm)", "Resultado"]]
            for r in self._datos_historial_actual:
                estado = "Germinó" if int(r["germino"]) == 1 else "No germinó"
                data.append([
                    truncar_texto(r["nombre_lote"], 25),
                    r["tipo_semilla"],
                    str(r["num_muestra"]),
                    r["tiempo_tratamiento"],
                    "" if r["cantidad_hojas"] is None else str(r["cantidad_hojas"]),
                    f"{float(r['longitud_raiz']):.2f}",
                    f"{float(r['longitud_tallo']):.2f}",
                    estado,
                ])

            col_widths = [125, 155, 40, 95, 55, 75, 75, 95]
            t = Table(data, colWidths=col_widths, repeatRows=1)

            table_style = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#059669")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("ALIGN", (0, 1), (1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ])

            for i in range(1, len(data)):
                bg = colors.HexColor("#f8fafc") if i % 2 == 0 else colors.white
                table_style.add("BACKGROUND", (0, i), (-1, i), bg)
                if data[i][7] == "Germinó":
                    table_style.add("TEXTCOLOR", (7, i), (7, i), colors.HexColor("#059669"))
                    table_style.add("FONTNAME", (7, i), (7, i), "Helvetica-Bold")
                else:
                    table_style.add("TEXTCOLOR", (7, i), (7, i), colors.HexColor("#dc2626"))

            t.setStyle(table_style)
            elements.append(t)
            doc.build(elements)

            messagebox.showinfo("✅ Exportación Exitosa", f"Archivo PDF guardado correctamente en:\n{ruta}", parent=self)
        except ImportError:
            messagebox.showerror("Librería faltante", "Instala reportlab ejecutando en tu terminal:\npip install reportlab", parent=self)
        except Exception as e:
            messagebox.showerror("Error al exportar", f"No se pudo generar el archivo PDF:\n{e}", parent=self)

    def _abrir_ventana_graficos_historial(self) -> None:
        """Abre una nueva ventana modal con múltiples gráficas basadas en el historial filtrado actual."""
        if not self._datos_historial_actual:
            messagebox.showwarning("Sin datos", "No hay registros en el historial con los filtros actuales para graficar.", parent=self)
            return

        # Convertir la lista actual del historial filtrado en DataFrame
        df_base = pd.DataFrame(self._datos_historial_actual)
        # Variedades con y sin nombre científico, y tiempos escritos distinto (45, 45 s), cuentan como uno
        df_base["tipo_semilla"] = df_base["tipo_semilla"].map(
            lambda v: next((t for t in TIPOS_SEMILLA if AN.clave_variedad(t) == AN.clave_variedad(v)), v))
        df_base["tiempo_tratamiento"] = df_base["tiempo_tratamiento"].map(self._tratamiento_normalizado)
        
        # Normalizar nombres de columnas para que coincidan con las gráficas
        if "tiempo_tratamiento" in df_base.columns and "Tratamiento" not in df_base.columns:
            df_base["Tratamiento"] = df_base["tiempo_tratamiento"]
        if "longitud_raiz" in df_base.columns and "Raíz (mm)" not in df_base.columns:
            df_base["Raíz (mm)"] = pd.to_numeric(df_base["longitud_raiz"], errors="coerce")
        if "longitud_tallo" in df_base.columns and "Tallo (mm)" not in df_base.columns:
            df_base["Tallo (mm)"] = pd.to_numeric(df_base["longitud_tallo"], errors="coerce")
        if "germino" in df_base.columns and "Estado" not in df_base.columns:
            df_base["Estado"] = df_base["germino"].apply(lambda x: "Germinó" if int(x) == 1 else "No germinó")

        if df_base.empty:
            return

        # Crear ventana nueva (CTkToplevel)
        ventana_graf = ctk.CTkToplevel(self)
        ventana_graf.title("📊 Análisis Gráfico del Historial Filtrado — BioGrow")
        ventana_graf.geometry("1100x700")
        ventana_graf.grab_set()

        from matplotlib.figure import Figure
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        p = TTK_PALETTES[self._modo_actual]
        fig = Figure(figsize=(10.5, 6.5), dpi=110, facecolor=p["bg_card"])
        fig.subplots_adjust(hspace=0.35, wspace=0.25, bottom=0.10, top=0.90, left=0.08, right=0.95)

        # Gráfico 1: Raíz por Tratamiento
        ax1 = fig.add_subplot(221)
        ax1.set_title("Promedio de Raíz (mm) por Tratamiento", color=p["text_primary"], fontsize=10, fontweight="bold")
        if "Tratamiento" in df_base.columns and "Raíz (mm)" in df_base.columns:
            df_base.groupby("Tratamiento")["Raíz (mm)"].mean().plot(kind="bar", ax=ax1, color="#2a78d6", alpha=0.85)
        ax1.tick_params(colors=p["text_primary"], labelsize=8)

        # Gráfico 2: Tallo por Tratamiento
        ax2 = fig.add_subplot(222)
        ax2.set_title("Promedio de Tallo (mm) por Tratamiento", color=p["text_primary"], fontsize=10, fontweight="bold")
        if "Tratamiento" in df_base.columns and "Tallo (mm)" in df_base.columns:
            df_base.groupby("Tratamiento")["Tallo (mm)"].mean().plot(kind="bar", ax=ax2, color="#10b981", alpha=0.85)
        ax2.tick_params(colors=p["text_primary"], labelsize=8)

        # Gráfico 3: Porcentaje de Germinación
        ax3 = fig.add_subplot(223)
        ax3.set_title("Porcentaje de Germinación (%)", color=p["text_primary"], fontsize=10, fontweight="bold")
        if "Tratamiento" in df_base.columns and "Estado" in df_base.columns:
            germ_s = df_base.groupby("Tratamiento")["Estado"].apply(lambda x: (x.str.lower().str.contains("germinó|germino|1|true")).mean() * 100)
            germ_s.plot(kind="bar", ax=ax3, color="#f59e0b", alpha=0.85)
        ax3.tick_params(colors=p["text_primary"], labelsize=8)

        # Gráfico 4: Dispersión Alométrica
        ax4 = fig.add_subplot(224)
        ax4.set_title("Correlación Raíz vs Tallo", color=p["text_primary"], fontsize=10, fontweight="bold")
        if "Raíz (mm)" in df_base.columns and "Tallo (mm)" in df_base.columns:
            ax4.scatter(df_base["Raíz (mm)"], df_base["Tallo (mm)"], alpha=0.6, color="#6366f1", s=25)
            ax4.set_xlabel("Raíz (mm)", color=p["text_primary"], fontsize=9)
            ax4.set_ylabel("Tallo (mm)", color=p["text_primary"], fontsize=9)
        ax4.tick_params(colors=p["text_primary"], labelsize=8)

        canvas = FigureCanvasTkAgg(fig, master=ventana_graf)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=15, pady=15)

    # -----------------------------------------------------------------------
    # Datos para el análisis: base de datos (pestaña 1) + archivos Excel
    # Los cálculos son los mismos del dashboard (motor_biogrow.py, vía analisis_biogrow.py)
    # -----------------------------------------------------------------------

    def _cargar_excels_externos(self) -> None:
        rutas = filedialog.askopenfilenames(
            parent=self,
            title="Seleccionar archivos Excel de experimentos (.xlsx)",
            filetypes=[("Archivos Excel", "*.xlsx")],
        )
        if not rutas:
            return
        for r in rutas:
            if r not in self._archivos_excel_cargados:
                self._archivos_excel_cargados.append(r)
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            self._df_excel_cache, self._advertencias_excel = AN.leer_excels(self._archivos_excel_cargados)
        finally:
            self.configure(cursor="")
        n_arch = len(self._archivos_excel_cargados)
        n_filas = len(self._df_excel_cache)
        self._actualizar_estado_excel()
        self._guardar_config_excel()
        if not self._df_excel_cache.empty and self._var_fuente_datos.get() == "Solo MySQL":
            self._var_fuente_datos.set("MySQL + Excel (Combinados)")
        self._unidad_tocada = False
        self._ejecutar_analisis_ui()

        msg = f"Se cargaron {n_arch} archivo(s) Excel con {n_filas} semillas reconocidas."
        if self._advertencias_excel:
            msg += "\n\nNotas del lector:\n• " + "\n• ".join(self._advertencias_excel[:8])
        messagebox.showinfo("Archivos Excel Cargados", msg, parent=self)

    def _limpiar_excels_externos(self) -> None:
        self._archivos_excel_cargados.clear()
        self._df_excel_cache = pd.DataFrame()
        self._advertencias_excel = []
        self._actualizar_estado_excel()
        self._guardar_config_excel()
        self._var_fuente_datos.set("Solo MySQL")
        self._unidad_tocada = False
        self._ejecutar_analisis_ui()

    # --- Recordar los Excel cargados entre sesiones -------------------------

    def _ruta_config_excel(self) -> str:
        carpeta = os.path.dirname(os.path.abspath(self.db.db_path))
        return os.path.join(carpeta, "biogrow_excels.json")

    def _leer_config_excel(self) -> dict:
        try:
            with open(self._ruta_config_excel(), encoding="utf-8") as f:
                datos = json.load(f)
            return datos if isinstance(datos, dict) else {}
        except (OSError, ValueError):
            return {}

    def _guardar_config_excel(self) -> None:
        recordar = self._var_recordar_excel.get() if hasattr(self, "_var_recordar_excel") else True
        datos = {"recordar": recordar, "archivos": list(self._archivos_excel_cargados) if recordar else []}
        try:
            with open(self._ruta_config_excel(), "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)
        except OSError:
            pass  # no poder recordar los archivos no debe impedir usar la app

    def _actualizar_estado_excel(self, faltantes: int = 0) -> None:
        n = len(self._archivos_excel_cargados)
        texto = f"📂 {n} Excel ({len(self._df_excel_cache)} semillas)" if n else "📂 0 Excel cargados"
        if faltantes:
            texto += f" · {faltantes} ya no se encontró"
        self._lbl_estado_excel.configure(text=texto)

    def _restaurar_excels(self) -> None:
        """Al abrir la app, vuelve a cargar los Excel de la sesión anterior (si 'Recordar Excel' está activo)."""
        cfg = self._leer_config_excel()
        if not cfg.get("recordar", True) or not cfg.get("archivos"):
            return
        rutas = [r for r in cfg["archivos"] if os.path.isfile(r)]
        faltantes = len(cfg["archivos"]) - len(rutas)
        if not rutas:
            self._guardar_config_excel()
            self._actualizar_estado_excel(faltantes)
            return
        self._archivos_excel_cargados = rutas
        try:
            self._df_excel_cache, self._advertencias_excel = AN.leer_excels(rutas)
        except Exception:
            self._archivos_excel_cargados, self._df_excel_cache = [], pd.DataFrame()
        self._actualizar_estado_excel(faltantes)
        if faltantes:
            self._guardar_config_excel()
        if not self._df_excel_cache.empty and self._var_fuente_datos.get() == "Solo MySQL":
            self._var_fuente_datos.set("MySQL + Excel (Combinados)")
        self._unidad_tocada = False
        if self._pag_calculos.winfo_ismapped() or self._pag_tablero.winfo_ismapped():
            self._ejecutar_analisis_ui()

    # --- Importar los Excel cargados a la base de datos ---------------------

    def _importar_excel_desde_historial(self) -> None:
        """Elige uno o más Excel y guarda sus lotes directo en la base (sin pasar por «Cargar Excel»)."""
        rutas = filedialog.askopenfilenames(parent=self, title="Importar Excel a la base de datos",
                                            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm *.xls")])
        if not rutas:
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            df, avisos = AN.leer_excels(list(rutas))
        except Exception as e:
            messagebox.showerror("Error de lectura", f"No se pudo leer el archivo:\n{e}", parent=self)
            return
        finally:
            self.configure(cursor="")
        if df.empty:
            detalle = ("\n\n• " + "\n• ".join(avisos[:6])) if avisos else ""
            messagebox.showwarning("Sin datos", "No se reconocieron semillas en el archivo. Usa el formato de la "
                                   "«📋 Plantilla Excel» (pestaña Cálculos y Modelos)." + detalle, parent=self)
            return
        self._importar_df_a_bd(df, desde_cargados=False)

    def _importar_excels_a_bd(self) -> None:
        if self._df_excel_cache is None or self._df_excel_cache.empty:
            messagebox.showwarning("Sin Excel", "Primero carga uno o más archivos Excel con «➕ Cargar Excel».",
                                   parent=self)
            return
        self._importar_df_a_bd(self._df_excel_cache, desde_cargados=True)

    def _importar_df_a_bd(self, df: pd.DataFrame, desde_cargados: bool) -> None:
        lotes, r = AN.preparar_importacion(df, self.db.nombres_de_lotes())
        notas = []
        if r["repetidos"]:
            notas.append(f"{len(r['repetidos'])} lote(s) ya estaban en la base y no se repiten.")
        if r["vacias"]:
            notas.append(f"{r['vacias']} caja(s) vacía(s) de plantilla no se importan.")
        if r["sin_pareja"]:
            notas.append(f"{r['sin_pareja']} columna(s) de archivos con un tratamiento por columna no se importan "
                         "(raíz y tallo no son de las mismas semillas).")
        if r["fecha_hoy"]:
            notas.append(f"{r['fecha_hoy']} lote(s) no traen fecha; se registran con la fecha de hoy.")
        if not lotes:
            messagebox.showinfo("Nada que importar", "No hay lotes nuevos para importar.\n\n" + "\n".join(notas),
                                parent=self)
            return
        texto = (f"Se van a agregar a la base de datos {r['lotes']} lote(s) con {r['semillas']} semillas "
                 f"de {len(r['archivos'])} archivo(s).\n\nAparecerán en el Historial de la primera pestaña y "
                 "se podrán editar o borrar como cualquier registro.")
        if notas:
            texto += "\n\n• " + "\n• ".join(notas)
        if not messagebox.askyesno("Importar Excel a la base de datos", texto + "\n\n¿Continuar?", parent=self):
            return
        ok, msg = self.db.guardar_lotes_importados(lotes)
        if not ok:
            messagebox.showerror("Error al importar", msg, parent=self)
            return
        if not desde_cargados:
            self._firma_bd = None
            self._cargar_historial()
            if self._pag_calculos.winfo_ismapped() or self._pag_tablero.winfo_ismapped():
                self._ejecutar_analisis_ui()
            messagebox.showinfo("✅ Importación terminada", msg, parent=self)
            return
        # Lo importado ya está en la base: se quita de los Excel cargados para no contarlo dos veces.
        importados = {l["archivo"] for l in lotes}
        self._archivos_excel_cargados = [a for a in self._archivos_excel_cargados
                                         if os.path.basename(a) not in importados]
        self._df_excel_cache, self._advertencias_excel = AN.leer_excels(self._archivos_excel_cargados)
        self._actualizar_estado_excel()
        self._guardar_config_excel()
        if self._df_excel_cache.empty:
            self._var_fuente_datos.set("Solo MySQL")
        self._firma_bd = None
        self._cargar_historial()
        self._ejecutar_analisis_ui()
        messagebox.showinfo("✅ Importación terminada",
                            msg + "\n\nLos archivos importados se quitaron de «Excel cargados» para no contarlos "
                            "dos veces; ahora sus datos se leen desde la base.", parent=self)

    def _obtener_dataframe_activo(self) -> pd.DataFrame:
        """Datos según la fuente elegida, leídos con el mismo lector del dashboard (en cm)."""
        modo = self._var_fuente_datos.get() if hasattr(self, "_var_fuente_datos") else "Solo MySQL"
        df_bd = pd.DataFrame()
        if not modo.startswith("Solo Excel"):
            registros = self.db.obtener_historial_con_id()
            firma = (len(registros), max((r.get("medicion_id") or 0 for r in registros), default=0),
                     hash(tuple((r.get("medicion_id"), r.get("longitud_raiz"), r.get("longitud_tallo"),
                                 r.get("germino"), r.get("tiempo_tratamiento"), r.get("temperatura"),
                                 r.get("nombre_lote"), r.get("tipo_semilla"), r.get("gas"), r.get("flujo"),
                                 r.get("orientacion"), r.get("morfologia"), r.get("tipo_tratamiento"))
                                for r in registros)))
            if getattr(self, "_firma_bd", None) != firma:
                self._df_bd_cache, _ = AN.leer_base(registros)
                self._firma_bd = firma
            df_bd = self._df_bd_cache
        return AN.combinar(df_bd, self._df_excel_cache, modo, self._nombres_tecnicos())

    # --- Nombres técnicos de las variedades ---------------------------------

    def _ruta_nombres_tecnicos(self) -> str:
        return os.path.join(os.path.dirname(os.path.abspath(self.db.db_path)), "biogrow_variedades.json")

    def _nombres_tecnicos(self) -> dict:
        return AN.leer_nombres_tecnicos(self._ruta_nombres_tecnicos())

    def _editar_nombres_tecnicos(self) -> None:
        actuales = self._nombres_tecnicos()
        base = {k for k in actuales if not str(k).startswith("_")}
        if getattr(self, "_df_activo", None) is not None and not self._df_activo.empty:
            base |= {re.sub(r"\s*\([^)]*\)$", "", v) for v in AN.variedades(self._df_activo)}
        base = sorted(b for b in base if b and b not in ("?", "Sin especificar"))
        top = ctk.CTkToplevel(self)
        top.title("Nombres técnicos y científicos de las variedades")
        top.configure(fg_color=UI_COLORS["bg_primary"])
        centrar_ventana(top, 520, 120 + 46 * len(base))
        top.transient(self)
        top.grab_set()
        card = ctk.CTkFrame(top, fg_color=UI_COLORS["bg_card"], corner_radius=12)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card, text="Nombre técnico o científico de cada variedad (vacío = solo el nombre común)",
                     font=ctk.CTkFont(size=12, weight="bold"), text_color=UI_COLORS["text_primary"]).grid(
            row=0, column=0, columnspan=2, padx=12, pady=(10, 8), sticky="w")
        entradas = {}
        for i, v in enumerate(base, start=1):
            ctk.CTkLabel(card, text=v, text_color=UI_COLORS["text_secondary"],
                         font=ctk.CTkFont(size=11, weight="bold")).grid(row=i, column=0, padx=12, pady=4, sticky="w")
            e = ctk.CTkEntry(card, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"],
                             text_color=UI_COLORS["text_primary"], placeholder_text="Ej: H-392 A o Zea mays var. amylacea")
            if actuales.get(v):
                e.insert(0, actuales[v])
            e.grid(row=i, column=1, padx=12, pady=4, sticky="ew")
            entradas[v] = e

        def guardar():
            datos = {v: e.get().strip() for v, e in entradas.items()}
            try:
                AN.guardar_nombres_tecnicos(self._ruta_nombres_tecnicos(), datos)
            except OSError as err:
                messagebox.showerror("No se pudo guardar", str(err), parent=top)
                return
            top.destroy()
            self._ejecutar_analisis_ui()
        ctk.CTkButton(card, text="💾 Guardar", fg_color=UI_COLORS["accent_green"], text_color="#ffffff",
                      command=guardar).grid(row=len(base) + 1, column=0, columnspan=2, padx=12, pady=(10, 12), sticky="ew")

    def _descargar_plantilla(self) -> None:
        ruta = filedialog.asksaveasfilename(parent=self, title="Guardar plantilla de Excel", defaultextension=".xlsx",
                                            initialfile="Plantilla_BioGrow.xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            with open(ruta, "wb") as f:
                f.write(AN.excel_machote())
            messagebox.showinfo("✅ Plantilla guardada",
                                f"Plantilla guardada en:\n{ruta}\n\nLlena la hoja 'Datos' (una semilla por fila) y "
                                "cárgala con «➕ Cargar Excel». En la hoja 'Léeme' están los valores permitidos.",
                                parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la plantilla:\n{e}", parent=self)

    # -----------------------------------------------------------------------
    # Página 2: Cálculos y Modelos (lo mismo que muestra el dashboard)
    # -----------------------------------------------------------------------

    def _etiqueta_control(self, padre, texto):
        ctk.CTkLabel(padre, text=texto, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                     text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))

    def _construir_seccion_calculos(self, padre: ctk.CTkFrame) -> None:
        padre.grid_columnconfigure(0, weight=1)
        padre.grid_rowconfigure(3, weight=1)
        self._unidad_tocada = False
        self._df_bd_cache = pd.DataFrame()
        recalcular = lambda *_: self._ejecutar_analisis_ui()

        ctrl = ctk.CTkFrame(padre, fg_color=UI_COLORS["bg_card"], corner_radius=12,
                            border_width=1, border_color=UI_COLORS["border"])
        ctrl.grid(row=0, column=0, pady=(0, 8), sticky="ew")

        # Fila 1: fuente de datos y archivos Excel
        top_bar = ctk.CTkFrame(ctrl, fg_color="transparent")
        top_bar.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(top_bar, text="🗄️ Fuente de Datos:", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                     text_color=UI_COLORS["text_primary"]).pack(side="left", padx=(0, 8))
        self._var_fuente_datos = tk.StringVar(value="Solo MySQL")
        ctk.CTkSegmentedButton(
            top_bar, values=["Solo MySQL", "MySQL + Excel (Combinados)", "Solo Excel"],
            variable=self._var_fuente_datos, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            selected_color=UI_COLORS["accent_indigo"], command=recalcular,
        ).pack(side="left", padx=(0, 12))
        ctk.CTkButton(top_bar, text="➕ Cargar Excel", height=30, width=115,
                      fg_color=UI_COLORS["accent_green"], hover_color=UI_COLORS["accent_hover"],
                      text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._cargar_excels_externos).pack(side="left", padx=(0, 6))
        ctk.CTkButton(top_bar, text="🧹 Quitar", width=75, height=30, fg_color=UI_COLORS["bg_secondary"],
                      text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"],
                      font=ctk.CTkFont(size=11), command=self._limpiar_excels_externos).pack(side="left", padx=(0, 6))
        ctk.CTkButton(top_bar, text="📥 Importar a la base", width=150, height=30,
                      fg_color=UI_COLORS["accent_amber"], hover_color="#b45309", text_color="#ffffff",
                      font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._importar_excels_a_bd).pack(side="left", padx=(0, 8))
        self._var_recordar_excel = tk.BooleanVar(value=self._leer_config_excel().get("recordar", True))
        ctk.CTkCheckBox(top_bar, text="Recordar Excel", variable=self._var_recordar_excel,
                        font=ctk.CTkFont(size=11), checkbox_width=18, checkbox_height=18,
                        command=self._guardar_config_excel).pack(side="left", padx=(0, 8))
        self._lbl_estado_excel = ctk.CTkLabel(top_bar, text="📂 0 Excel cargados", font=ctk.CTkFont(size=11),
                                              text_color=UI_COLORS["text_secondary"])
        self._lbl_estado_excel.pack(side="left", padx=(4, 0))
        ctk.CTkButton(top_bar, text="🔄 Recalcular", width=100, height=30, fg_color=UI_COLORS["accent_indigo"],
                      text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._ejecutar_analisis_ui).pack(side="right")
        ctk.CTkButton(top_bar, text="📄 Reporte PDF", width=120, height=30,
                      fg_color=UI_COLORS["accent_red"], hover_color="#b91c1c", text_color="#ffffff",
                      font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._exportar_reporte_calculos).pack(side="right", padx=(0, 8))

        # Fila 2: qué analizar
        param_bar = ctk.CTkFrame(ctrl, fg_color="transparent")
        param_bar.pack(fill="x", padx=16, pady=4)
        self._etiqueta_control(param_bar, "Variable:")
        self._var_variable_calc = tk.StringVar(value="Raíz")
        ctk.CTkComboBox(param_bar, variable=self._var_variable_calc, values=["Raíz", "Tallo", "Germinó"],
                        width=105, state="readonly", command=recalcular).pack(side="left", padx=(0, 12))
        self._etiqueta_control(param_bar, "Comparar por:")
        self._var_agrupacion_calc = tk.StringVar(value=list(AN.OPCIONES_COMPARAR)[0])
        ctk.CTkComboBox(param_bar, variable=self._var_agrupacion_calc, values=list(AN.OPCIONES_COMPARAR),
                        width=265, state="readonly", command=recalcular).pack(side="left", padx=(0, 12))
        self._etiqueta_control(param_bar, "Variedad:")
        self._var_filtro_var_calc = tk.StringVar(value="Todas las variedades")
        self._cmb_variedad_calc = ctk.CTkComboBox(param_bar, variable=self._var_filtro_var_calc,
                                                  values=["Todas las variedades"], width=185, state="readonly",
                                                  command=recalcular)
        self._cmb_variedad_calc.pack(side="left", padx=(0, 12))
        self._etiqueta_control(param_bar, "Cada dato es:")
        self._var_unidad_calc = tk.StringVar(value="Por caja (promedio de sus semillas)")
        self._cmb_unidad_calc = ctk.CTkComboBox(
            param_bar, variable=self._var_unidad_calc,
            values=["Por caja (promedio de sus semillas)", "Por semilla"], width=235, state="readonly",
            command=recalcular)
        self._cmb_unidad_calc.pack(side="left", padx=(0, 12))
        self._etiqueta_control(param_bar, "Unidad:")
        self._var_long_calc = tk.StringVar(value="mm")

        def _unidad_cambiada(_=None):
            self._unidad_tocada = True
            self._ejecutar_analisis_ui()
        ctk.CTkSegmentedButton(param_bar, values=["mm", "cm"], variable=self._var_long_calc, width=80,
                               command=_unidad_cambiada).pack(side="left", padx=(0, 12))
        self._etiqueta_control(param_bar, "α:")
        self._var_alfa_calc = tk.StringVar(value="0.05")
        ctk.CTkSegmentedButton(param_bar, values=["0.05", "0.01"], variable=self._var_alfa_calc, width=100,
                               command=recalcular).pack(side="left")

        # Fila 3: descartes (mismas reglas del dashboard)
        desc_bar = ctk.CTkFrame(ctrl, fg_color="transparent")
        desc_bar.pack(fill="x", padx=16, pady=(4, 12))
        self._etiqueta_control(desc_bar, "🧹 Semillas que no germinaron:")
        self._var_no_germ = tk.StringVar(value="Excluirlas del promedio")
        ctk.CTkSegmentedButton(desc_bar, values=["Excluirlas del promedio", "Incluirlas con longitud 0"],
                               variable=self._var_no_germ, command=recalcular).pack(side="left", padx=(0, 14))
        self._var_ceros = tk.BooleanVar(value=False)
        self._var_excluir_atipicos = tk.BooleanVar(value=False)
        self._var_cajas_vacias = tk.BooleanVar(value=True)
        for texto, var in [("Excluir germinadas con valor = 0", self._var_ceros),
                           ("Excluir atípicos de cada grupo (1.5×IQR)", self._var_excluir_atipicos),
                           ("Excluir cajas vacías", self._var_cajas_vacias)]:
            ctk.CTkCheckBox(desc_bar, text=texto, variable=var, font=ctk.CTkFont(size=11),
                            checkbox_width=18, checkbox_height=18, command=recalcular).pack(side="left", padx=(0, 14))
        ctk.CTkButton(desc_bar, text="🏷️ Nombres técnicos", width=140, height=28, fg_color=UI_COLORS["bg_secondary"],
                      text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"], font=ctk.CTkFont(size=11),
                      command=self._editar_nombres_tecnicos).pack(side="right")
        ctk.CTkButton(desc_bar, text="📋 Plantilla Excel", width=130, height=28, fg_color=UI_COLORS["bg_secondary"],
                      text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"], font=ctk.CTkFont(size=11),
                      command=self._descargar_plantilla).pack(side="right", padx=(0, 8))

        # Dictamen y conclusión
        self._txt_conclusion_anova = ctk.CTkTextbox(
            padre, height=58, fg_color=UI_COLORS["bg_card"], border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(family="Segoe UI", size=12),
            corner_radius=10, wrap="word")
        self._txt_conclusion_anova.grid(row=1, column=0, pady=(0, 8), sticky="ew")

        # Tarjetas con los estadísticos principales
        tarjetas = ctk.CTkFrame(padre, fg_color="transparent")
        tarjetas.grid(row=2, column=0, pady=(0, 8), sticky="ew")
        tarjetas.grid_columnconfigure(tuple(range(6)), weight=1, uniform="tarj")
        self._tarjetas_calc = {}
        for i, (clave, titulo, color) in enumerate([
                ("anova", "🧮 ANOVA · F calculada vs F crítica", "accent_red"),
                ("p", "📉 p-valor", "accent_indigo"),
                ("gl", "🔢 Grados de libertad", "accent_amber"),
                ("cme", "📐 CM del error", "accent_indigo"),
                ("tukey", "🔠 Tukey", "accent_green"),
                ("lsd", "🔡 LSD de Fisher", "accent_green")]):
            card = ctk.CTkFrame(tarjetas, fg_color=UI_COLORS["bg_card"], corner_radius=10, border_width=2,
                                border_color=UI_COLORS[color])
            card.grid(row=0, column=i, padx=(0 if i == 0 else 4, 0 if i == 5 else 4), sticky="nsew")
            ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
                         text_color=UI_COLORS["text_secondary"], anchor="w").pack(anchor="w", padx=12, pady=(8, 0))
            valor = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
                                 text_color=UI_COLORS["text_primary"], anchor="w")
            valor.pack(anchor="w", padx=12)
            sub = ctk.CTkLabel(card, text="", font=ctk.CTkFont(family="Segoe UI", size=10),
                               text_color=UI_COLORS["text_secondary"], anchor="w", justify="left", wraplength=210)
            sub.pack(anchor="w", padx=12, pady=(0, 8))
            self._tarjetas_calc[clave] = (card, valor, sub)

        # Sub-pestañas con las mismas secciones del dashboard
        self._tabs_estadistica = ctk.CTkTabview(
            padre, fg_color=UI_COLORS["bg_card"], segmented_button_fg_color=UI_COLORS["bg_secondary"],
            segmented_button_selected_color=UI_COLORS["accent_green"],
            segmented_button_unselected_color=UI_COLORS["bg_secondary"], text_color=UI_COLORS["text_primary"],
            corner_radius=12, border_width=1, border_color=UI_COLORS["border"])
        self._tabs_estadistica.grid(row=3, column=0, sticky="nsew")
        t_res = self._tabs_estadistica.add("📌 Resumen")
        t_desc = self._tabs_estadistica.add("📊 Promedio, DE y varianza")
        t_anov = self._tabs_estadistica.add("🧮 ANOVA")
        t_tuk = self._tabs_estadistica.add("🔠 Tukey")
        t_lsd = self._tabs_estadistica.add("🔡 LSD")

        def _nota(contenedor, fila=0, texto=""):
            lbl = ctk.CTkLabel(contenedor, text=texto, anchor="w", justify="left", wraplength=1380,
                               font=ctk.CTkFont(family="Segoe UI", size=11), text_color=UI_COLORS["text_secondary"])
            lbl.grid(row=fila, column=0, sticky="ew", padx=6, pady=(2, 4))
            return lbl

        for t in (t_res, t_desc, t_anov, t_tuk, t_lsd):
            t.grid_columnconfigure(0, weight=1)

        t_res.grid_rowconfigure(1, weight=3)
        t_res.grid_rowconfigure(3, weight=1)
        self._nota_resumen = _nota(t_res, 0, "Grupos que comparten una letra no difieren entre sí. Tukey es la "
                                   "prueba más estricta; LSD separa más grupos.")
        self._tree_ranking = self._crear_tabla_estadistica(t_res, 1)
        self.after(400, self._restaurar_excels)
        ctk.CTkLabel(t_res, text="¿Qué grupos difieren? (Tukey)", anchor="w",
                     font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                     text_color=UI_COLORS["text_primary"]).grid(row=2, column=0, sticky="ew", padx=6, pady=(6, 0))
        self._txt_pares = ctk.CTkTextbox(t_res, height=90, fg_color=UI_COLORS["bg_secondary"], wrap="word",
                                         text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11))
        self._txt_pares.grid(row=3, column=0, sticky="nsew", padx=6, pady=(2, 6))

        t_desc.grid_rowconfigure(1, weight=1)
        _nota(t_desc, 0, "Varianza y desviación estándar muestrales (se divide entre n − 1). "
                         "CV = DE como % del promedio. Error estándar = DE / √n.")
        self._tree_desc = self._crear_tabla_estadistica(t_desc, 1)

        t_anov.grid_rowconfigure(1, weight=1)
        self._nota_anova = _nota(t_anov)
        self._tree_anova = self._crear_tabla_estadistica(t_anov, 1)

        t_tuk.grid_rowconfigure(1, weight=3)
        t_tuk.grid_rowconfigure(3, weight=2)
        self._nota_tukey = _nota(t_tuk)
        self._tree_tukey = self._crear_tabla_estadistica(t_tuk, 1)
        _nota(t_tuk, 2, "Letras de Tukey (misma letra = no difieren):")
        self._tree_letras_tukey = self._crear_tabla_estadistica(t_tuk, 3)

        t_lsd.grid_rowconfigure(1, weight=3)
        t_lsd.grid_rowconfigure(3, weight=2)
        self._nota_lsd = _nota(t_lsd)
        self._tree_lsd = self._crear_tabla_estadistica(t_lsd, 1)
        for t in (self._tree_ranking, self._tree_desc, self._tree_anova, self._tree_tukey, self._tree_lsd):
            self._activar_fondo_tabla(t, "calculos", "BioGrow.Treeview")
        _nota(t_lsd, 2, "Letras de LSD (misma letra = no difieren):")
        self._tree_letras_lsd = self._crear_tabla_estadistica(t_lsd, 3)

    def _crear_tabla_estadistica(self, contenedor, fila: int, selectmode: str = "browse") -> ttk.Treeview:
        marco = ctk.CTkFrame(contenedor, fg_color="transparent")
        marco.grid(row=fila, column=0, sticky="nsew", padx=4, pady=2)
        marco.grid_columnconfigure(0, weight=1)
        marco.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(marco, columns=(), show="headings", style="BioGrow.Treeview", selectmode=selectmode)
        vsb = ttk.Scrollbar(marco, orient="vertical", command=tree.yview, style="Dark.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(marco, orient="horizontal", command=tree.xview, style="Dark.Horizontal.TScrollbar")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self._tablas_estadisticas = getattr(self, "_tablas_estadisticas", []) + [tree]
        return tree

    def _llenar_tabla(self, tree: ttk.Treeview, df: Optional[pd.DataFrame], dec: int = 3, ids=None) -> None:
        tree.delete(*tree.get_children())
        cols, filas = AN.formatear_tabla(df, dec)
        tree["columns"] = cols
        for i, c in enumerate(cols):
            largo = max([len(c)] + [len(f[i]) for f in filas[:300]]) if filas else len(c)
            tree.heading(c, text=c, anchor="center")
            tree.column(c, width=min(max(70, int(largo * 7.4) + 18), 340), minwidth=55,
                        anchor="w" if i == 0 or c in ("Grupo", "Grupo A", "Grupo B", "Caja", "Motivo") else "center",
                        stretch=True)
        for k, f in enumerate(filas):
            texto = " ".join(f)
            if "✔" in texto:
                tag = "si"
            elif "✘" in texto:
                tag = "no"
            else:
                tag = "alt" if k % 2 else "norm"
            iid = ids[k] if ids is not None else None
            tree.insert("", "end", iid=iid, values=f, tags=(tag,))

    def _estilizar_tablas_estadisticas(self) -> None:
        p = TTK_PALETTES[self._modo_actual]
        for tree in getattr(self, "_tablas_estadisticas", []):
            tree.tag_configure("si", foreground=p["accent_green"], background=self._bg_fila(tree, p["bg_card"]))
            tree.tag_configure("no", foreground=p["accent_red"], background=self._bg_fila(tree, p["bg_card"]))
            tree.tag_configure("alt", background=self._bg_fila(tree, p["row_alt"]), foreground=p["text_primary"])
            tree.tag_configure("norm", background=self._bg_fila(tree, p["bg_card"]), foreground=p["text_primary"])

    def _opciones_descartes(self) -> dict:
        return {
            "cajas_vacias": self._var_cajas_vacias.get(),
            "no_germinadas": "excluir" if self._var_no_germ.get().startswith("Excluir") else "incluir",
            "ceros_germinadas": self._var_ceros.get(),
            "atipicos": self._var_excluir_atipicos.get(),
            "factor_iqr": 1.5,
        }

    def _analizar(self, df: pd.DataFrame, filtros_extra: Optional[dict] = None) -> dict:
        variedad = self._var_filtro_var_calc.get()
        f = dict(variedades_sel=[variedad] if variedad and variedad != "Todas las variedades" else None)
        f.update(filtros_extra or {})
        datos = AN.filtrar(df, **f) if not df.empty else df
        return AN.analizar(datos, self._var_variable_calc.get(), self._var_agrupacion_calc.get(),
                           self._var_unidad_calc.get(), float(self._var_alfa_calc.get()),
                           self._var_long_calc.get(), self._opciones_descartes())

    def _ejecutar_analisis_ui(self) -> None:
        if not hasattr(self, "_txt_conclusion_anova"):
            return
        self._estilizar_tablas_estadisticas()
        try:
            df = self._obtener_dataframe_activo()
        except Exception as e:  # el análisis nunca debe impedir usar la pestaña de registro
            df = pd.DataFrame()
            messagebox.showerror("Error al leer los datos", str(e), parent=self)
        self._df_activo = df

        # Listas que dependen de los datos
        vars_disp = ["Todas las variedades"] + AN.variedades(df)
        self._cmb_variedad_calc.configure(values=vars_disp)
        if self._var_filtro_var_calc.get() not in vars_disp:
            self._var_filtro_var_calc.set("Todas las variedades")
        if not self._unidad_tocada and not df.empty:
            self._var_long_calc.set(AN.unidad_por_defecto(df))
        self._actualizar_filtros_tablero(df)

        try:
            res = self._analizar(df)
        except Exception as e:
            res = {"ok": False, "mensaje": f"No se pudo hacer el análisis: {e}", "avisos": []}
        self._ultimo_resultado_stats = res
        self._mostrar_resultados_calculos(res)
        self._dibujar_grafica_tablero()

    def _escribir(self, caja: ctk.CTkTextbox, texto: str) -> None:
        caja.configure(state="normal")
        caja.delete("1.0", tk.END)
        caja.insert("1.0", texto)
        caja.configure(state="disabled")

    def _mostrar_resultados_calculos(self, res: dict) -> None:
        tablas = [self._tree_ranking, self._tree_desc, self._tree_anova, self._tree_tukey, self._tree_letras_tukey,
                  self._tree_lsd, self._tree_letras_lsd]
        if res.get("forzar_semilla"):
            self._var_unidad_calc.set("Por semilla")
            self._cmb_unidad_calc.configure(state="disabled")
        else:
            self._cmb_unidad_calc.configure(state="disabled" if res.get("variable") == "Germinó" else "readonly")
            if res.get("variable") == "Germinó":
                self._var_unidad_calc.set("Por caja (promedio de sus semillas)")

        self._actualizar_tarjetas_calc(res)
        if not res.get("ok"):
            for t in tablas:
                self._llenar_tabla(t, None)
            self._escribir(self._txt_conclusion_anova, res.get("mensaje") or "Sin datos.")
            self._escribir(self._txt_pares, "")
            return

        an, u, u2, alfa = res.get("anova"), res["u"], res["u2"], res["alfa"]
        m = res["metricas"]
        germ = m["% germinación global"]
        cabecera = (f"{m['Semillas analizadas']:,} semillas · {m['Cajas']} cajas · {m['Grupos comparados']} grupos · "
                    f"{list(m.values())[3]} datos ({res['unidad_txt']}) · germinación global "
                    f"{'—' if pd.isna(germ) else f'{germ:.1f} %'} · {m['Registros descartados']} registros descartados · "
                    f"comparando por {res['agrupacion'].lower()}")
        if an:
            dictamen = (("✔ Sí hay" if an["significativo"] else "✘ No hay") +
                        f" diferencias significativas en {res['nombre_var'].lower()} entre los grupos "
                        f"(F calculada = {an['F calculada']:.3f} {'>' if an['significativo'] else '≤'} "
                        f"F crítica = {an['F crítica']:.3f}; GL = {an['GL entre']} y {an['GL error']}; "
                        f"p = {an['p-valor']:.4f}; α = {alfa}).")
        else:
            dictamen = "Sin ANOVA."
        self._escribir(self._txt_conclusion_anova, f"{dictamen}\n{cabecera}")

        self._llenar_tabla(self._tree_ranking, res["ranking"], dec=2)
        self._escribir(self._txt_pares, "\n".join(f"• {f.replace('**', '')}" for f in res["frases"]) if res["frases"]
                       else "• Ningún par de grupos difiere según Tukey.")
        self._llenar_tabla(self._tree_desc, res["tabla_descriptivos"])
        self._llenar_tabla(self._tree_anova, res["tabla_anova"])
        if an:
            self._nota_anova.configure(text=(
                f"¿Al menos un grupo tiene un promedio distinto? Si F calculada > F crítica (o p < α) sí. "
                f"F crítica = F({1 - alfa:.2f}; {an['GL entre']}, {an['GL error']}) = {an['F crítica']:.3f}. "
                f"GL entre grupos = k − 1 = {an['GL entre']}; GL del error = n − k = {an['GL error']}."))
        tk_, ls_ = res["tukey"], res["lsd"]
        if not tk_.empty:
            self._nota_tukey.configure(text=(
                f"Si |A − B| > DHS (o q calculada > q crítica), los dos grupos difieren. q crítica (α = {alfa}) = "
                f"{tk_['q crítica'].iloc[0]:.3f} · k = {an['k']} grupos · GL del error = {an['GL error']} · "
                f"CM del error = {an['CM error']:.3f} {u2}. DHS = q·√(CMe/2·(1/nA + 1/nB))."))
            if not ls_.empty:
                self._nota_lsd.configure(text=(
                    f"Si |A − B| > LSD (o t calculada > t crítica), los dos grupos difieren. t crítica (α = {alfa}, "
                    f"dos colas) = {ls_['t crítica'].iloc[0]:.3f} · GL del error = {an['GL error']}. "
                    f"LSD = t·√(CMe·(1/nA + 1/nB)). LSD no corrige por hacer muchas comparaciones."
                    + ("" if an["significativo"] else "  ⚠ El ANOVA no es significativo: no reportes estas diferencias.")))
        else:
            self._nota_tukey.configure(text="Se necesitan al menos 2 grupos con 2 o más datos.")
            self._nota_lsd.configure(text="Se necesitan al menos 2 grupos con 2 o más datos.")
        self._llenar_tabla(self._tree_tukey, res["tabla_tukey"])
        self._llenar_tabla(self._tree_letras_tukey,
                           res["letras_tukey"].rename(columns={"Media": f"Promedio ({u})"}) if an else None)
        self._llenar_tabla(self._tree_lsd, res["tabla_lsd"])
        self._llenar_tabla(self._tree_letras_lsd,
                           res["letras_lsd"].rename(columns={"Media": f"Promedio ({u})"}) if an else None)

    def _actualizar_tarjetas_calc(self, res: dict) -> None:
        if not hasattr(self, "_tarjetas_calc"):
            return
        an = res.get("anova") if res.get("ok") else None
        def poner(clave, valor, sub="", color=None):
            card, lv, ls = self._tarjetas_calc[clave]
            lv.configure(text=valor, text_color=color or UI_COLORS["text_primary"])
            ls.configure(text=sub)
        if not an:
            for k in self._tarjetas_calc:
                poner(k, "—", "Se necesitan 2 grupos con 2 o más datos" if k == "anova" else "")
            return
        u, u2, alfa = res["u"], res["u2"], res["alfa"]
        sig = an["significativo"]
        color = UI_COLORS["accent_green"] if sig else UI_COLORS["accent_red"]
        poner("anova", f"{an['F calculada']:.2f} {'>' if sig else '≤'} {an['F crítica']:.2f}",
              ("✔ Sí hay diferencias" if sig else "✘ No hay diferencias") + f" (α = {alfa})", color)
        poner("p", f"{an['p-valor']:.4f}", f"{'menor' if an['p-valor'] < alfa else 'mayor o igual'} que α = {alfa}", color)
        poner("gl", f"{an['GL entre']} y {an['GL error']}", f"entre grupos (k − 1) y error (n − k) · n = {an['n total']}")
        poner("cme", f"{an['CM error']:.3f}", f"{u2} · varianza dentro de los grupos")
        for clave, tabla, crit, margen in [("tukey", res["tukey"], "q crítica", "DHS"),
                                           ("lsd", res["lsd"], "t crítica", "LSD")]:
            if tabla is None or tabla.empty:
                poner(clave, "—")
                continue
            n_sig = int((tabla["Resultado"] == AN.M.SIG).sum())
            mn, mx = tabla[margen].min(), tabla[margen].max()
            rango = f"{mn:.2f}" if abs(mx - mn) < 0.005 else f"{mn:.2f}–{mx:.2f}"
            poner(clave, f"{crit.split()[0]} crít = {tabla[crit].iloc[0]:.3f}",
                  f"{margen} = {rango} {u} · {n_sig} de {len(tabla)} pares difieren")

    def _exportar_reporte_calculos(self) -> None:
        res = self._ultimo_resultado_stats
        if not res or not res.get("ok"):
            messagebox.showwarning("Sin análisis", (res or {}).get("mensaje") or
                                   "Todavía no hay un análisis para exportar.", parent=self)
            return
        var = res["variable"].replace("í", "i").replace("ó", "o")
        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar reporte de cálculos en PDF",
            defaultextension=".pdf",
            initialfile=f"Reporte_Calculos_{var}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
        )
        if not ruta:
            return
        op = self._opciones_descartes()
        descartes = [("semillas no germinadas: " + ("excluidas del promedio" if op["no_germinadas"] == "excluir"
                                                    else f"incluidas con 0 {res['unidad_long']}"))]
        if op["ceros_germinadas"]:
            descartes.append("germinadas con valor 0 excluidas")
        if op["atipicos"]:
            descartes.append("atípicos de cada grupo excluidos (1.5×IQR)")
        if op["cajas_vacias"]:
            descartes.append("cajas vacías excluidas")
        config = {
            "fuente": self._var_fuente_datos.get().replace("MySQL", "Base de datos"),
            "archivos": ", ".join(os.path.basename(a) for a in self._archivos_excel_cargados)
            if not self._var_fuente_datos.get().startswith("Solo MySQL") else "",
            "variedad": self._var_filtro_var_calc.get(),
            "descartes": "; ".join(descartes),
            "proyecto": AN.resumen_proyecto(res.get("df_filtrado"), self._var_fuente_datos.get().replace("MySQL", "Base de datos")),
        }
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            REP.generar_pdf(ruta, res, config, TTK_PALETTES["light"])
            messagebox.showinfo("✅ Reporte generado", f"Reporte de cálculos guardado en:\n{ruta}", parent=self)
        except ImportError:
            messagebox.showerror("Librería faltante", "Instala reportlab ejecutando en tu terminal:\npip install reportlab",
                                 parent=self)
        except Exception as e:
            messagebox.showerror("Error al exportar", f"No se pudo generar el reporte PDF:\n{e}", parent=self)
        finally:
            self.configure(cursor="")

    # -----------------------------------------------------------------------
    # Página 3: Tablero y Gráficas (las gráficas del dashboard, en matplotlib)
    # -----------------------------------------------------------------------

    def _construir_seccion_tablero(self, padre: ctk.CTkFrame) -> None:
        padre.grid_columnconfigure((0, 1, 2, 3), weight=1)
        padre.grid_rowconfigure(2, weight=1)
        redibujar = lambda *_: self._dibujar_grafica_tablero()

        ctrl_graf = ctk.CTkFrame(padre, fg_color=UI_COLORS["bg_card"], corner_radius=10,
                                 border_width=1, border_color=UI_COLORS["border"])
        ctrl_graf.grid(row=0, column=0, columnspan=4, pady=(0, 8), sticky="ew")

        fila1 = ctk.CTkFrame(ctrl_graf, fg_color="transparent")
        fila1.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkLabel(fila1, text="📊 Tipo de Gráfica:", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                     text_color=UI_COLORS["text_primary"]).pack(side="left", padx=(6, 6))
        self._var_tipo_grafica = tk.StringVar(value=GR.TIPOS[0])
        ctk.CTkComboBox(fila1, variable=self._var_tipo_grafica, values=GR.TIPOS, width=390, state="readonly",
                        command=redibujar).pack(side="left", padx=6)
        self._etiqueta_control(fila1, "  Letras:")
        self._var_letras_graf = tk.StringVar(value="Tukey")
        ctk.CTkSegmentedButton(fila1, values=["Tukey", "LSD"], variable=self._var_letras_graf, width=110,
                               command=redibujar).pack(side="left", padx=4)
        self._etiqueta_control(fila1, "  Barras de error:")
        self._var_error_graf = tk.StringVar(value="Desviación estándar")
        ctk.CTkComboBox(fila1, variable=self._var_error_graf,
                        values=["Desviación estándar", "Error estándar (SEM)", "IC 95%"], width=170,
                        state="readonly", command=redibujar).pack(side="left", padx=4)
        ctk.CTkButton(fila1, text="🗂️ Varias gráficas → PDF / Excel", width=200, fg_color=UI_COLORS["accent_indigo"],
                      text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._exportar_varias_graficas).pack(side="right", padx=6)
        ctk.CTkButton(fila1, text="💾 Guardar gráfica (PNG/PDF/Excel)", width=210, fg_color=UI_COLORS["accent_green"],
                      hover_color=UI_COLORS["accent_hover"], text_color="#ffffff",
                      font=ctk.CTkFont(size=11, weight="bold"),
                      command=self._guardar_grafica_archivo).pack(side="right", padx=6)

        # Opciones de la gráfica de tendencia
        fila_t = ctk.CTkFrame(ctrl_graf, fg_color="transparent")
        fila_t.pack(fill="x", padx=10, pady=(2, 2))
        self._etiqueta_control(fila_t, "📈 Tendencia (gráfica 11) — Eje X:")
        self._var_eje_tend = tk.StringVar(value=GR.EJES_TENDENCIA[0])
        ctk.CTkComboBox(fila_t, variable=self._var_eje_tend, values=GR.EJES_TENDENCIA, width=200, state="readonly",
                        command=redibujar).pack(side="left", padx=(0, 10))
        self._etiqueta_control(fila_t, "Series:")
        self._var_series_tend = tk.StringVar(value="Variedad")
        self._series_elegidas: Optional[set] = None

        def _series_cambiadas(_=None):
            self._series_elegidas = None
            self._lbl_series_tend.configure(text="todas")
            self._dibujar_grafica_tablero()
        ctk.CTkComboBox(fila_t, variable=self._var_series_tend, values=list(GR.SERIES_TENDENCIA), width=170,
                        state="readonly", command=_series_cambiadas).pack(side="left", padx=(0, 6))
        ctk.CTkButton(fila_t, text="Elegir series…", width=110, height=26, fg_color=UI_COLORS["bg_secondary"],
                      text_color=UI_COLORS["text_primary"], hover_color=UI_COLORS["border"],
                      font=ctk.CTkFont(size=11), command=self._elegir_series_tendencia).pack(side="left", padx=(0, 6))
        self._lbl_series_tend = ctk.CTkLabel(fila_t, text="todas", font=ctk.CTkFont(size=11),
                                             text_color=UI_COLORS["text_secondary"])
        self._lbl_series_tend.pack(side="left")

        # Filtros rápidos: se suman a los de 'Cálculos y Modelos' (variable, comparación y descartes)
        fila2 = ctk.CTkFrame(ctrl_graf, fg_color="transparent")
        fila2.pack(fill="x", padx=10, pady=(4, 10))
        self._etiqueta_control(fila2, "🔍 Variedad:")
        self._var_filtro_var_tablero = tk.StringVar(value="Todas las variedades")
        self._cmb_var_tablero = ctk.CTkComboBox(fila2, variable=self._var_filtro_var_tablero,
                                                values=["Todas las variedades"], width=190, state="readonly",
                                                command=redibujar)
        self._cmb_var_tablero.pack(side="left", padx=(0, 10))
        self._etiqueta_control(fila2, "Tratamiento:")
        self._var_filtro_trat_tablero = tk.StringVar(value="Todos los tratamientos")
        self._cmb_trat_tablero = ctk.CTkComboBox(fila2, variable=self._var_filtro_trat_tablero,
                                                 values=["Todos los tratamientos"], width=170, state="readonly",
                                                 command=redibujar)
        self._cmb_trat_tablero.pack(side="left", padx=(0, 10))
        self._etiqueta_control(fila2, "📅 Desde:")
        self._var_fecha_desde = tk.StringVar(value="Todas")
        self._cmb_desde = ctk.CTkComboBox(fila2, variable=self._var_fecha_desde, values=["Todas"], width=115,
                                          state="readonly", command=redibujar)
        self._cmb_desde.pack(side="left", padx=(0, 6))
        self._etiqueta_control(fila2, "Hasta:")
        self._var_fecha_hasta = tk.StringVar(value="Todas")
        self._cmb_hasta = ctk.CTkComboBox(fila2, variable=self._var_fecha_hasta, values=["Todas"], width=115,
                                          state="readonly", command=redibujar)
        self._cmb_hasta.pack(side="left", padx=(0, 10))
        self._etiqueta_control(fila2, "💨 Gas:")
        self._var_filtro_gas_tablero = tk.StringVar(value="Todos los gases")
        self._cmb_gas_tablero = ctk.CTkComboBox(fila2, variable=self._var_filtro_gas_tablero,
                                                values=["Todos los gases"], width=140, state="readonly",
                                                command=redibujar)
        self._cmb_gas_tablero.pack(side="left", padx=(0, 10))
        self._lbl_info_tablero = ctk.CTkLabel(fila2, text="", font=ctk.CTkFont(size=11),
                                              text_color=UI_COLORS["text_secondary"])
        self._lbl_info_tablero.pack(side="left", padx=(6, 0))

        self._lbl_kpi_total = self._crear_kpi_card(padre, 1, 0, "🌱 Semillas Analizadas", "0", UI_COLORS["accent_indigo"])
        self._lbl_kpi_germ = self._crear_kpi_card(padre, 1, 1, "✓ Germinación Global", "0 (0%)", UI_COLORS["accent_green"])
        self._lbl_kpi_lotes = self._crear_kpi_card(padre, 1, 2, "📦 Cajas / Lotes", "0", UI_COLORS["accent_amber"])
        self._lbl_kpi_fval = self._crear_kpi_card(padre, 1, 3, "🧮 ANOVA (F / p-valor)", "—", UI_COLORS["accent_red"])

        self._chart_container = ctk.CTkFrame(padre, fg_color=UI_COLORS["bg_card"], corner_radius=12,
                                             border_width=1, border_color=UI_COLORS["border"])
        self._chart_container.grid(row=2, column=0, columnspan=4, pady=(8, 0), sticky="nsew")
        self._chart_container.grid_columnconfigure(0, weight=1)
        self._chart_container.grid_rowconfigure(1, weight=1)
        self._canvas_widget = None

    def _crear_kpi_card(self, padre, row: int, col: int, titulo: str, valor: str, color_borde) -> ctk.CTkLabel:
        padx = (0, 6) if col == 0 else ((6, 6) if col in (1, 2) else (6, 0))
        card = ctk.CTkFrame(padre, fg_color=UI_COLORS["bg_card"], corner_radius=10, border_width=2,
                            border_color=color_borde)
        card.grid(row=row, column=col, padx=padx, pady=4, sticky="nsew")
        ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                     text_color=UI_COLORS["text_secondary"]).pack(anchor="w", padx=14, pady=(10, 2))
        lbl_val = ctk.CTkLabel(card, text=valor, font=ctk.CTkFont(family="Segoe UI", size=19, weight="bold"),
                               text_color=UI_COLORS["text_primary"])
        lbl_val.pack(anchor="w", padx=14, pady=(0, 10))
        return lbl_val

    def _actualizar_filtros_tablero(self, df: pd.DataFrame) -> None:
        if not hasattr(self, "_cmb_var_tablero"):
            return
        for cmb, var, todos, valores in [
            (self._cmb_var_tablero, self._var_filtro_var_tablero, "Todas las variedades", AN.variedades(df)),
            (self._cmb_trat_tablero, self._var_filtro_trat_tablero, "Todos los tratamientos", AN.tratamientos(df)),
            (self._cmb_desde, self._var_fecha_desde, "Todas", AN.fechas(df)),
            (self._cmb_hasta, self._var_fecha_hasta, "Todas", AN.fechas(df)),
            (self._cmb_gas_tablero, self._var_filtro_gas_tablero, "Todos los gases",
             sorted(g for g in df["Gas"].dropna().astype(str).unique() if g != "Sin dato") if "Gas" in df else []),
        ]:
            opciones = [todos] + valores
            cmb.configure(values=opciones)
            if var.get() not in opciones:
                var.set(todos)

    def _filtros_tablero(self) -> dict:
        f = {}
        v = self._var_filtro_var_tablero.get()
        if v and v != "Todas las variedades":
            f["variedades_sel"] = [v]
        t = self._var_filtro_trat_tablero.get()
        if t and t != "Todos los tratamientos":
            f["tratamiento"] = t
        if self._var_fecha_desde.get() not in ("", "Todas"):
            f["desde"] = self._var_fecha_desde.get()
        if self._var_fecha_hasta.get() not in ("", "Todas"):
            f["hasta"] = self._var_fecha_hasta.get()
        g = self._var_filtro_gas_tablero.get()
        if g and g != "Todos los gases":
            f["gas"] = g
        return f

    def _opciones_tendencia(self) -> dict:
        return {"eje": self._var_eje_tend.get(), "series": self._var_series_tend.get(),
                "seleccion": sorted(self._series_elegidas) if self._series_elegidas else None}

    def _elegir_series_tendencia(self) -> None:
        col = GR.SERIES_TENDENCIA.get(self._var_series_tend.get())
        df = getattr(self, "_df_activo", pd.DataFrame())
        if not col or df.empty or col not in df:
            messagebox.showinfo("Series", "Elige primero qué serie usar (Variedad, Tipo de gas…).", parent=self)
            return
        valores = sorted(v for v in df[col].dropna().astype(str).unique() if v not in ("Sin dato", "nan"))
        top = ctk.CTkToplevel(self)
        top.title("Series de la tendencia")
        top.configure(fg_color=UI_COLORS["bg_primary"])
        centrar_ventana(top, 380, min(640, 150 + 34 * len(valores)))
        top.transient(self)
        top.grab_set()
        marco = ctk.CTkScrollableFrame(top, fg_color=UI_COLORS["bg_card"])
        marco.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        variables = {}
        for v in valores:
            var = tk.BooleanVar(value=self._series_elegidas is None or v in self._series_elegidas)
            ctk.CTkCheckBox(marco, text=v, variable=var, font=ctk.CTkFont(size=11)).pack(anchor="w", pady=3, padx=6)
            variables[v] = var

        def aplicar():
            elegidas = {v for v, var in variables.items() if var.get()}
            self._series_elegidas = None if len(elegidas) in (0, len(valores)) else elegidas
            self._lbl_series_tend.configure(text="todas" if self._series_elegidas is None
                                            else f"{len(elegidas)} de {len(valores)}")
            top.destroy()
            if not self._var_tipo_grafica.get().startswith("11"):
                self._var_tipo_grafica.set(GR.TIPOS[10])
            self._dibujar_grafica_tablero()
        ctk.CTkButton(top, text="Aplicar", fg_color=UI_COLORS["accent_green"], text_color="#ffffff",
                      command=aplicar).pack(fill="x", padx=12, pady=(0, 12))

    def _actualizar_kpis_tablero(self, res: Optional[dict] = None) -> None:
        res = res if res is not None else self._ultimo_resultado_stats
        if not res or "metricas" not in res:
            for lbl in (self._lbl_kpi_total, self._lbl_kpi_germ, self._lbl_kpi_lotes):
                lbl.configure(text="0")
            self._lbl_kpi_fval.configure(text="Sin datos")
            return
        m = res["metricas"]
        d = res["df_semilla"]
        col = "Germinó (conteo)" if "Germinó (conteo)" in d.columns else "Germinó"
        germ = int((pd.to_numeric(d[col], errors="coerce") == 1).sum()) if not d.empty else 0
        pct = m["% germinación global"]
        self._lbl_kpi_total.configure(text=f"{m['Semillas analizadas']:,}")
        self._lbl_kpi_germ.configure(text=f"{germ:,} ({0 if pd.isna(pct) else pct:.1f}%)")
        self._lbl_kpi_lotes.configure(text=f"{m['Cajas']:,}")
        an = res.get("anova")
        if an:
            estado = "✔ Sig." if an["significativo"] else "✘ No sig."
            self._lbl_kpi_fval.configure(text=f"F={an['F calculada']:.2f} (p={an['p-valor']:.3g}) {estado}")
        else:
            self._lbl_kpi_fval.configure(text="Insuf. grupos")

    def _dibujar_grafica_tablero(self) -> None:
        if not hasattr(self, "_chart_container") or not hasattr(self, "_df_activo"):
            return
        filtros = self._filtros_tablero()
        if filtros:
            try:
                res = self._analizar(self._df_activo, filtros)
            except Exception as e:
                res = {"ok": False, "mensaje": f"No se pudo hacer el análisis: {e}"}
            self._lbl_info_tablero.configure(text="Filtros activos (se suman a Cálculos y Modelos).")
        else:
            res = self._ultimo_resultado_stats or {"ok": False, "mensaje": "Sin datos."}
            self._lbl_info_tablero.configure(text="Misma configuración que Cálculos y Modelos.")
        self._actualizar_kpis_tablero(res)
        self._res_tablero = res

        for child in self._chart_container.winfo_children():
            child.destroy()
        p = TTK_PALETTES[self._modo_actual]
        try:
            fig = GR.dibujar(self._var_tipo_grafica.get(), res, p, self._var_letras_graf.get(),
                             self._var_error_graf.get(), self._opciones_tendencia())
        except Exception as e:
            fig = GR._mensaje(p, f"No se pudo dibujar esta gráfica: {e}")
        if res.get("ok"):
            FD.fondo_en_figura(fig, "tablero", self._modo_actual)  # solo en pantalla, no en lo exportado
        self._fig_actual = fig

        alto_px = int(fig.get_figheight() * fig.dpi)
        if alto_px > 620:  # gráficas con muchos pares: se desplazan verticalmente
            contenedor = ctk.CTkScrollableFrame(self._chart_container, fg_color=p["bg_card"])
            contenedor.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
            contenedor.grid_columnconfigure(0, weight=1)
            canvas = FigureCanvasTkAgg(fig, master=contenedor)
            canvas.get_tk_widget().configure(height=alto_px)
            canvas.get_tk_widget().grid(row=0, column=0, sticky="ew")
        else:
            canvas = FigureCanvasTkAgg(fig, master=self._chart_container)
            canvas.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        canvas.draw()
        texto = GR.descripcion_cuartiles(res) if (self._var_tipo_grafica.get().startswith("9") and res.get("ok")) else ""
        if texto:
            caja = ctk.CTkTextbox(self._chart_container, height=118, fg_color=UI_COLORS["bg_secondary"],
                                  text_color=UI_COLORS["text_primary"], wrap="word",
                                  font=ctk.CTkFont(family="Segoe UI", size=11))
            caja.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))
            caja.insert("1.0", "Descripción de los datos\n" + texto)
            caja.tag_add("titulo", "1.0", "1.end")
            caja.configure(state="disabled")
        barra = ctk.CTkFrame(self._chart_container, fg_color="transparent")
        barra.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 0))
        toolbar = NavigationToolbar2Tk(canvas, barra, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(side="left")
        self._canvas_widget = canvas

    def _figura_para_exportar(self, tipo: str, res: dict):
        """La misma gráfica, en colores claros (fondo blanco) para documentos."""
        return GR.dibujar(tipo, res, TTK_PALETTES["light"], self._var_letras_graf.get(),
                          self._var_error_graf.get(), self._opciones_tendencia())

    def _proyecto_actual(self, res: dict) -> pd.DataFrame:
        fuente = self._var_fuente_datos.get().replace("MySQL", "Base de datos")
        return AN.resumen_proyecto(res.get("df_filtrado") if res else None, fuente)

    def _guardar_grafica_archivo(self) -> None:
        if self._fig_actual is None:
            messagebox.showwarning("Sin gráfica", "Primero genera una gráfica en el tablero.", parent=self)
            return
        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar gráfica actual",
            defaultextension=".png",
            initialfile=f"Grafica_BioGrow_{datetime.now().strftime('%Y%m%d_%H%M')}.png",
            filetypes=[("Imagen PNG (Alta resolución)", "*.png"), ("Documento PDF", "*.pdf"), ("Imagen JPG", "*.jpg"),
                       ("Excel (gráfica + datos)", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            if ruta.lower().endswith(".xlsx"):
                res = getattr(self, "_res_tablero", None) or {}
                tipo = self._var_tipo_grafica.get()
                EX.exportar_excel(ruta, [(tipo, self._figura_para_exportar(tipo, res))], res,
                                  self._proyecto_actual(res), incluir_tablas=True)
                messagebox.showinfo("✅ Gráfica Guardada", f"Excel con la gráfica, sus datos y los cálculos:\n{ruta}",
                                    parent=self)
            else:
                self._fig_actual.savefig(ruta, dpi=300, bbox_inches="tight", facecolor=self._fig_actual.get_facecolor())
                messagebox.showinfo("✅ Gráfica Guardada", f"Gráfica exportada en alta resolución (300 DPI):\n{ruta}",
                                    parent=self)
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar la gráfica:\n{e}", parent=self)

    def _exportar_varias_graficas(self) -> None:
        res = getattr(self, "_res_tablero", None)
        if not res or not res.get("ok"):
            messagebox.showwarning("Sin datos", "No hay un análisis para graficar con la selección actual.", parent=self)
            return
        top = ctk.CTkToplevel(self)
        top.title("Exportar varias gráficas")
        top.configure(fg_color=UI_COLORS["bg_primary"])
        centrar_ventana(top, 520, 640)
        top.transient(self)
        top.grab_set()
        ctk.CTkLabel(top, text="Elige las gráficas que irán en un solo archivo",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=UI_COLORS["text_primary"]).pack(
            anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(top, text="Usan la configuración actual (variable, comparación, filtros, letras y tendencia).",
                     font=ctk.CTkFont(size=11), text_color=UI_COLORS["text_secondary"]).pack(anchor="w", padx=16)
        marco = ctk.CTkScrollableFrame(top, fg_color=UI_COLORS["bg_card"], height=330)
        marco.pack(fill="both", expand=True, padx=16, pady=8)
        por_defecto = {"1", "2", "9", "10", "12"}
        elegidas = {}
        for t in GR.TIPOS:
            var = tk.BooleanVar(value=t.split(".")[0] in por_defecto)
            ctk.CTkCheckBox(marco, text=t, variable=var, font=ctk.CTkFont(size=11)).pack(anchor="w", pady=3, padx=6)
            elegidas[t] = var
        opciones = ctk.CTkFrame(top, fg_color="transparent")
        opciones.pack(fill="x", padx=16)
        self._etiqueta_control(opciones, "Formato:")
        var_formato = tk.StringVar(value="PDF")
        ctk.CTkSegmentedButton(opciones, values=["PDF", "Excel", "PDF y Excel"], variable=var_formato).pack(side="left")
        var_tablas = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(top, text="Incluir en el Excel las tablas de cálculos y los datos de cada grupo",
                        variable=var_tablas, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=16, pady=(8, 0))

        def exportar():
            tipos = [t for t, v in elegidas.items() if v.get()]
            if not tipos:
                messagebox.showwarning("Sin gráficas", "Marca al menos una gráfica.", parent=top)
                return
            formato = var_formato.get()
            ext = ".xlsx" if formato == "Excel" else ".pdf"
            ruta = filedialog.asksaveasfilename(
                parent=top, title="Guardar gráficas", defaultextension=ext,
                initialfile=f"Graficas_BioGrow_{datetime.now().strftime('%Y%m%d_%H%M')}{ext}",
                filetypes=[("Excel", "*.xlsx")] if ext == ".xlsx" else [("PDF", "*.pdf")])
            if not ruta:
                return
            top.configure(cursor="watch")
            top.update_idletasks()
            try:
                figuras = [(t, self._figura_para_exportar(t, res)) for t in tipos]
                proyecto = self._proyecto_actual(res)
                hechos = []
                if formato in ("PDF", "PDF y Excel"):
                    ruta_pdf = os.path.splitext(ruta)[0] + ".pdf"
                    EX.exportar_pdf(ruta_pdf, figuras, res, proyecto)
                    hechos.append(ruta_pdf)
                if formato in ("Excel", "PDF y Excel"):
                    ruta_xl = os.path.splitext(ruta)[0] + ".xlsx"
                    EX.exportar_excel(ruta_xl, figuras, res, proyecto, var_tablas.get())
                    hechos.append(ruta_xl)
            except Exception as e:
                top.configure(cursor="")
                messagebox.showerror("Error al exportar", f"No se pudieron exportar las gráficas:\n{e}", parent=top)
                return
            top.destroy()
            messagebox.showinfo("✅ Gráficas exportadas", f"{len(tipos)} gráfica(s) guardadas en:\n" + "\n".join(hechos),
                                parent=self)
        ctk.CTkButton(top, text="💾 Exportar", fg_color=UI_COLORS["accent_green"], text_color="#ffffff",
                      font=ctk.CTkFont(size=12, weight="bold"), command=exportar).pack(fill="x", padx=16, pady=12)

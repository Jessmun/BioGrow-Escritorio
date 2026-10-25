"""
interfaz.py — Interfaz gráfica de BioGrow Desktop
- Colores duales adaptativos (Modo Claro / Modo Oscuro sin zonas oscuras residuales).
- Selector de páginas en la esquina superior izquierda.
- Botón de cambio de tema (reemplaza al switch).
- Modo Administrador con Edición y Eliminación de registros y lotes.
- Exportación a Excel (.xlsx) y PDF (.pdf) con diseño profesional.
"""
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
)

import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from analytics_engine import (
    cargar_archivos_excel,
    convertir_mysql_a_df,
    combinar_fuentes,
    preparar_datos_analisis,
    ejecutar_suite_completa,
    generar_texto_conclusion,
    SIG,
)
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
TIPOS_SEMILLA = [
    "Semilla de maíz blanco",
    "Semilla de maíz rosa",
    "Semilla de maíz cacahuazintle",
    "Semilla de maíz criollo",
    "Semilla de lenteja",
]
ADMIN_PIN = "admin123"


class BioGrowApp(ctk.CTk):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self._modo_actual = "dark"
        self._admin_activo = False

        self._filas_captura: list = []
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
            self._tree_historial.tag_configure("germino_si", foreground=p["accent_green"], background=p["bg_card"])
            self._tree_historial.tag_configure("germino_no", foreground=p["accent_red"], background=p["bg_card"])
            self._tree_historial.tag_configure("alt_si", background=p["row_alt"], foreground=p["accent_green"])
            self._tree_historial.tag_configure("alt_no", background=p["row_alt"], foreground=p["accent_red"])

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
        ).grid(row=0, column=0, columnspan=2, padx=(14, 8), pady=(10, 2), sticky="w")

        self._entry_lote = ctk.CTkEntry(
            frame,
            placeholder_text="Ej: Lote 16-Mayo / Criollo M1",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_lote.grid(row=1, column=0, columnspan=2, padx=(14, 8), pady=(0, 8), sticky="ew")

        ctk.CTkLabel(
            frame,
            text="Variedad de Semilla",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=2, padx=(8, 14), pady=(10, 2), sticky="w")

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
        self._combo_semilla.grid(row=1, column=2, padx=(8, 14), pady=(0, 8), sticky="ew")

        # --- FILA 2 y 3: Fecha | Tiempo de Tratamiento | Temperatura (°C) ---
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
        ).grid(row=2, column=1, padx=6, pady=(4, 2), sticky="w")

        self._entry_tratamiento = ctk.CTkEntry(
            frame,
            placeholder_text="Ej: 45s, Control, T-1M",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_tratamiento.grid(row=3, column=1, padx=6, pady=(0, 12), sticky="ew")

        ctk.CTkLabel(
            frame,
            text="🌡️ Temperatura (°C)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=2, column=2, padx=(6, 14), pady=(4, 2), sticky="w")

        self._entry_temperatura = ctk.CTkEntry(
            frame,
            placeholder_text="Ej: 35, 40.5 o Ambiente",
            fg_color=UI_COLORS["input_bg"],
            border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            placeholder_text_color=UI_COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self._entry_temperatura.insert(0, "Ambiente")
        self._entry_temperatura.grid(row=3, column=2, padx=(6, 14), pady=(0, 12), sticky="ew")

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

        ctk.CTkLabel(
            contenedor_tabla,
            text="📊  Tabla de Captura — 27 Muestras (Doble clic o Enter para editar)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, pady=(0, 4), sticky="w")

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
            height=18,
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

        for i in range(1, NUM_MUESTRAS + 1):
            self._filas_captura.append({"raiz": None, "tallo": None, "hojas": None})
            tag = "row_alt" if i % 2 == 0 else "row_norm"
            self._tree_captura.insert("", "end", iid=str(i), values=(i, "", "", "", ""), tags=(tag,))

    def _refrescar_colores_captura(self) -> None:
        for idx in range(NUM_MUESTRAS):
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
            if siguiente < NUM_MUESTRAS:
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
        elif event.keysym == "Down" and fila_idx < NUM_MUESTRAS - 1:
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

        if nueva_fila < NUM_MUESTRAS:
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
        self._entry_temperatura.insert(0, "Ambiente")
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
        if not hasattr(self, "_sel_variedades"):
            self._sel_variedades = set(TIPOS_SEMILLA)
        if not hasattr(self, "_sel_tratamientos"):
            self._sel_tratamientos = {"Control", "30s", "45s", "60s"}
        if not hasattr(self, "_sel_germino"):
            self._sel_germino = {"1", "0"}

        # 2. Botón desplegable para Variedades (Pasando TIPOS_SEMILLA completos)
        ctk.CTkButton(
            filtros_frame, text="🌱 Variedades...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Variedades", self._sel_variedades, TIPOS_SEMILLA)
        ).grid(row=0, column=1, padx=(0, 6), sticky="ew")

        # 3. Botón desplegable para Tratamientos (Pasando la lista completa de tratamientos)
        ctk.CTkButton(
            filtros_frame, text="⏱️ Tratamientos...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Tratamientos", self._sel_tratamientos, ["Control", "30s", "45s", "60s"])
        ).grid(row=0, column=2, padx=(0, 6), sticky="ew")

        # 4. Botón desplegable para Estado de Germinación
        ctk.CTkButton(
            filtros_frame, text="✓ Germinación...",
            fg_color=UI_COLORS["input_bg"], hover_color=UI_COLORS["border"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"], font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self._abrir_selector_multiple("Germinación", self._sel_germino, [("1", "Germinó"), ("0", "No germinó")])
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

    def _abrir_ventana_graficos_historial(self) -> None:
        """Abre una nueva ventana modal con múltiples gráficas basadas en el historial filtrado actual."""
        if not self._datos_historial_actual:
            messagebox.showwarning("Sin datos", "No hay registros en el historial con los filtros actuales para graficar.", parent=self)
            return

        # Convertir la lista actual del historial filtrado en DataFrame
        df_base = pd.DataFrame(self._datos_historial_actual)
        
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

    def _abrir_selector_multiple(self, titulo: str, conjunto_ref: set, opciones_disponibles: list) -> None:
        """Ventana modal con Checkboxes para selección múltiple de filtros."""
        top = ctk.CTkToplevel(self)
        top.title(f"Seleccionar {titulo}")
        top.geometry("340x380")
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
            
            var = tk.BooleanVar(value=(val_key in conjunto_ref))
            vars_dict[val_key] = var
            
            chk = ctk.CTkCheckBox(frame_chk, text=val_lbl, variable=var)
            chk.pack(anchor="w", pady=4, padx=5)

        def aplicar_seleccion():
            conjunto_ref.clear()
            for k, v in vars_dict.items():
                if v.get():
                    conjunto_ref.add(str(k))
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

        columnas = ("fecha", "lote", "semilla", "muestra", "tratamiento", "temperatura", "hojas", "raiz", "tallo", "resultado")
        
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
            "semilla":     ("Tipo Semilla",   130, "center"),
            "muestra":     ("No.",             42, "center"),
            "tratamiento": ("Tratamiento",     85, "center"),
            "temperatura": ("Temp. (°C)",      80, "center"),
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
        centrar_ventana(modal, 460, 480)
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
        cmb_sem = ctk.CTkComboBox(
            card, values=TIPOS_SEMILLA, state="readonly",
            fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"]
        )
        cmb_sem.set(reg["tipo_semilla"] if reg["tipo_semilla"] in TIPOS_SEMILLA else TIPOS_SEMILLA[0])
        cmb_sem.grid(row=4, column=0, padx=20, pady=(2, 10), sticky="ew")

        # Tratamiento
        ctk.CTkLabel(card, text="Tiempo de Tratamiento:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).grid(row=5, column=0, padx=20, sticky="w")
        ent_trat = ctk.CTkEntry(card, fg_color=UI_COLORS["input_bg"], border_color=UI_COLORS["border"], text_color=UI_COLORS["text_primary"])
        ent_trat.insert(0, reg["tiempo_tratamiento"])
        ent_trat.grid(row=6, column=0, padx=20, pady=(2, 10), sticky="ew")

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

            exito, msg = self.db.actualizar_medicion(
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
        btn_box.grid(row=8, column=0, padx=20, pady=(10, 16), sticky="ew")
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

            # Filtrar por Variedad múltiple
            if hasattr(self, "_sel_variedades") and row["tipo_semilla"] not in self._sel_variedades:
                continue

            # Filtrar por Tratamiento múltiple
            if hasattr(self, "_sel_tratamientos") and row["tiempo_tratamiento"] not in self._sel_tratamientos:
                continue

            # Filtrar por Germinación múltiple
            if hasattr(self, "_sel_germino") and str(row["germino"]) not in self._sel_germino:
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
                    row["tiempo_tratamiento"],
                    row.get("temperatura", "Ambiente"),
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
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Historial Biométrico"

            # Título del reporte
            ws.merge_cells("A1:H1")
            ws["A1"] = "BIOGROW ANALYTICS — REPORTE DE GERMINACIÓN Y BIOMETRÍA DE SEMILLAS"
            ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            ws["A1"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28

            # Subtítulo con fecha y resumen
            total = len(self._datos_historial_actual)
            germ = sum(1 for r in self._datos_historial_actual if int(r["germino"]) == 1)
            pct = (germ / total * 100) if total > 0 else 0
            ws.merge_cells("A2:H2")
            ws["A2"] = f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   Total Muestras: {total}   |   Germinadas: {germ} ({pct:.1f}%)"
            ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="334155")
            ws["A2"].fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 20

            # Encabezados de tabla
            headers = ["Lote Origen", "Variedad de Semilla", "No. Muestra", "Tratamiento", "Cant. Hojas", "Raíz (mm)", "Tallo (mm)", "Estado"]
            ws.append([])
            ws.append(headers)
            ws.row_dimensions[4].height = 24

            header_fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style="thin", color="CBD5E1"),
                right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="CBD5E1"),
                bottom=Side(style="thin", color="CBD5E1"),
            )

            for col_num in range(1, 9):
                c = ws.cell(row=4, column=col_num)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

            alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            si_font  = Font(name="Calibri", size=10, bold=True, color="059669")
            no_font  = Font(name="Calibri", size=10, bold=True, color="DC2626")
            reg_font = Font(name="Calibri", size=10, color="0F172A")

            for idx, r in enumerate(self._datos_historial_actual, start=5):
                estado = "Germinó" if int(r["germino"]) == 1 else "No germinó"
                fila = [
                    r["nombre_lote"],
                    r["tipo_semilla"],
                    int(r["num_muestra"]),
                    r["tiempo_tratamiento"],
                    int(r["cantidad_hojas"]) if r["cantidad_hojas"] is not None else 0,
                    float(r["longitud_raiz"]),
                    float(r["longitud_tallo"]),
                    estado,
                ]
                ws.append(fila)
                ws.row_dimensions[idx].height = 20

                for col_num in range(1, 9):
                    celda = ws.cell(row=idx, column=col_num)
                    celda.border = thin_border
                    if idx % 2 == 1:
                        celda.fill = alt_fill
                    celda.alignment = Alignment(horizontal="left" if col_num in (1, 2) else "center", vertical="center")
                    if col_num in (6, 7):
                        celda.number_format = "0.00"
                    if col_num == 8:
                        celda.font = si_font if int(r["germino"]) == 1 else no_font
                    else:
                        celda.font = reg_font

            anchos = [24, 28, 14, 18, 14, 14, 14, 16]
            for i, ancho in enumerate(anchos, start=1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho

            wb.save(ruta)
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
    # Gestión Híbrida de Datos (MySQL + Archivos Excel de dashboardst.py)
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

        self._df_excel_cache, self._advertencias_excel = cargar_archivos_excel(self._archivos_excel_cargados)
        n_arch = len(self._archivos_excel_cargados)
        n_filas = len(self._df_excel_cache)

        if hasattr(self, "_lbl_estado_excel"):
            self._lbl_estado_excel.configure(text=f"📂 {n_arch} Excel ({n_filas} semillas)")

        self._var_fuente_datos.set("MySQL + Excel (Combinados)" if not self._df_excel_cache.empty else "Solo MySQL")
        self._ejecutar_analisis_ui()

        msg = f"Se cargaron {n_arch} archivo(s) Excel con {n_filas} mediciones reconocidas."
        if self._advertencias_excel:
            msg += "\n\nNotas del lector:\n• " + "\n• ".join(self._advertencias_excel[:6])
        messagebox.showinfo("Archivos Excel Cargados", msg, parent=self)

    def _limpiar_excels_externos(self) -> None:
        self._archivos_excel_cargados.clear()
        self._df_excel_cache = pd.DataFrame()
        self._advertencias_excel.clear()
        if hasattr(self, "_lbl_estado_excel"):
            self._lbl_estado_excel.configure(text="📂 0 Excel cargados")
        self._var_fuente_datos.set("Solo MySQL")
        self._ejecutar_analisis_ui()

    def _obtener_dataframe_activo(self) -> pd.DataFrame:
        registros_db = self.db.obtener_historial_con_id()
        df_mysql = convertir_mysql_a_df(registros_db)
        modo = self._var_fuente_datos.get() if hasattr(self, "_var_fuente_datos") else "Solo MySQL"
        df = combinar_fuentes(df_mysql, self._df_excel_cache, modo)

        # Filtro de variedad si está seleccionado
        if not df.empty and hasattr(self, "_var_filtro_var_calc"):
            f_var = self._var_filtro_var_calc.get()
            if f_var and f_var != "Todas las variedades":
                df = df[df["Variedad"].astype(str).str.contains(f_var, case=False, na=False)].copy()
        return df

    # -----------------------------------------------------------------------
    # Página 2: Cálculos y Modelos (ANOVA, Tukey, LSD, Cuartiles y Descriptivos)
    # -----------------------------------------------------------------------

    def _construir_seccion_calculos(self, padre: ctk.CTkFrame) -> None:
        padre.grid_columnconfigure(0, weight=1)
        padre.grid_rowconfigure(2, weight=1)

        # 1. Barra de Control Estadístico y Fuente de Datos
        ctrl = ctk.CTkFrame(
            padre, fg_color=UI_COLORS["bg_card"],
            corner_radius=12, border_width=1, border_color=UI_COLORS["border"]
        )
        ctrl.grid(row=0, column=0, pady=(0, 8), sticky="ew")
        ctrl.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        # Fila 0: Fuente de datos + Botones Excel
        top_bar = ctk.CTkFrame(ctrl, fg_color="transparent")
        top_bar.grid(row=0, column=0, columnspan=6, padx=16, pady=(12, 6), sticky="ew")

        ctk.CTkLabel(
            top_bar, text="🗄️ Fuente de Datos:",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=UI_COLORS["text_primary"]
        ).pack(side="left", padx=(0, 8))

        self._var_fuente_datos = tk.StringVar(value="Solo MySQL")
        ctk.CTkSegmentedButton(
            top_bar,
            values=["Solo MySQL", "MySQL + Excel (Combinados)", "Solo Excel"],
            variable=self._var_fuente_datos,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            selected_color=UI_COLORS["accent_indigo"],
            command=lambda _: self._ejecutar_analisis_ui(),
        ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            top_bar, text="➕ Cargar / Combinar Excel (.xlsx)",
            height=30, fg_color=UI_COLORS["accent_green"], hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
            command=self._cargar_excels_externos,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            top_bar, text="🧹 Quitar Excel", width=95, height=30,
            fg_color=UI_COLORS["bg_secondary"], text_color=UI_COLORS["text_primary"],
            hover_color=UI_COLORS["border"], font=ctk.CTkFont(size=11),
            command=self._limpiar_excels_externos,
        ).pack(side="left", padx=(0, 10))

        self._lbl_estado_excel = ctk.CTkLabel(
            top_bar, text="📂 0 Excel cargados",
            font=ctk.CTkFont(size=11), text_color=UI_COLORS["text_secondary"]
        )
        self._lbl_estado_excel.pack(side="left", padx=(4, 0))

        ctk.CTkButton(
            top_bar, text="🔄 Recalcular", width=100, height=30,
            fg_color=UI_COLORS["accent_indigo"], text_color="#ffffff",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._ejecutar_analisis_ui,
        ).pack(side="right")

        # Fila 1: Parámetros del modelo estadístico
        param_bar = ctk.CTkFrame(ctrl, fg_color="transparent")
        param_bar.grid(row=1, column=0, columnspan=6, padx=16, pady=(4, 12), sticky="ew")

        # Variable
        ctk.CTkLabel(param_bar, text="Variable:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self._var_variable_calc = tk.StringVar(value="Raíz")
        ctk.CTkComboBox(
            param_bar, variable=self._var_variable_calc,
            values=["Raíz", "Tallo", "Hojas", "Germinó"], width=120, state="readonly",
            command=lambda _: self._ejecutar_analisis_ui()
        ).pack(side="left", padx=(0, 12))

        # Comparar por
        ctk.CTkLabel(param_bar, text="Comparar por:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self._var_agrupacion_calc = tk.StringVar(value="Tratamiento")
        ctk.CTkComboBox(
            param_bar, variable=self._var_agrupacion_calc,
            values=[
                "Tratamiento",
                "Variedad de Semilla",
                "Lote / Hoja",
                "Parte Expuesta (Clase)",
                "Tratamiento + Variedad",
                "Origen (MySQL vs Excel)",
            ],
            width=185, state="readonly",
            command=lambda _: self._ejecutar_analisis_ui()
        ).pack(side="left", padx=(0, 12))

        # Filtro de variedad
        ctk.CTkLabel(param_bar, text="Filtrar Variedad:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self._var_filtro_var_calc = tk.StringVar(value="Todas las variedades")
        ctk.CTkComboBox(
            param_bar, variable=self._var_filtro_var_calc,
            values=["Todas las variedades", "blanco", "rosa", "cacahuazintle", "criollo", "lenteja"],
            width=155, state="readonly",
            command=lambda _: self._ejecutar_analisis_ui()
        ).pack(side="left", padx=(0, 12))

        # Unidad experimental
        ctk.CTkLabel(param_bar, text="Réplica:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self._var_unidad_calc = tk.StringVar(value="Por Caja/Lote (Recomendado)")
        ctk.CTkComboBox(
            param_bar, variable=self._var_unidad_calc,
            values=["Por Caja/Lote (Recomendado)", "Por Semilla Individual"],
            width=195, state="readonly",
            command=lambda _: self._ejecutar_analisis_ui()
        ).pack(side="left", padx=(0, 12))

        # Nivel alfa
        ctk.CTkLabel(param_bar, text="α:", font=ctk.CTkFont(size=11, weight="bold"), text_color=UI_COLORS["text_secondary"]).pack(side="left", padx=(0, 4))
        self._var_alfa_calc = tk.StringVar(value="0.05")
        ctk.CTkComboBox(
            param_bar, variable=self._var_alfa_calc,
            values=["0.05", "0.01"], width=75, state="readonly",
            command=lambda _: self._ejecutar_analisis_ui()
        ).pack(side="left", padx=(0, 12))

        # Checkboxes de limpieza
        self._var_excluir_nogerm = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            param_bar, text="Excluir no germinadas (0 mm)", variable=self._var_excluir_nogerm,
            font=ctk.CTkFont(size=11), checkbox_width=18, checkbox_height=18,
            command=self._ejecutar_analisis_ui
        ).pack(side="left", padx=(0, 10))

        self._var_excluir_atipicos = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            param_bar, text="Quitar atípicos (1.5×IQR)", variable=self._var_excluir_atipicos,
            font=ctk.CTkFont(size=11), checkbox_width=18, checkbox_height=18,
            command=self._ejecutar_analisis_ui
        ).pack(side="left")

        # 2. Banner de dictamen rápido ANOVA
        self._txt_conclusion_anova = ctk.CTkTextbox(
            padre, height=82, fg_color=UI_COLORS["bg_card"],
            border_width=1, border_color=UI_COLORS["border"],
            text_color=UI_COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
            corner_radius=10,
        )
        self._txt_conclusion_anova.grid(row=1, column=0, pady=(0, 8), sticky="ew")

        # 3. Sub-pestañas con las 5 tablas estadísticas de dashboardst.py
        self._tabs_estadistica = ctk.CTkTabview(
            padre,
            fg_color=UI_COLORS["bg_card"],
            segmented_button_fg_color=UI_COLORS["bg_secondary"],
            segmented_button_selected_color=UI_COLORS["accent_green"],
            segmented_button_unselected_color=UI_COLORS["bg_secondary"],
            text_color=UI_COLORS["text_primary"],
            corner_radius=12,
            border_width=1,
            border_color=UI_COLORS["border"],
        )
        self._tabs_estadistica.grid(row=2, column=0, sticky="nsew")

        t_rank = self._tabs_estadistica.add("🏆 1. Ranking, Media, DE y Letras")
        t_anov = self._tabs_estadistica.add("🧮 2. Tabla ANOVA (1 Vía)")
        t_tuk  = self._tabs_estadistica.add("🔠 3. Prueba de Tukey HSD")
        t_lsd  = self._tabs_estadistica.add("🔡 4. Prueba LSD de Fisher")
        t_cuar = self._tabs_estadistica.add("📦 5. Análisis por Cuartiles e IQR")

        self._tree_ranking = self._crear_treeview_estadistico(
            t_rank,
            [("Lugar", 55), ("Grupo", 180), ("n", 60), ("Promedio ± DE", 140),
             ("Varianza", 95), ("CV %", 80), ("% Germ.", 130), ("Tukey", 80), ("LSD", 80)]
        )
        self._tree_anova = self._crear_treeview_estadistico(
            t_anov,
            [("Fuente", 210), ("GL", 65), ("SC", 110), ("CM", 110),
             ("F calc", 95), ("F crít", 95), ("Comparación", 140), ("p-valor", 95), ("Conclusión", 140)]
        )
        self._tree_tukey = self._crear_treeview_estadistico(
            t_tuk,
            [("Grupo 1", 160), ("Grupo 2", 160), ("Conclusión", 120), ("|Diferencia|", 95),
             ("DHS", 95), ("Comparación", 130), ("q calc", 85), ("q crít", 85), ("p ajustada", 95)]
        )
        self._tree_lsd = self._crear_treeview_estadistico(
            t_lsd,
            [("Grupo 1", 160), ("Grupo 2", 160), ("Conclusión", 120), ("|Diferencia|", 95),
             ("LSD", 95), ("Comparación", 130), ("t calc", 85), ("t crít", 85), ("p-valor", 95)]
        )
        self._tree_cuartiles = self._crear_treeview_estadistico(
            t_cuar,
            [("Grupo", 170), ("n", 55), ("Mínimo", 80), ("Q1", 80), ("Mediana", 85),
             ("Q3", 80), ("Máximo", 80), ("IQR", 80), ("Lím. Inf", 85), ("Lím. Sup", 85),
             ("Atípicos", 75), ("Bowley", 80), ("Forma", 160)]
        )

    def _crear_treeview_estadistico(self, contenedor, columnas_cfg: list) -> ttk.Treeview:
        contenedor.grid_columnconfigure(0, weight=1)
        contenedor.grid_rowconfigure(0, weight=1)
        cols = [c[0] for c in columnas_cfg]
        tree = ttk.Treeview(contenedor, columns=cols, show="headings", style="BioGrow.Treeview")
        for nombre, ancho in columnas_cfg:
            tree.heading(nombre, text=nombre, anchor="center")
            tree.column(nombre, width=ancho, minwidth=50, anchor="center", stretch=True)

        vsb = ttk.Scrollbar(contenedor, orient="vertical", command=tree.yview, style="Dark.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(contenedor, orient="horizontal", command=tree.xview, style="Dark.Horizontal.TScrollbar")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=(4, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(4, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        return tree

    def _poblar_treeview_df(self, tree: ttk.Treeview, df: pd.DataFrame, columnas: list) -> None:
        for item in tree.get_children():
            tree.delete(item)
        if df is None or df.empty:
            return
        for idx, (_, row) in enumerate(df.iterrows()):
            vals = []
            for c in columnas:
                v = row.get(c, "")
                if pd.isna(v):
                    vals.append("")
                elif isinstance(v, (int, np.integer)):
                    vals.append(str(int(v)))
                elif isinstance(v, (float, np.floating)):
                    vals.append(f"{v:.4f}" if "p" in c.lower() else f"{v:.3f}")
                else:
                    vals.append(str(v))
            es_par = (idx % 2 == 1)
            texto_fila = " ".join(vals)
            if "✔" in texto_fila:
                tag = "alt_si" if es_par else "germino_si"
            elif "✘" in texto_fila:
                tag = "alt_no" if es_par else "germino_no"
            else:
                tag = "row_alt" if es_par else "row_norm"
            tree.insert("", "end", values=vals, tags=(tag,))

    def _ejecutar_analisis_ui(self) -> None:
        df_activo = self._obtener_dataframe_activo()
        variable = self._var_variable_calc.get()
        agrupacion = self._var_agrupacion_calc.get()
        por_caja = self._var_unidad_calc.get().startswith("Por Caja")
        alfa = float(self._var_alfa_calc.get())
        excluir_ng = self._var_excluir_nogerm.get()
        excluir_at = self._var_excluir_atipicos.get()

        df_sem, df_an = preparar_datos_analisis(
            df_activo, variable, agrupacion,
            excluir_no_germinadas=excluir_ng,
            por_caja=por_caja,
            excluir_atipicos=excluir_at,
        )
        res = ejecutar_suite_completa(df_sem, df_an, variable, alfa)
        self._ultimo_resultado_stats = {
            "df_activo": df_activo,
            "df_sem": df_sem,
            "df_an": df_an,
            "variable": variable,
            "agrupacion": agrupacion,
            "por_caja": por_caja,
            "alfa": alfa,
            "res": res,
        }

        unidad_txt = "cajas/lotes" if (por_caja or variable == "Germinó") else "semillas"
        resumen_txt = generar_texto_conclusion(variable, agrupacion, unidad_txt, res, alfa)
        self._txt_conclusion_anova.configure(state="normal")
        self._txt_conclusion_anova.delete("1.0", tk.END)
        self._txt_conclusion_anova.insert("1.0", resumen_txt)
        self._txt_conclusion_anova.configure(state="disabled")

        self._poblar_treeview_df(
            self._tree_ranking, res["ranking"],
            ["Lugar", "Grupo", "n", "Promedio ± DE", "Varianza", "CV %", "% Germ.", "Tukey", "LSD"]
        )
        self._poblar_treeview_df(
            self._tree_anova, res["anova"]["tabla"] if res["anova"] else pd.DataFrame(),
            ["Fuente", "GL", "SC", "CM", "F calc", "F crít", "Comparación", "p-valor", "Conclusión"]
        )
        self._poblar_treeview_df(
            self._tree_tukey, res["tukey"],
            ["Grupo 1", "Grupo 2", "Conclusión", "|Diferencia|", "DHS", "Comparación", "q calc", "q crít", "p ajustada"]
        )
        self._poblar_treeview_df(
            self._tree_lsd, res["lsd"],
            ["Grupo 1", "Grupo 2", "Conclusión", "|Diferencia|", "LSD", "Comparación", "t calc", "t crít", "p-valor"]
        )
        self._poblar_treeview_df(
            self._tree_cuartiles, res["cuartiles"],
            ["Grupo", "n", "Mínimo", "Q1", "Mediana", "Q3", "Máximo", "IQR", "Lím. Inf", "Lím. Sup", "Atípicos", "Bowley", "Forma"]
        )

        self._actualizar_kpis_tablero()
        self._dibujar_grafica_tablero()

    # -----------------------------------------------------------------------
    # Página 3: Tablero y Gráficas Interactivas (Matplotlib + Letras Tukey/LSD)
    # -----------------------------------------------------------------------

    def _construir_seccion_tablero(self, padre: ctk.CTkFrame) -> None:
        padre.grid_columnconfigure((0, 1, 2, 3), weight=1)
        padre.grid_rowconfigure(2, weight=1)

        # Marco Principal de Controles del Gráfico
        ctrl_graf = ctk.CTkFrame(
            padre, fg_color=UI_COLORS["bg_card"],
            corner_radius=10, border_width=1, border_color=UI_COLORS["border"]
        )
        ctrl_graf.grid(row=0, column=0, columnspan=4, pady=(0, 8), sticky="ew")

        # --- FILA 1: Tipo de Gráfica, Letras, Barras de Error y Botón Guardar ---
        fila1 = ctk.CTkFrame(ctrl_graf, fg_color="transparent")
        fila1.pack(fill="x", padx=10, pady=(10, 4))

        ctk.CTkLabel(
            fila1, text="📊 Tipo de Gráfica:",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=UI_COLORS["text_primary"],
        ).pack(side="left", padx=(6, 6))

        self._var_tipo_grafica = tk.StringVar(value="1. Barras: Promedio + Letras + Patrón")
        ctk.CTkComboBox(
            fila1,
            variable=self._var_tipo_grafica,
            values=[
                "1. Barras: Promedio + Letras + Patrón",
                "2. Cajas (Boxplot): Distribución + Letras",
                "3. Tukey Cluster: Intervalos Simultáneos (DHS)",
                "4. LSD de Fisher: ¿Qué pares superan la LSD?",
                "5. Dispersión: Datos alrededor de su Promedio (±1 DE)",
                "6. Comparativa: Raíz vs Tallo por Grupo",
                "7. Germinación: Germinaron vs No Germinaron (%)",
                "8. Correlación Alométrica: Raíz vs Tallo (OLS)",
            ],
            width=290,
            state="readonly",
            command=lambda _: self._dibujar_grafica_tablero(),
        ).pack(side="left", padx=6)

        ctk.CTkLabel(
            fila1, text="Letras:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(10, 4))

        self._var_letras_graf = tk.StringVar(value="Tukey")
        ctk.CTkSegmentedButton(
            fila1, values=["Tukey", "LSD"], variable=self._var_letras_graf,
            width=110, command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        ctk.CTkLabel(
            fila1, text="Barras de error:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(10, 4))

        self._var_error_graf = tk.StringVar(value="Desviación estándar")
        ctk.CTkComboBox(
            fila1, variable=self._var_error_graf,
            values=["Desviación estándar", "Error estándar (SEM)", "IC 95%"],
            width=150, state="readonly",
            command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            fila1, text="💾 Guardar Gráfica (PNG/PDF)",
            width=170, fg_color=UI_COLORS["accent_green"], hover_color=UI_COLORS["accent_hover"],
            text_color="#ffffff", font=ctk.CTkFont(size=11, weight="bold"),
            command=self._guardar_grafica_archivo,
        ).pack(side="right", padx=6)

        # --- FILA 2: Filtros Rápidos Locales (Variedad, Tratamiento, Desde y Hasta) ---
        fila2 = ctk.CTkFrame(ctrl_graf, fg_color="transparent")
        fila2.pack(fill="x", padx=10, pady=(4, 10))

        ctk.CTkLabel(
            fila2, text="🔍 Variedad:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(6, 2))

        lista_vars = ["Todas las variedades"]
        try:
            df_temp = self._obtener_df_activo() if hasattr(self, "_obtener_df_activo") else pd.DataFrame()
            if not df_temp.empty:
                col_v = "Tipo Semilla" if "Tipo Semilla" in df_temp.columns else ("Variedad de Semilla" if "Variedad de Semilla" in df_temp.columns else None)
                if col_v:
                    lista_vars += sorted(df_temp[col_v].dropna().unique().tolist())
        except Exception:
            pass
        if len(lista_vars) == 1:
            lista_vars += ["Semilla de maíz blanco", "Semilla de maíz rosa", "Semilla de maíz cacahuazintle", "Semilla de lenteja"]

        self._var_filtro_var_tablero = tk.StringVar(value="Todas las variedades")
        ctk.CTkComboBox(
            fila2, variable=self._var_filtro_var_tablero,
            values=lista_vars, width=170, state="readonly",
            command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        ctk.CTkLabel(
            fila2, text="Tratamiento:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(8, 2))

        self._var_filtro_trat_tablero = tk.StringVar(value="Todos los tratamientos")
        ctk.CTkComboBox(
            fila2, variable=self._var_filtro_trat_tablero,
            values=["Todos los tratamientos", "Control", "30s", "45s", "60s"], width=130, state="readonly",
            command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        # Fechas únicas disponibles
        lista_fechas = []
        try:
            if not df_temp.empty and "fecha_registro" in df_temp.columns:
                fechas_unicas = pd.to_datetime(df_temp["fecha_registro"], errors="coerce").dt.strftime("%Y-%m-%d").dropna().unique().tolist()
                lista_fechas = sorted(list(set(fechas_unicas)))
        except Exception:
            pass

        ctk.CTkLabel(
            fila2, text="📅 Desde:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(8, 2))

        self._var_fecha_desde = tk.StringVar(value="Todas")
        ctk.CTkComboBox(
            fila2, variable=self._var_fecha_desde,
            values=["Todas"] + lista_fechas, width=115, state="readonly",
            command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        ctk.CTkLabel(
            fila2, text="Hasta:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(side="left", padx=(6, 2))

        self._var_fecha_hasta = tk.StringVar(value="Todas")
        ctk.CTkComboBox(
            fila2, variable=self._var_fecha_hasta,
            values=["Todas"] + lista_fechas, width=115, state="readonly",
            command=lambda _: self._dibujar_grafica_tablero()
        ).pack(side="left", padx=4)

        # 4 Tarjetas KPI
        self._lbl_kpi_total = self._crear_kpi_card(padre, 1, 0, "🌱 Semillas Activas", "0", UI_COLORS["accent_indigo"])
        self._lbl_kpi_germ  = self._crear_kpi_card(padre, 1, 1, "✓ Germinación Global", "0 (0%)", UI_COLORS["accent_green"])
        self._lbl_kpi_lotes = self._crear_kpi_card(padre, 1, 2, "📦 Cajas / Lotes", "0", UI_COLORS["accent_amber"])
        self._lbl_kpi_fval  = self._crear_kpi_card(padre, 1, 3, "🧮 ANOVA (F / p-valor)", "—", UI_COLORS["accent_red"])

        # Contenedor del Lienzo Matplotlib
        self._chart_container = ctk.CTkFrame(
            padre, fg_color=UI_COLORS["bg_card"],
            corner_radius=12, border_width=1, border_color=UI_COLORS["border"]
        )
        self._chart_container.grid(row=2, column=0, columnspan=4, pady=(8, 0), sticky="nsew")
        self._chart_container.grid_columnconfigure(0, weight=1)
        self._chart_container.grid_rowconfigure(0, weight=1)

        self._canvas_widget = None

    def _crear_kpi_card(self, padre, row: int, col: int, titulo: str, valor: str, color_borde) -> ctk.CTkLabel:
        padx = (0, 6) if col == 0 else ((6, 6) if col in (1, 2) else (6, 0))
        card = ctk.CTkFrame(
            padre, fg_color=UI_COLORS["bg_card"],
            corner_radius=10, border_width=2, border_color=color_borde
        )
        card.grid(row=row, column=col, padx=padx, pady=4, sticky="nsew")

        ctk.CTkLabel(
            card, text=titulo,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=UI_COLORS["text_secondary"],
        ).pack(anchor="w", padx=14, pady=(10, 2))

        lbl_val = ctk.CTkLabel(
            card, text=valor,
            font=ctk.CTkFont(family="Segoe UI", size=19, weight="bold"),
            text_color=UI_COLORS["text_primary"],
        )
        lbl_val.pack(anchor="w", padx=14, pady=(0, 10))
        return lbl_val

    def _actualizar_kpis_tablero(self) -> None:
        if not self._ultimo_resultado_stats or self._ultimo_resultado_stats["df_activo"].empty:
            stats_db = self.db.contar_registros()
            total = stats_db.get("total", 0)
            germ = stats_db.get("germinadas", 0)
            lotes = stats_db.get("lotes", 0)
            pct = round((germ / total * 100), 1) if total > 0 else 0.0
            self._lbl_kpi_total.configure(text=str(total))
            self._lbl_kpi_germ.configure(text=f"{germ} ({pct}%)")
            self._lbl_kpi_lotes.configure(text=str(lotes))
            self._lbl_kpi_fval.configure(text="Sin datos")
            return

        df_sem = self._ultimo_resultado_stats["df_sem"]
        anova = self._ultimo_resultado_stats["res"].get("anova")
        total = len(df_sem)
        germ = int((df_sem["Germinó"] == 1).sum()) if "Germinó" in df_sem.columns else 0
        cajas = df_sem["Corrida"].nunique() if "Corrida" in df_sem.columns else 0
        pct = (germ / total * 100) if total > 0 else 0.0

        self._lbl_kpi_total.configure(text=f"{total:,}")
        self._lbl_kpi_germ.configure(text=f"{germ:,} ({pct:.1f}%)")
        self._lbl_kpi_lotes.configure(text=f"{cajas:,}")
        if anova:
            estado = "✔ Sig." if anova["significativo"] else "✘ No sig."
            self._lbl_kpi_fval.configure(text=f"F={anova['F calculada']:.2f} (p={anova['p-valor']:.3g}) {estado}")
        else:
            self._lbl_kpi_fval.configure(text="Insuf. grupos")

    def _dibujar_grafica_tablero(self) -> None:
        if not hasattr(self, "_chart_container"):
            return

        # 1. Obtener DataFrame base activo
        df_base = self._obtener_df_activo() if hasattr(self, "_obtener_df_activo") else pd.DataFrame()
        if df_base.empty and self._ultimo_resultado_stats:
            df_base = self._ultimo_resultado_stats.get("df_an", pd.DataFrame())

        # 2. Filtrar localmente por Variedad
        if hasattr(self, "_var_filtro_var_tablero"):
            v_sel = self._var_filtro_var_tablero.get()
            if v_sel and v_sel != "Todas las variedades":
                col_v = "Tipo Semilla" if "Tipo Semilla" in df_base.columns else "Variedad de Semilla"
                if col_v in df_base.columns:
                    df_base = df_base[df_base[col_v] == v_sel]

        # 3. Filtrar localmente por Tratamiento
        if hasattr(self, "_var_filtro_trat_tablero"):
            t_sel = self._var_filtro_trat_tablero.get()
            if t_sel and t_sel != "Todos los tratamientos":
                if "Tratamiento" in df_base.columns:
                    df_base = df_base[df_base["Tratamiento"] == t_sel]

        # 4. Filtrar localmente por Rango de Fechas (Desde / Hasta)
        if hasattr(self, "_var_fecha_desde") and hasattr(self, "_var_fecha_hasta") and "fecha_registro" in df_base.columns:
            f_desde = self._var_fecha_desde.get()
            f_hasta = self._var_fecha_hasta.get()
            
            fechas_dt = pd.to_datetime(df_base["fecha_registro"], errors="coerce")
            if f_desde and f_desde != "Todas":
                df_base = df_base[fechas_dt >= pd.to_datetime(f_desde)]
            if f_hasta and f_hasta != "Todas":
                df_base = df_base[fechas_dt <= pd.to_datetime(f_hasta)]

        # 5. Análisis estadístico flexible (permite 1 o más grupos)
        variable = self._var_variable.get() if hasattr(self, "_var_variable") else "Raíz"
        agrupacion = self._var_agrupacion.get() if hasattr(self, "_var_agrupacion") else "Tratamiento"
        
        try:
            from analytics_engine import preparar_datos_analisis, ejecutar_suite_completa
            df_sem_f, df_an_f = preparar_datos_analisis(
                df_base, variable, agrupacion, 
                replicas_modo="caja", excluir_atipicos=False
            )
            res_f = ejecutar_suite_completa(df_sem_f, df_an_f, variable, 0.05)
            
            # Actualizar KPIs en tiempo real
            total_sem = len(df_base)
            germ_glob = int((df_base["Estado"] == "Germinó").sum()) if "Estado" in df_base.columns else 0
            pct_germ = (germ_glob / total_sem * 100) if total_sem > 0 else 0.0
            
            if hasattr(self, "_lbl_kpi_total"):
                self._lbl_kpi_total.configure(text=str(total_sem))
            if hasattr(self, "_lbl_kpi_germ"):
                self._lbl_kpi_germ.configure(text=f"{germ_glob} ({pct_germ:.1f}%)")
            if hasattr(self, "_lbl_kpi_lotes"):
                cajas_n = df_base["Lote Origen"].nunique() if "Lote Origen" in df_base.columns else 1
                self._lbl_kpi_lotes.configure(text=str(cajas_n))
            if hasattr(self, "_lbl_kpi_fval") and res_f.get("anova"):
                an_res = res_f["anova"]
                self._lbl_kpi_fval.configure(text=f"F={an_res['F calculada']:.2f} (p={an_res['p-valor']:.3f})")
            elif hasattr(self, "_lbl_kpi_fval"):
                self._lbl_kpi_fval.configure(text="1 grupo (N/A)")

        except Exception:
            datos_st = self._ultimo_resultado_stats
            if not datos_st:
                return
            df_an_f = datos_st["df_an"]
            res_f = datos_st["res"]

        for child in self._chart_container.winfo_children():
            child.destroy()

        p = TTK_PALETTES[self._modo_actual]
        paleta = ["#2a78d6", "#10b981", "#f59e0b", "#6366f1", "#ec4899", "#14b8a6", "#f97316", "#8b5cf6"]

        # 4. Alta resolución (DPI 130) y espacio inferior amplio para nombres largos
        fig = Figure(figsize=(10.5, 5.2), dpi=130, facecolor=p["bg_card"])
        fig.subplots_adjust(bottom=0.32, left=0.10, right=0.95, top=0.90)
        
        ax = fig.add_subplot(111)
        ax.set_facecolor(p["bg_card"])
        for spine in ax.spines.values():
            spine.set_color(p["border"])
        ax.tick_params(colors=p["text_primary"], labelsize=9)
        ax.xaxis.label.set_color(p["text_primary"])
        ax.yaxis.label.set_color(p["text_primary"])
        ax.title.set_color(p["text_primary"])
        ax.grid(True, linestyle="--", alpha=0.22, color=p["text_muted"])

        # Limpiar valores nulos o 'nan' en el grupo
        df_an = df_an_f.copy()
        if not df_an.empty and "Grupo" in df_an.columns:
            df_an["Grupo"] = df_an["Grupo"].fillna("Sin especificar").astype(str)
            df_an = df_an[df_an["Grupo"] != "nan"]

        res = res_f
        desc = res["descriptivos"]
        if not desc.empty:
            desc["Grupo"] = desc["Grupo"].fillna("Sin especificar").astype(str)
            desc = desc[desc["Grupo"] != "nan"]

        anova = res["anova"]
        tipo_g = self._var_tipo_grafica.get()
        usar_tukey = self._var_letras_graf.get() == "Tukey"
        letras_df = res["letras_tukey"] if usar_tukey else res["letras_lsd"]
        
        mapa_letras = {}
        if not letras_df.empty and "Grupo" in letras_df.columns:
            mapa_letras = dict(zip(letras_df["Grupo"].astype(str), letras_df["Letras"]))

        u = "%" if variable == "Germinó" else ("hojas" if variable == "Hojas" else "mm")

        if df_an.empty or desc.empty or len(desc) < 1:
            ax.text(0.5, 0.5, "No hay datos suficientes para mostrar con los filtros actualizados.\nSelecciona otra opción en 'Comparar por' o 'Filtrar Variedad'.",
                    ha="center", va="center", color=p["text_secondary"], fontsize=11, transform=ax.transAxes)
        elif tipo_g.startswith("1."):
            # Gráfica 1: Barras con medias, errores y letras
            d = desc.sort_values("Media", ascending=False).reset_index(drop=True)
            err_mode = self._var_error_graf.get()
            if err_mode == "Error estándar (SEM)":
                err = d["SEM"].fillna(0).values
            elif err_mode == "IC 95%":
                err = (d["IC sup"] - d["Media"]).fillna(0).values
            else:
                err = d["Desviación estándar"].fillna(0).values

            colores = ["#8a8984" if ("control" in g.lower() or "patrón" in g.lower()) else paleta[i % len(paleta)]
                       for i, g in enumerate(d["Grupo"])]
            x_pos = np.arange(len(d))
            bars = ax.bar(x_pos, d["Media"], yerr=err, capsize=5, color=colores, edgecolor=p["border"], alpha=0.9)
            
            ax.set_xticks(x_pos)
            ax.set_xticklabels(d["Grupo"], rotation=30, ha="right", fontsize=8.5)
            ax.set_ylabel(f"{variable} promedio ({u})")

            tope = float((d["Media"] + err).max()) if len(d) else 1.0
            for idx, (bar, (_, row)) in enumerate(zip(bars, d.iterrows())):
                let = mapa_letras.get(str(row["Grupo"]), "")
                y_top = row["Media"] + err[idx]
                ax.text(bar.get_x() + bar.get_width() / 2, y_top + tope * 0.03,
                        f"{row['Media']:.1f} {let}".strip(), ha="center", va="bottom",
                        fontsize=8.5, fontweight="bold", color=p["text_primary"])

            ctrl_rows = d[d["Grupo"].str.lower().str.contains("control|patrón|patron")]
            if not ctrl_rows.empty:
                val_ctrl = float(ctrl_rows["Media"].iloc[0])
                ax.axhline(val_ctrl, color="#f59e0b", linestyle="--", linewidth=1.8, label=f"Control: {val_ctrl:.2f} {u}")
                ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"])

            ax.set_ylim(0, tope * 1.25 if tope > 0 else 1)
            sub_anova = f"F={anova['F calculada']:.2f}, p={anova['p-valor']:.4f}" if anova else "Sin ANOVA"
            ax.set_title(f"{variable} por {agrupacion} — Letras de {self._var_letras_graf.get()} ({sub_anova})", fontweight="bold", fontsize=11)

        elif tipo_g.startswith("2."):
            # Gráfica 2: Boxplot (Distribución)
            grupos = [g_vals["Valor"].values for _, g_vals in df_an.groupby("Grupo")]
            nombres = [g_name for g_name, _ in df_an.groupby("Grupo")]
            
            bp = ax.boxplot(grupos, labels=nombres, patch_artist=True, medianprops=dict(color="white", linewidth=2))
            for i, patch in enumerate(bp["boxes"]):
                patch.set(facecolor=paleta[i % len(paleta)], edgecolor=p["border"], alpha=0.85)

            ax.set_xticklabels(nombres, rotation=30, ha="right", fontsize=8.5)
            ax.set_ylabel(f"{variable} ({u})")
            ax.set_title(f"Distribución (Boxplot) de {variable} por {agrupacion}", fontweight="bold", fontsize=11)

        elif tipo_g.startswith("3."):
            # Gráfica 3: Violín
            try:
                grupos = [g_vals["Valor"].values for _, g_vals in df_an.groupby("Grupo")]
                nombres = [g_name for g_name, _ in df_an.groupby("Grupo")]
                parts = ax.violinplot(grupos, showmeans=True, showextrema=True, showmedians=True)
                for i, pc in enumerate(parts['bodies']):
                    pc.set_facecolor(paleta[i % len(paleta)])
                    pc.set_edgecolor(p["border"])
                    pc.set_alpha(0.8)

                ax.set_xticks(np.arange(1, len(nombres) + 1))
                ax.set_xticklabels(nombres, rotation=30, ha="right", fontsize=8.5)
                ax.set_ylabel(f"{variable} ({u})")
                ax.set_title(f"Diagrama de Violín — {variable} por {agrupacion}", fontweight="bold", fontsize=11)
            except Exception:
                ax.text(0.5, 0.5, "No hay suficientes datos para generar el diagrama de violín.", ha="center", va="center", color=p["text_secondary"])

        else:
            # Gráfica 4: Germinación apilada / comparativa
            if "Germinó" in res.get("columnas_disponibles", []):
                germ_tab = df_an.groupby("Grupo")["Valor"].agg(
                    Total="count",
                    Germinadas=lambda x: (x == 1).sum()
                ).reset_index()
                germ_tab["No germinadas"] = germ_tab["Total"] - germ_tab["Germinadas"]
                
                x_pos = np.arange(len(germ_tab))
                b1 = ax.bar(x_pos, germ_tab["Germinadas"], label="Germinadas", color="#10b981", edgecolor=p["border"], alpha=0.9)
                b2 = ax.bar(x_pos, germ_tab["No germinadas"], bottom=germ_tab["Germinadas"], label="No germinadas", color="#ef4444", edgecolor=p["border"], alpha=0.9)

                ax.set_xticks(x_pos)
                ax.set_xticklabels(germ_tab["Grupo"], rotation=30, ha="right", fontsize=8.5)
                ax.set_ylabel("Cantidad de semillas")
                ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"])
                ax.set_title(f"Conteo de Germinación por {agrupacion}", fontweight="bold", fontsize=11)
            else:
                ax.text(0.5, 0.5, "Gráfica no disponible para la métrica seleccionada.", ha="center", va="center", color=p["text_secondary"])

        canvas = FigureCanvasTkAgg(fig, master=self._chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)

    def _guardar_grafica_archivo(self) -> None:
        if self._fig_actual is None:
            messagebox.showwarning("Sin gráfica", "Primero genera una gráfica en el tablero.", parent=self)
            return
        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar gráfica actual",
            defaultextension=".png",
            initialfile=f"Grafica_BioGrow_{datetime.now().strftime('%Y%m%d_%H%M')}.png",
            filetypes=[("Imagen PNG (Alta resolución)", "*.png"), ("Documento PDF", "*.pdf"), ("Imagen JPG", "*.jpg")],
        )
        if ruta:
            self._fig_actual.savefig(ruta, dpi=300, bbox_inches="tight")
            messagebox.showinfo("✅ Gráfica Guardada", f"Gráfica exportada en alta resolución (300 DPI):\n{ruta}", parent=self)
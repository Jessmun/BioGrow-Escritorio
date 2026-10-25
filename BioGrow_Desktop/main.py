"""
main.py — Punto de entrada de BioGrow Desktop (Versión SQLite Local)
Ejecutar con: python main.py
Compilar con: pyinstaller BioGrow_Analytics.spec
"""
import sys
import os

# Asegurar que el directorio del módulo esté en el path de Python
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

from database import DatabaseManager
from interfaz import BioGrowApp


def _mostrar_error_inicio(titulo: str, mensaje: str) -> None:
    """Muestra un mensaje de error gráfico antes de que exista la ventana principal."""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(titulo, mensaje, parent=root)
    root.destroy()


def main() -> None:
    """Punto de entrada principal de BioGrow Analytics (SQLite)."""
    # 1. Inicializar la base de datos local
    try:
        db = DatabaseManager()
        db.init_db()
    except Exception as e:
        _mostrar_error_inicio(
            "Error de base de datos — BioGrow Analytics",
            "BioGrow Analytics no pudo inicializar la base de datos local.\n\n"
            f"Detalle:\n{e}",
        )
        return

    # 2. Arrancar la aplicación de escritorio
    try:
        app = BioGrowApp(db)
    except Exception as e:
        import tkinter as tk
        ventana = getattr(tk, "_default_root", None)
        if ventana is not None:
            try:
                ventana.destroy()
            except tk.TclError:
                pass

        _mostrar_error_inicio(
            "Error al iniciar — BioGrow Analytics",
            "BioGrow Analytics no pudo iniciar la interfaz.\n\n"
            f"Detalle:\n{e}",
        )
        return

    app.mainloop()


if __name__ == "__main__":
    main()
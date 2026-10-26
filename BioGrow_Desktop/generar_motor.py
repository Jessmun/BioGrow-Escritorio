"""
generar_motor.py — Regenera motor_biogrow.py a partir de dashboardst.py.

Copia, sin cambios, todas las funciones de cálculo del dashboard que no usan Streamlit ni Plotly,
para que la app de escritorio haga exactamente los mismos análisis.

Uso (desde la carpeta BioGrow_Desktop):
    python generar_motor.py ruta/a/dashboardst.py motor_biogrow.py
"""
import ast, sys, re


def main():
    origen, destino = sys.argv[1], sys.argv[2]
    src = open(origen, encoding="utf-8").read()
    arbol = ast.parse(src)
    EXCLUIR = {"_cargar", "_ejecutar", "main", "html_descargas", "_tabla_html", "_celda_txt", "_esc",
               "_nombre_figura", "_sin_encimados", "_partir_linea", "aplicar_tipografia", "_estilo",
               "_linea_patron", "_fig_para_exportar"}
    INCLUIR_FORZADO = {"cajas_vacias"}
    UI = ("st.", "go.", "px.", "make_subplots", "components.")
    partes, nombres = [], []
    for n in arbol.body:
        seg = ast.get_source_segment(src, n)
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            if any(m in seg for m in ("streamlit", "plotly")):
                continue
            partes.append(seg)
        elif isinstance(n, ast.FunctionDef):
            if n.name in EXCLUIR:
                continue
            cuerpo = re.sub(r'""".*?"""', "", seg, flags=re.S)
            if n.name not in INCLUIR_FORZADO and any(u in cuerpo for u in UI):
                continue
            partes.append(seg); nombres.append(n.name)
        elif isinstance(n, (ast.Assign, ast.AnnAssign)):
            if any(u in seg for u in UI):
                continue
            partes.append(seg)
        elif isinstance(n, ast.Try):
            if not any(m in seg for m in ("streamlit", "plotly")):
                partes.append(seg)
        elif isinstance(n, ast.If):  # if __name__ == "__main__"
            continue
        elif isinstance(n, ast.Expr) and isinstance(getattr(n, "value", None), ast.Constant):
            continue
    cabecera = '''"""
    motor_biogrow.py — Motor de lectura y análisis estadístico de BioGrow.

    Copia fiel de las funciones de cálculo de dashboardst.py (dashboard de Streamlit), sin
    Streamlit ni Plotly, para que la app de escritorio haga exactamente los mismos análisis:
    lectura de Excel (todos los formatos), descartes, promedio por caja, descriptivos, ANOVA,
    Tukey, LSD, letras y cuartiles. Generado automáticamente: si cambias un cálculo en
    dashboardst.py, vuelve a generar este archivo en lugar de editarlo a mano.
    Internamente las longitudes se leen en cm (igual que el dashboard).
    """
    '''
    cabecera = "\n".join(l[4:] if l.startswith("    ") else l for l in cabecera.split("\n"))
    open(destino, "w", encoding="utf-8").write(cabecera + "\n\n".join(partes) + "\n")
    print(len(nombres), "funciones copiadas")


if __name__ == "__main__":
    main()

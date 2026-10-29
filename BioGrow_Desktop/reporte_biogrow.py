"""
reporte_biogrow.py — Reporte PDF de 'Cálculos y Modelos'.

Incluye la configuración del análisis, el dictamen del ANOVA, el ranking con letras, la gráfica
de promedios con letras, promedio/DE/varianza, la tabla ANOVA y las pruebas de Tukey y LSD
(valores críticos, GL, comparaciones y letras). Usa el mismo estilo que el PDF del historial.
"""
from __future__ import annotations

import io
import os
from datetime import datetime

import pandas as pd

import analisis_biogrow as AN
import graficas_biogrow as GR

VERDE = "#059669"
GRIS_TEXTO = "#475569"
OSCURO = "#0f172a"


def _fuentes():
    """Registra DejaVu Sans (viene con matplotlib) para que se vean ±, α, ², ✔, ✘ y −."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import matplotlib
    carpeta = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
    try:
        if "DejaVu" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(carpeta, "DejaVuSans.ttf")))
            pdfmetrics.registerFont(TTFont("DejaVu-Bold", os.path.join(carpeta, "DejaVuSans-Bold.ttf")))
        return "DejaVu", "DejaVu-Bold"
    except Exception:
        return "Helvetica", "Helvetica-Bold"


def _esc(texto) -> str:
    return (str(texto).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _tabla(df, estilos, fuente, negrita, ancho_total, columnas=None, dec=3, tam=7.5):
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, Table, TableStyle
    from reportlab.lib.styles import ParagraphStyle

    if df is None or df.empty:
        return Paragraph("Sin datos para esta tabla.", estilos["normal"])
    if columnas:
        df = df[[c for c in columnas if c in df.columns]]
    cols, filas = AN.formatear_tabla(df, dec)
    st_celda = ParagraphStyle("celda", fontName=fuente, fontSize=tam, leading=tam + 2)
    st_enc = ParagraphStyle("enc", fontName=negrita, fontSize=tam, leading=tam + 2, textColor=colors.white)
    # anchos proporcionales al contenido
    largos = [max([len(c)] + [len(f[i]) for f in filas[:200]]) for i, c in enumerate(cols)]
    largos = [min(max(l, 8), 42) for l in largos]
    total = sum(largos)
    anchos = [ancho_total * l / total for l in largos]
    datos = [[Paragraph(_esc(c), st_enc) for c in cols]]
    for f in filas:
        datos.append([Paragraph(_esc(v), st_celda) for v in f])
    t = Table(datos, colWidths=anchos, repeatRows=1)
    estilo = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(VERDE)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])
    for i, f in enumerate(filas, start=1):
        estilo.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f8fafc") if i % 2 == 0 else colors.white)
        texto = " ".join(f)
        if "✔" in texto:
            estilo.add("TEXTCOLOR", (0, i), (-1, i), colors.HexColor(VERDE))
        elif "✘" in texto:
            estilo.add("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#b91c1c"))
    t.setStyle(estilo)
    return t


def generar_pdf(ruta: str, res: dict, config: dict, paleta_clara: dict) -> None:
    """config: textos de la configuración (fuente, archivos, variable, comparar, variedad, unidad, alfa, descartes)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer

    fuente, negrita = _fuentes()
    est = {
        "titulo": ParagraphStyle("t", fontName=negrita, fontSize=16, textColor=colors.HexColor(OSCURO), spaceAfter=4),
        "sub": ParagraphStyle("s", fontName=fuente, fontSize=9.5, textColor=colors.HexColor(GRIS_TEXTO), spaceAfter=10),
        "h2": ParagraphStyle("h2", fontName=negrita, fontSize=12.5, textColor=colors.HexColor(VERDE),
                             spaceBefore=10, spaceAfter=5),
        "normal": ParagraphStyle("n", fontName=fuente, fontSize=9, leading=12.5, textColor=colors.HexColor(OSCURO)),
        "nota": ParagraphStyle("no", fontName=fuente, fontSize=8, leading=11, textColor=colors.HexColor(GRIS_TEXTO)),
        "dictamen": ParagraphStyle("d", fontName=negrita, fontSize=10.5, leading=14, spaceAfter=3),
    }
    doc = SimpleDocTemplate(ruta, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30,
                            bottomMargin=34, title="Reporte de análisis BioGrow", author="BioGrow Analytics")
    ancho = landscape(letter)[0] - 60
    an, u, alfa = res.get("anova"), res["u"], res["alfa"]
    m = res["metricas"]
    germ = m["% germinación global"]
    el = []

    el.append(Paragraph("BioGrow Analytics — Reporte de Análisis Estadístico", est["titulo"]))
    el.append(Paragraph(_esc(
        f"{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {res['nombre_var']} ({u})  |  Comparación: {res['agrupacion']}  |  "
        f"Cada dato: una {'caja' if res['por_caja'] else 'semilla'}  |  α = {alfa}  |  "
        f"Datos: {config.get('fuente', '')}"
        + (f" ({config['archivos']})" if config.get("archivos") else "")
        + (f"  |  Variedad: {config['variedad']}" if config.get("variedad", "").lower() != "todas las variedades" else "")),
        est["sub"]))

    if an:
        color = VERDE if an["significativo"] else "#b91c1c"
        dictamen = (("✔ Sí hay" if an["significativo"] else "✘ No hay") +
                    f" diferencias significativas en {res['nombre_var'].lower()} entre los grupos "
                    f"(F calculada = {an['F calculada']:.3f} {'>' if an['significativo'] else '≤'} F crítica = "
                    f"{an['F crítica']:.3f}; GL = {an['GL entre']} y {an['GL error']}; p = {an['p-valor']:.4f}; α = {alfa}).")
        el.append(Paragraph(f"<font color='{color}'>{_esc(dictamen)}</font>", est["dictamen"]))
    el.append(Paragraph(_esc(
        f"{m['Semillas analizadas']:,} semillas · {m['Cajas']} cajas · {m['Grupos comparados']} grupos · "
        f"{list(m.values())[3]} datos ({res['unidad_txt']}) · germinación global "
        f"{'—' if pd.isna(germ) else f'{germ:.1f} %'} · {m['Registros descartados']} registros descartados"),
        est["normal"]))

    proyecto = config.get("proyecto")
    if proyecto is not None and not proyecto.empty:
        el.append(Paragraph("Información del proyecto", est["h2"]))
        el.append(_tabla(proyecto, est, fuente, negrita, ancho * 0.8, tam=8))

    el.append(Paragraph("1. Ranking con letras de Tukey y LSD", est["h2"]))
    el.append(_tabla(res["ranking"], est, fuente, negrita, ancho, dec=2, tam=8))
    try:
        fig = GR.barras(res, paleta_clara, "Tukey", "Desviación estándar")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=170, facecolor="white")
        buf.seek(0)
        w = ancho * 0.72
        el.append(Spacer(1, 6))
        el.append(Image(buf, width=w, height=w * fig.get_figheight() / fig.get_figwidth()))
    except Exception:
        pass

    el.append(KeepTogether([Paragraph("2. Promedio, desviación estándar y varianza", est["h2"]),
                            _tabla(res["tabla_descriptivos"], est, fuente, negrita, ancho, tam=7.5)]))
    if an:
        el.append(KeepTogether([Paragraph("3. ANOVA de una vía", est["h2"]),
                                _tabla(res["tabla_anova"], est, fuente, negrita, ancho, tam=8)]))

    for num, nombre, tabla, letras, clave_crit in [
        (4, "Prueba de Tukey", res["tabla_tukey"], res.get("letras_tukey"), "q crítica"),
        (5, "Prueba LSD de Fisher", res["tabla_lsd"], res.get("letras_lsd"), "t crítica"),
    ]:
        if tabla is None or tabla.empty or not an:
            continue
        el.append(PageBreak())
        el.append(Paragraph(f"{num}. {nombre} (α = {alfa})", est["h2"]))
        sig = int(tabla["Conclusión"].astype(str).str.contains("✔").sum()) if "Conclusión" in tabla else 0
        el.append(Paragraph(_esc(f"{clave_crit} = {tabla[clave_crit].iloc[0]:.3f} · GL del error = {an['GL error']} · "
                                 f"{sig} de {len(tabla)} pares difieren"), est["nota"]))
        el.append(Spacer(1, 4))
        if letras is not None and not letras.empty:
            el.append(_tabla(letras.rename(columns={"Media": f"Promedio ({u})"}), est, fuente, negrita,
                             ancho * 0.6, tam=8))
            el.append(Spacer(1, 8))
        margen = "DHS" if num == 4 else "LSD"
        est_ = "q" if num == 4 else "t"
        p_col = "p ajustada" if num == 4 else "p-valor"
        cols = ["Grupo A", "Grupo B", "Conclusión", f"|A − B| ({u})", f"{margen} ({u})", f"|A − B| vs {margen}",
                f"{est_} calculada", f"{est_} crítica", p_col]
        el.append(_tabla(tabla, est, fuente, negrita, ancho, columnas=cols, tam=7.5))

    def pie(canvas, doc_):
        canvas.saveState()
        canvas.setFont(fuente, 7.5)
        canvas.setFillColor(colors.HexColor(GRIS_TEXTO))
        canvas.drawString(30, 18, "BioGrow Analytics — Reporte de análisis estadístico")
        canvas.drawRightString(landscape(letter)[0] - 30, 18, f"Página {doc_.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=pie, onLaterPages=pie)

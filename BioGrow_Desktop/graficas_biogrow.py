"""
graficas_biogrow.py — Gráficas del tablero (matplotlib) con la misma lógica que dashboardst.py.

Cada función recibe el resultado de analisis_biogrow.analizar() y la paleta de la app
(modo claro u oscuro) y devuelve una Figure de matplotlib.
"""
from __future__ import annotations

import textwrap

import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.patches import Rectangle
from scipy import stats

import motor_biogrow as M

PALETA = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
GRIS_CONTROL = "#8a8984"
ROJO, VERDE, GRIS = "#e34948", "#1baf7a", "#8a8984"

TIPOS = [
    "1. Barras: Promedio + Letras + Patrón",
    "2. Cajas (Boxplot): Distribución + Letras",
    "3. Tukey Cluster: Intervalos Simultáneos (DHS)",
    "4. LSD de Fisher: ¿Qué pares superan la LSD?",
    "5. Dispersión: Datos alrededor de su Promedio (±1 DE)",
    "6. Comparativa: Raíz vs Tallo por Grupo",
    "7. Germinación: Germinaron vs No Germinaron (%)",
    "8. Correlación Alométrica: Raíz vs Tallo (OLS)",
    "9. Cuartiles: Q1, Mediana, Q3, IQR y Atípicos",
]


def _etiqueta(g, ancho=16):
    partes = str(g).split(" | ")
    return "\n".join("\n".join(textwrap.wrap(p, ancho)) or p for p in partes)


def _es_control(g):
    t = str(g).lower()
    return "control" in t or "patrón" in t or "patron" in t


def _colores(grupos):
    salida, k = [], 0
    for g in grupos:
        if _es_control(g):
            salida.append(GRIS_CONTROL)
        else:
            salida.append(PALETA[k % len(PALETA)])
            k += 1
    return salida


def _figura(p, alto=5.4, ancho=11.5):
    fig = Figure(figsize=(ancho, alto), dpi=100, facecolor=p["bg_card"])
    ax = fig.add_subplot(111)
    _estilo_ejes(ax, p)
    return fig, ax


def _estilo_ejes(ax, p):
    ax.set_facecolor(p["bg_card"])
    for lado, spine in ax.spines.items():
        spine.set_color(p["border"] if lado in ("left", "bottom") else p["bg_card"])
    ax.tick_params(colors=p["text_primary"], labelsize=9)
    ax.xaxis.label.set_color(p["text_primary"])
    ax.yaxis.label.set_color(p["text_primary"])
    ax.title.set_color(p["text_primary"])
    ax.grid(True, axis="y", linestyle="--", alpha=0.25, color=p["text_muted"])
    ax.set_axisbelow(True)


def _titulo(ax, titulo, subtitulo, p):
    lineas = len(textwrap.wrap(subtitulo, 150)) if subtitulo else 0
    ax.set_title(titulo, fontweight="bold", fontsize=12, loc="left", color=p["text_primary"], pad=14 + 13 * lineas)
    if subtitulo:
        subtitulo = "\n".join(textwrap.wrap(subtitulo, 150))
        ax.annotate(subtitulo, xy=(0, 1), xycoords="axes fraction", xytext=(0, 5), textcoords="offset points",
                    fontsize=9, color=p["text_secondary"], va="bottom", ha="left")


def _ticks(ax, grupos):
    rot = 30 if len(grupos) > 7 else 0
    ax.set_xticks(np.arange(len(grupos)))
    ax.set_xticklabels([_etiqueta(g) for g in grupos], rotation=rot, ha="right" if rot else "center",
                       fontsize=8.5 if len(grupos) > 7 else 9)


def _mensaje(p, texto):
    fig, ax = _figura(p)
    ax.axis("off")
    ax.text(0.5, 0.5, texto, ha="center", va="center", fontsize=11, color=p["text_secondary"],
            transform=ax.transAxes, wrap=True)
    return fig


def _sub_anova(res):
    an = res.get("anova")
    if not an:
        return "Sin ANOVA (se necesitan 2 grupos con 2 o más datos)"
    estado = "hay diferencias" if an["significativo"] else "no hay diferencias"
    return (f"ANOVA: F = {an['F calculada']:.2f} (F crítica = {an['F crítica']:.2f}; GL {an['GL entre']} y "
            f"{an['GL error']}; p = {an['p-valor']:.4f}) → {estado} · α = {res['alfa']} · "
            f"cada dato: una {'caja' if res['por_caja'] else 'semilla'}")


def _letras(res, prueba):
    let = res.get("letras_tukey" if prueba == "Tukey" else "letras_lsd")
    if let is None or let.empty:
        return {}
    return dict(zip(let["Grupo"].astype(str), let["Letras"]))


def _ajustar(fig, n_grupos):
    abajo = 0.30 if n_grupos > 7 else 0.16
    fig.subplots_adjust(left=0.08, right=0.98, top=0.86, bottom=abajo)


# ---------------------------------------------------------------------------

def barras(res, p, prueba="Tukey", error="Desviación estándar"):
    d = res["desc"].sort_values("Media", ascending=False).reset_index(drop=True)
    u = res["u"]
    if error.startswith("Error"):
        err = d["SEM"].fillna(0).to_numpy()
    elif error.startswith("IC"):
        err = (d["IC sup"] - d["Media"]).fillna(0).to_numpy()
    else:
        err = d["Desviación estándar"].fillna(0).to_numpy()
    grupos = d["Grupo"].astype(str).tolist()
    fig, ax = _figura(p)
    x = np.arange(len(d))
    ax.bar(x, d["Media"], yerr=err, capsize=5, color=_colores(grupos), edgecolor=p["border"], width=0.65,
           error_kw=dict(ecolor=p["text_primary"], elinewidth=1.1))
    mapa = _letras(res, prueba)
    tope = float(np.nanmax(d["Media"] + err)) if len(d) else 1.0
    for i, g in enumerate(grupos):
        y = d["Media"].iloc[i] + err[i]
        m = d["Media"].iloc[i]
        if m > tope * 0.12:
            ax.text(i, m * 0.5, f"{m:.2f}", ha="center", va="center", fontsize=8.5, rotation=90,
                    color="white", fontweight="bold")
        else:
            ax.text(i, y + tope * 0.015, f"{m:.2f}", ha="center", va="bottom", fontsize=8.5,
                    color=p["text_secondary"])
        if mapa.get(g):
            ax.text(i, y + tope * (0.03 if m > tope * 0.12 else 0.09), mapa[g], ha="center", va="bottom", fontsize=12,
                    fontweight="bold", color=p["text_primary"],
                    bbox=dict(boxstyle="round,pad=0.25", fc=p["bg_card"], ec=p["border"]))
    ctrl = [i for i, g in enumerate(grupos) if _es_control(g)]
    if len(ctrl) == 1:
        v = float(d["Media"].iloc[ctrl[0]])
        ax.axhline(v, color="#f59e0b", linestyle="--", linewidth=1.6, label=f"Control: {v:.2f} {u}")
        ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"], fontsize=9)
    ax.set_ylim(0, tope * 1.25 if tope > 0 else 1)
    _ticks(ax, grupos)
    ax.set_ylabel(f"{res['nombre_var']} ({u}) — promedio ± {error.split(' (')[0].lower()}")
    _titulo(ax, f"{res['nombre_var']} por {res['agrupacion'].lower()} — letras de {prueba}", _sub_anova(res), p)
    _ajustar(fig, len(grupos))
    return fig


def cajas(res, p, prueba="Tukey"):
    var, u = res["variable"], res["u"]
    datos = res["df_an"]
    orden = res["desc"].sort_values("Media", ascending=False)["Grupo"].astype(str).tolist()
    series = [pd.to_numeric(datos.loc[datos["Grupo"].astype(str) == g, var], errors="coerce").dropna().to_numpy()
              for g in orden]
    fig, ax = _figura(p)
    bp = ax.boxplot(series, positions=np.arange(len(orden)), widths=0.55, patch_artist=True,
                    showmeans=True, meanprops=dict(marker="D", markerfacecolor="white",
                                                   markeredgecolor=p["text_primary"], markersize=5),
                    medianprops=dict(color=p["text_primary"], linewidth=1.8),
                    whiskerprops=dict(color=p["text_secondary"]), capprops=dict(color=p["text_secondary"]),
                    flierprops=dict(marker="x", markeredgecolor=ROJO, markersize=6))
    for patch, c in zip(bp["boxes"], _colores(orden)):
        patch.set(facecolor=c, edgecolor=p["border"], alpha=0.8)
    mapa = _letras(res, prueba)
    maximo = max((s.max() for s in series if len(s)), default=1.0)
    for i, (g, s) in enumerate(zip(orden, series)):
        if mapa.get(g) and len(s):
            ax.text(i, s.max() + maximo * 0.04, mapa[g], ha="center", va="bottom", fontsize=12,
                    fontweight="bold", color=p["text_primary"],
                    bbox=dict(boxstyle="round,pad=0.25", fc=p["bg_card"], ec=p["border"]))
    ax.set_ylim(top=maximo * 1.2 if maximo > 0 else 1)
    _ticks(ax, orden)
    ax.set_ylabel(f"{res['nombre_var']} ({u})")
    _titulo(ax, f"Distribución de {res['nombre_var'].lower()} por grupo — letras de {prueba}",
            "Caja: Q1 a Q3 · línea: mediana · ◆: promedio · ✕: atípicos (1.5×IQR) · " + _sub_anova(res), p)
    _ajustar(fig, len(orden))
    return fig


def _pares(res, prueba):
    tabla = res.get("tukey" if prueba == "Tukey" else "lsd")
    if tabla is None or tabla.empty:
        return None
    margen = "DHS" if prueba == "Tukey" else "LSD"
    mapa = _letras(res, prueba)
    filas = []
    for _, f in tabla.iterrows():
        la, lb = mapa.get(str(f["Grupo 1"]), ""), mapa.get(str(f["Grupo 2"]), "")
        filas.append({
            "Par": f"{f['Grupo 1']}" + (f" ({la})" if la else "") + "  −  " + f"{f['Grupo 2']}" + (f" ({lb})" if lb else ""),
            "Diferencia": f["Media 1"] - f["Media 2"], "Abs": abs(f["Media 1"] - f["Media 2"]),
            "Margen": f[margen], "Sig": f["Resultado"] == M.SIG,
        })
    return pd.DataFrame(filas)


def cluster_tukey(res, p):
    d = _pares(res, "Tukey")
    if d is None:
        return _mensaje(p, "La gráfica de Tukey necesita al menos 2 grupos con 2 o más datos.")
    d = d.sort_values("Diferencia").reset_index(drop=True)
    alto = max(5.4, 1.6 + 0.3 * len(d))
    fig, ax = _figura(p, alto=alto)
    y = np.arange(len(d))
    for i, f in d.iterrows():
        c = ROJO if f["Sig"] else GRIS
        ax.plot([f["Diferencia"] - f["Margen"], f["Diferencia"] + f["Margen"]], [i, i], color=c, linewidth=2.2)
        ax.plot(f["Diferencia"], i, marker="D", color=c, markersize=7, markeredgecolor=p["text_primary"])
    ax.axvline(0, color=p["text_muted"], linestyle="--", linewidth=1.3)
    ax.set_yticks(y)
    ax.set_yticklabels(d["Par"], fontsize=8)
    ax.set_ylim(-0.7, len(d) - 0.3)
    ax.grid(True, axis="x", linestyle="--", alpha=0.25, color=p["text_muted"])
    ax.grid(False, axis="y")
    ax.set_xlabel(f"Diferencia de promedios, A − B ({res['u']})")
    _titulo(ax, "Tukey: diferencia de cada par ± DHS (intervalo simultáneo)",
            "◆ rojo: sí difieren (la línea no cruza el 0) · ◆ gris: no difieren · entre paréntesis: letra de cada grupo", p)
    fig.subplots_adjust(left=0.36, right=0.98, top=1 - 0.75 / alto, bottom=0.6 / alto)
    return fig


def pares_lsd(res, p):
    d = _pares(res, "LSD")
    if d is None:
        return _mensaje(p, "La gráfica LSD necesita al menos 2 grupos con 2 o más datos.")
    d = d.sort_values("Abs").reset_index(drop=True)
    alto = max(5.4, 1.6 + 0.3 * len(d))
    fig, ax = _figura(p, alto=alto)
    y = np.arange(len(d))
    ax.barh(y, d["Abs"], color=[VERDE if s else GRIS for s in d["Sig"]], edgecolor=p["border"], height=0.62)
    ax.scatter(d["Margen"], y, marker="|", s=260, color=p["text_primary"], linewidths=3, zorder=3,
               label="LSD (diferencia mínima significativa)")
    ax.set_yticks(y)
    ax.set_yticklabels(d["Par"], fontsize=8)
    ax.set_ylim(-0.7, len(d) - 0.3)
    ax.grid(True, axis="x", linestyle="--", alpha=0.25, color=p["text_muted"])
    ax.grid(False, axis="y")
    ax.set_xlabel(f"|A − B| ({res['u']})")
    ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"], fontsize=8.5,
              loc="lower right")
    nota = " · ANOVA no significativo: no reportes estas diferencias" if (res.get("anova") and not res["anova"]["significativo"]) else ""
    _titulo(ax, "LSD de Fisher: ¿qué pares superan la diferencia mínima?",
            "Barra verde: la diferencia supera la LSD (sí difieren) · gris: no la supera · │: valor de la LSD" + nota, p)
    fig.subplots_adjust(left=0.36, right=0.98, top=1 - 0.75 / alto, bottom=0.6 / alto)
    return fig


def dispersion(res, p):
    var, u = res["variable"], res["u"]
    datos = res["df_an"]
    d = res["desc"].sort_values("Media", ascending=False).reset_index(drop=True)
    grupos = d["Grupo"].astype(str).tolist()
    fig, ax = _figura(p)
    rng = np.random.default_rng(7)
    for i, (g, c) in enumerate(zip(grupos, _colores(grupos))):
        v = pd.to_numeric(datos.loc[datos["Grupo"].astype(str) == g, var], errors="coerce").dropna().to_numpy()
        m, s = d["Media"].iloc[i], d["Desviación estándar"].iloc[i]
        if pd.notna(s):
            ax.add_patch(Rectangle((i - 0.32, m - s), 0.64, 2 * s, color=c, alpha=0.18, linewidth=0))
        ax.scatter(i + rng.uniform(-0.22, 0.22, len(v)), v, s=16, color=c, alpha=0.75, edgecolors="none")
        ax.plot([i - 0.32, i + 0.32], [m, m], color=c, linewidth=3)
    _ticks(ax, grupos)
    ax.set_ylabel(f"{res['nombre_var']} ({u})")
    _titulo(ax, f"Datos de cada grupo alrededor de su promedio — {res['nombre_var'].lower()}",
            f"Cada punto: una {'caja' if res['por_caja'] else 'semilla'} · línea gruesa: promedio · franja: ± 1 desviación estándar", p)
    _ajustar(fig, len(grupos))
    return fig


def raiz_vs_tallo(res, p, error="Desviación estándar"):
    datos = res["ambos"]
    filas = []
    for g, sub in datos.groupby("Grupo"):
        for v in ("Raíz", "Tallo"):
            s = pd.to_numeric(sub[v], errors="coerce").dropna()
            if len(s):
                sd = s.std(ddof=1) if len(s) > 1 else 0.0
                filas.append({"Grupo": str(g), "Variable": v, "Media": s.mean(),
                              "Err": sd / np.sqrt(len(s)) if error.startswith("Error") else sd})
    r = pd.DataFrame(filas)
    if r.empty or r["Variable"].nunique() < 2:
        return _mensaje(p, "Esta gráfica necesita raíz y tallo medidos en las mismas semillas.")
    orden = r[r["Variable"] == "Raíz"].sort_values("Media", ascending=False)["Grupo"].tolist()
    orden += [g for g in r["Grupo"].unique() if g not in orden]
    fig, ax = _figura(p)
    x = np.arange(len(orden))
    for k, (v, c) in enumerate([("Raíz", PALETA[0]), ("Tallo", PALETA[2])]):
        sub = r[r["Variable"] == v].set_index("Grupo").reindex(orden)
        ax.bar(x + (k - 0.5) * 0.38, sub["Media"], width=0.38, yerr=sub["Err"].fillna(0), capsize=3,
               color=c, edgecolor=p["border"], label=v, error_kw=dict(ecolor=p["text_primary"], elinewidth=1))
        for xi, m in zip(x, sub["Media"]):
            if pd.notna(m):
                ax.text(xi + (k - 0.5) * 0.38, m, f"{m:.1f}", ha="center", va="bottom", fontsize=7.5,
                        color=p["text_secondary"])
    _ticks(ax, orden)
    ax.set_ylabel(f"Longitud promedio ({res['unidad_long']}) ± {error.split(' (')[0].lower()}")
    ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"])
    _titulo(ax, "Raíz y tallo juntos, por grupo",
            "Promedio por semilla germinada · si una barra es mucho más alta, el tratamiento favoreció esa parte", p)
    _ajustar(fig, len(orden))
    return fig


def germinacion(res, p):
    d = res["df_semilla"]
    col = M.columna_germinacion(d)
    g = d.dropna(subset=[col]).groupby("Grupo")[col].agg(["sum", "count"])
    if g.empty:
        return _mensaje(p, "No hay datos de germinación.")
    g["pct"] = g["sum"] / g["count"] * 100
    g = g.sort_values("pct", ascending=False)
    grupos = g.index.astype(str).tolist()
    fig, ax = _figura(p)
    x = np.arange(len(g))
    ax.bar(x, g["sum"], color=VERDE, edgecolor=p["border"], label="Germinaron", width=0.65)
    ax.bar(x, g["count"] - g["sum"], bottom=g["sum"], color=GRIS, edgecolor=p["border"],
           label="No germinaron", width=0.65, alpha=0.75)
    for xi, (tot, pct) in enumerate(zip(g["count"], g["pct"])):
        ax.text(xi, tot, f"{pct:.1f} %", ha="center", va="bottom", fontsize=9, fontweight="bold",
                color=p["text_primary"])
    ax.set_ylim(0, g["count"].max() * 1.15)
    _ticks(ax, grupos)
    ax.set_ylabel("Semillas")
    ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"], ncol=2,
              loc="lower right", bbox_to_anchor=(1, 1.0), fontsize=8.5)
    _titulo(ax, "Germinación por grupo: germinaron vs no germinaron",
            "Número de semillas; arriba de cada barra, el % de germinación (germinadas / sembradas)", p)
    _ajustar(fig, len(grupos))
    return fig


def correlacion(res, p):
    d = res["ambos"].dropna(subset=["Raíz", "Tallo"])
    if len(d) < 3:
        return _mensaje(p, "Esta gráfica necesita raíz y tallo medidos en las mismas semillas.")
    fig, ax = _figura(p)
    grupos = sorted(d["Grupo"].astype(str).unique())
    colores = dict(zip(grupos, _colores(grupos)))
    for g in grupos:
        s = d[d["Grupo"].astype(str) == g]
        ax.scatter(s["Raíz"], s["Tallo"], s=18, alpha=0.6, color=colores[g], label=g, edgecolors="none")
    x, y = d["Raíz"].astype(float), d["Tallo"].astype(float)
    reg = stats.linregress(x, y)
    xs = np.linspace(x.min(), x.max(), 50)
    ax.plot(xs, reg.intercept + reg.slope * xs, color=p["text_primary"], linewidth=2,
            label=f"OLS: tallo = {reg.intercept:.2f} + {reg.slope:.3f}·raíz")
    ax.grid(True, axis="x", linestyle="--", alpha=0.25, color=p["text_muted"])
    ul = res["unidad_long"]
    ax.set_xlabel(f"Raíz ({ul})")
    ax.set_ylabel(f"Tallo ({ul})")
    if len(grupos) <= 12:
        ax.legend(facecolor=p["bg_secondary"], edgecolor=p["border"], labelcolor=p["text_primary"],
                  fontsize=7.5, loc="upper left", ncol=2 if len(grupos) > 6 else 1)
    _titulo(ax, "Correlación entre longitud de raíz y de tallo, por semilla",
            f"r de Pearson = {reg.rvalue:.3f} · R² = {reg.rvalue ** 2:.3f} · p = {reg.pvalue:.4g} · n = {len(d)} semillas", p)
    fig.subplots_adjust(left=0.08, right=0.98, top=0.86, bottom=0.11)
    return fig


def cuartiles(res, p):
    """Igual que grafico_cuartiles del dashboard: caja de Q1 a Q3, mediana, promedio,
    bigotes hasta el último dato dentro de los límites y atípicos marcados."""
    tc = res.get("cuartiles")
    if tc is None or tc.empty:
        return _mensaje(p, "No hay datos para calcular cuartiles.")
    var, u, factor = res["variable"], res["u"], res.get("factor_iqr", 1.5)
    if var == "Germinó" and not res["por_caja"]:
        return _mensaje(p, "Los cuartiles no aplican a la germinación por semilla: cada semilla vale 0 % o 100 %.\n"
                           "Se pueden calcular cuando cada dato es una caja (2 o más cajas por grupo).")
    datos = res["datos_cuartiles"]
    grupos = tc["Grupo"].astype(str).tolist()
    stats_ = [{"label": g, "q1": f["Q1 (25 %)"], "med": f["Mediana (Q2)"], "q3": f["Q3 (75 %)"],
               "whislo": f["_bigote_inf"], "whishi": f["_bigote_sup"], "mean": f["Promedio"], "fliers": []}
              for g, (_, f) in zip(grupos, tc.iterrows())]
    fig, ax = _figura(p)
    bp = ax.bxp(stats_, positions=np.arange(len(grupos)), widths=0.55, patch_artist=True, showmeans=True,
                meanline=True, meanprops=dict(color=p["text_primary"], linestyle="--", linewidth=1.4),
                medianprops=dict(color=p["text_primary"], linewidth=2),
                whiskerprops=dict(color=p["text_secondary"]), capprops=dict(color=p["text_secondary"]))
    for patch, c in zip(bp["boxes"], _colores(grupos)):
        patch.set(facecolor=c, edgecolor=p["border"], alpha=0.8)
    for i, (g, (_, f)) in enumerate(zip(grupos, tc.iterrows())):
        x = pd.to_numeric(datos.loc[datos["Grupo"].astype(str) == g, var], errors="coerce").dropna()
        fuera = x[(x < f["Límite inferior"]) | (x > f["Límite superior"])]
        if len(fuera):
            ax.scatter([i] * len(fuera), fuera, marker="x", color=ROJO, s=40, zorder=3)
    _ticks(ax, grupos)
    ax.set_ylabel(f"{res['nombre_var']} ({u})")
    nivel = "semilla" if var != "Germinó" else ("caja" if res["por_caja"] else "semilla")
    _titulo(ax, f"Cuartiles de {res['nombre_var'].lower()} por grupo (ordenados por mediana)",
            f"Caja: de Q1 a Q3 (50 % central) · línea: mediana · punteada: promedio · bigotes: último dato dentro de "
            f"Q1 − {factor:g}·IQR y Q3 + {factor:g}·IQR · ✕ rojo: atípico · cada dato: una {nivel}", p)
    _ajustar(fig, len(grupos))
    fig.subplots_adjust(top=0.80)
    return fig


def descripcion_cuartiles(res) -> str:
    """Resumen corto de los cuartiles para leer junto a la gráfica."""
    tc, diag = res.get("cuartiles"), res.get("diag_cuartiles")
    if tc is None or tc.empty or (res["variable"] == "Germinó" and not res["por_caja"]):
        return ""
    u = res["u"]
    lineas = []
    if diag:
        partes = [f"{len(tc)} grupos", f"{diag['atipicos']} dato(s) atípico(s) en total"]
        if diag["iqr_max"] and pd.notna(diag["razon_iqr"]):
            partes.append(f"el grupo más disperso ({diag['iqr_max'][0]}, IQR {diag['iqr_max'][1]:.2f} {u}) tiene "
                          f"{diag['razon_iqr']:.1f} veces el IQR del menos disperso ({diag['iqr_min'][0]}, "
                          f"{diag['iqr_min'][1]:.2f} {u})")
        if diag["sesgados"]:
            partes.append("muy sesgados: " + ", ".join(diag["sesgados"]))
        lineas.append("En general: " + "; ".join(partes) + ".")
    mayor, menor = tc.iloc[0], tc.iloc[-1]
    lineas.append(f"Mediana más alta: {mayor['Grupo']} ({mayor['Mediana (Q2)']:.2f} {u}); "
                  f"más baja: {menor['Grupo']} ({menor['Mediana (Q2)']:.2f} {u}).")
    for _, f in tc.iterrows():
        at = int(f["Atípicos bajos"] + f["Atípicos altos"])
        lineas.append(f"• {f['Grupo']} (n = {f['n']}): mediana {f['Mediana (Q2)']:.2f}, Q1–Q3 = "
                      f"{f['Q1 (25 %)']:.2f}–{f['Q3 (75 %)']:.2f} {u} (IQR {f['IQR (Q3 − Q1)']:.2f}), "
                      f"{at} atípico(s), {str(f['Forma de la distribución']).lower()}.")
    return "\n".join(lineas)


def dibujar(tipo, res, p, prueba="Tukey", error="Desviación estándar"):
    if not res or not res.get("ok"):
        return _mensaje(p, (res or {}).get("mensaje") or "Sin datos para graficar.")
    n = tipo.split(".")[0]
    if n == "1":
        return barras(res, p, prueba, error)
    if n == "2":
        return cajas(res, p, prueba)
    if n == "3":
        return cluster_tukey(res, p)
    if n == "4":
        return pares_lsd(res, p)
    if n == "5":
        return dispersion(res, p)
    if n == "6":
        return raiz_vs_tallo(res, p, error)
    if n == "7":
        return germinacion(res, p)
    if n == "9":
        return cuartiles(res, p)
    return correlacion(res, p)

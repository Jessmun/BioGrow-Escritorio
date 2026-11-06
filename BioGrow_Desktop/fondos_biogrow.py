"""
fondos_biogrow.py — Ilustraciones difuminadas detrás de las tablas y gráficas de cada pestaña.

El fondo oscuro de cada ilustración se vuelve transparente (máscara por luminosidad), la imagen
se atenúa y se desvanece en los bordes, y se coloca a la derecha del área para no tapar la
primera columna. Así queda como marca de agua detrás de los datos, sin ocupar espacio propio.

- componer_tabla(): imagen del tamaño exacto de una tabla (fondo de ttk.Treeview).
- fondo_en_figura(): agrega la ilustración detrás de una gráfica de matplotlib (solo en pantalla).
"""
from __future__ import annotations

import os

from PIL import Image, ImageChops, ImageFilter

CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imagenes")

# pestaña -> (archivo, posición vertical: 0 arriba · 0.5 centro · 1 abajo)
FONDOS = {
    "registro": ("fondo_registro.png", 0.5),   # semilla germinando (registro de la germinación)
    "calculos": ("fondo_calculos.png", 1.0),   # dos plantas comparadas (pruebas estadísticas)
    "tablero": ("fondo_tablero.png", 1.0),     # maíz con líneas de flujo (gráficas y tendencias)
}

# opacidad (ilustración, resplandor difuminado) por tema; bajas para que el texto se lea bien
OPACIDAD_TABLA = {"dark": (0.32, 0.18), "light": (0.28, 0.10)}
OPACIDAD_GRAFICA = {"dark": 0.28, "light": 0.24}

_cache_original: dict = {}


def _original(clave: str):
    if clave not in _cache_original:
        archivo, _ = FONDOS.get(clave, (None, 0))
        ruta = os.path.join(CARPETA, archivo) if archivo else ""
        _cache_original[clave] = Image.open(ruta).convert("RGB") if ruta and os.path.exists(ruta) else None
    return _cache_original[clave]


def disponible(clave: str) -> bool:
    return _original(clave) is not None


def _hex(c: str):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _mascara_luz(img: Image.Image, gamma: float = 0.9) -> Image.Image:
    """Alfa según el brillo: lo oscuro del dibujo desaparece y lo luminoso queda."""
    r, g, b = img.split()
    luz = ImageChops.lighter(ImageChops.lighter(r, g), b)
    return luz.point(lambda v: 0 if v < 28 else min(255, int(((v - 28) / 227) ** gamma * 255)))


def _desvanecer(ancho: int, alto: int, borde_v: float = 0.12, borde_h: float = 0.12) -> Image.Image:
    """Máscara que se desvanece hacia los cuatro bordes."""
    def rampa(n, frac):
        k = max(1, int(n * frac))
        return [min(1.0, i / k, (n - 1 - i) / k) for i in range(n)]
    col = Image.new("L", (1, alto))
    col.putdata([int(255 * x) for x in rampa(alto, borde_v)])
    fila = Image.new("L", (ancho, 1))
    fila.putdata([int(255 * x) for x in rampa(ancho, borde_h)])
    return ImageChops.multiply(col.resize((ancho, alto)), fila.resize((ancho, alto)))


def _escalar(m: Image.Image, factor: float) -> Image.Image:
    return m.point(lambda v: int(v * factor))


def componer_tabla(clave: str, ancho: int, alto: int, color_fondo: str, tema: str):
    """Imagen RGB (ancho x alto) del color de la tabla con la ilustración difuminada a la derecha."""
    img = _original(clave)
    if img is None or ancho < 40 or alto < 40:
        return None
    _, alin = FONDOS[clave]
    op_img, op_glow = OPACIDAD_TABLA.get(tema, OPACIDAD_TABLA["dark"])
    lienzo = Image.new("RGB", (ancho, alto), _hex(color_fondo))

    # ilustración: alta como la tabla, pero nunca más de la mitad del ancho
    esc = min(alto * 0.94 / img.height, ancho * 0.5 / img.width)
    w, h = max(1, int(img.width * esc)), max(1, int(img.height * esc))
    x = ancho - w - int(ancho * 0.03)
    y = int((alto - h) * (0.5 if alin == 0.5 else 0.96))

    # resplandor: la misma ilustración más grande y muy difuminada, para que no se vea un recorte
    gw, gh = int(w * 1.5), int(h * 1.15)
    glow = img.resize((max(1, gw), max(1, gh)), Image.BILINEAR).filter(ImageFilter.GaussianBlur(max(6, gw // 10)))
    m = ImageChops.multiply(_escalar(_mascara_luz(glow), op_glow), _desvanecer(gw, gh, 0.25, 0.25))
    lienzo.paste(glow, (x - (gw - w) // 2, y - (gh - h) // 2), m)

    nitida = img.resize((w, h), Image.LANCZOS)
    m = ImageChops.multiply(_escalar(_mascara_luz(nitida), op_img), _desvanecer(w, h, 0.14, 0.14))
    lienzo.paste(nitida, (x, y), m)
    return lienzo


def ilustracion_rgba(clave: str, tema: str, alto_px: int = 600):
    """Ilustración con transparencia (RGBA) para colocarla detrás de una gráfica."""
    img = _original(clave)
    if img is None:
        return None
    esc = min(1.0, alto_px / img.height) if img.height > alto_px else 1.0
    if esc != 1.0:
        img = img.resize((max(1, int(img.width * esc)), max(1, int(img.height * esc))), Image.LANCZOS)
    op = OPACIDAD_GRAFICA.get(tema, OPACIDAD_GRAFICA["dark"])
    alfa = ImageChops.multiply(_escalar(_mascara_luz(img), op), _desvanecer(img.width, img.height, 0.14, 0.14))
    salida = img.copy()
    salida.putalpha(alfa)
    return salida


def fondo_en_figura(fig, clave: str, tema: str) -> None:
    """Pone la ilustración abajo a la derecha, detrás de las áreas de gráfica (que quedan transparentes).
    Solo para la vista en pantalla: los archivos exportados se generan aparte, sin ilustración."""
    try:
        import numpy as np
        img = ilustracion_rgba(clave, tema)
        if img is None:
            return
        for ax in fig.axes:
            ax.set_facecolor("none")
        alto_fig = fig.get_figheight()
        frac = min(0.92, 6.2 / alto_fig) if alto_fig > 0 else 0.92  # en gráficas muy altas no se estira
        fondo = fig.add_axes([0.0, 0.0, 0.985, frac], zorder=-10)
        fondo.imshow(np.asarray(img), interpolation="bilinear")
        fondo.set_aspect("equal")
        fondo.set_anchor("SE")
        fondo.axis("off")
        fondo.set_navigate(False)
    except Exception:
        pass  # decorativo: nunca debe impedir que se dibuje la gráfica

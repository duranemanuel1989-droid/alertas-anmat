#!/usr/bin/env python3
"""
Genera una imagen (PNG) tipo grilla con los registros del Boletin,
para mandarla como foto por Telegram (mas facil de leer que texto).
"""

from PIL import Image, ImageDraw, ImageFont

# --- Fuentes (DejaVu viene en ubuntu / GitHub Actions) ---
_FONT_DIR = "/usr/share/fonts/truetype/dejavu"


def _font(nombre, size):
    try:
        return ImageFont.truetype(f"{_FONT_DIR}/{nombre}", size)
    except Exception:
        return ImageFont.load_default()


F_TITULO = _font("DejaVuSans-Bold.ttf", 30)
F_HEADER = _font("DejaVuSans-Bold.ttf", 21)
F_CELDA = _font("DejaVuSans.ttf", 20)

# --- Colores ---
AZUL = (46, 109, 164)
AZUL_OSC = (33, 79, 122)
BLANCO = (255, 255, 255)
FILA_A = (255, 255, 255)
FILA_B = (238, 244, 251)
LINEA = (200, 210, 222)
TEXTO = (33, 37, 41)

# --- Columnas: (clave, titulo, ancho px) ---
COLUMNAS = [
    ("nombre", "Nombre", 300),
    ("razon", "Empresa", 210),
    ("marca", "Marca", 150),
    ("tramite", "Trámite", 190),
    ("pm", "PM", 120),
    ("modelo", "Modelo/s", 250),
]

PAD = 10          # padding interno de celda
MARGEN = 16       # margen del lienzo
INTERLINEA = 6    # separacion entre lineas de texto


def _wrap(draw, texto, font, ancho_max):
    """Parte el texto en varias lineas para que entre en 'ancho_max'."""
    texto = texto or ""
    lineas = []
    for parrafo in texto.split("\n"):
        palabras = parrafo.split(" ")
        actual = ""
        for p in palabras:
            prueba = (actual + " " + p).strip()
            if draw.textlength(prueba, font=font) <= ancho_max or not actual:
                if draw.textlength(p, font=font) > ancho_max and not actual:
                    while draw.textlength(p, font=font) > ancho_max and len(p) > 1:
                        corte = len(p)
                        while corte > 1 and draw.textlength(p[:corte], font=font) > ancho_max:
                            corte -= 1
                        lineas.append(p[:corte])
                        p = p[corte:]
                    actual = p
                else:
                    actual = prueba
            else:
                lineas.append(actual)
                actual = p
        lineas.append(actual)
    return lineas or [""]


def render_tabla(registros, titulo, out_path):
    """Dibuja la grilla y la guarda en out_path (PNG)."""
    ancho_cols = [c[2] for c in COLUMNAS]
    ancho_tabla = sum(ancho_cols)
    lh = F_CELDA.size + INTERLINEA  # alto de linea

    tmp = Image.new("RGB", (10, 10))
    d = ImageDraw.Draw(tmp)

    # Precalcular lineas y alto de cada fila
    filas_wrap = []
    for reg in registros:
        celdas = []
        alto_celdas = []
        for clave, _t, w in COLUMNAS:
            wl = _wrap(d, str(reg.get(clave, "")), F_CELDA, w - 2 * PAD)
            celdas.append(wl)
            alto_celdas.append(len(wl))
        filas_wrap.append((celdas, max(alto_celdas)))

    alto_titulo = F_TITULO.size + 2 * PAD + 8
    alto_header = F_HEADER.size + 2 * PAD
    alto_filas = sum(nl * lh + 2 * PAD for _c, nl in filas_wrap)
    W = ancho_tabla + 2 * MARGEN
    H = MARGEN + alto_titulo + alto_header + alto_filas + MARGEN

    img = Image.new("RGB", (W, H), BLANCO)
    draw = ImageDraw.Draw(img)

    # Titulo
    draw.rectangle([MARGEN, MARGEN, MARGEN + ancho_tabla, MARGEN + alto_titulo], fill=AZUL_OSC)
    draw.text((MARGEN + PAD, MARGEN + PAD), titulo, font=F_TITULO, fill=BLANCO)

    y = MARGEN + alto_titulo

    # Header de columnas
    draw.rectangle([MARGEN, y, MARGEN + ancho_tabla, y + alto_header], fill=AZUL)
    x = MARGEN
    for (_k, t, w) in COLUMNAS:
        draw.text((x + PAD, y + PAD), t, font=F_HEADER, fill=BLANCO)
        x += w
    y += alto_header

    # Filas
    for i, (celdas, nl) in enumerate(filas_wrap):
        alto_fila = nl * lh + 2 * PAD
        bg = FILA_A if i % 2 == 0 else FILA_B
        draw.rectangle([MARGEN, y, MARGEN + ancho_tabla, y + alto_fila], fill=bg)
        x = MARGEN
        for ci, (_k, _t, w) in enumerate(COLUMNAS):
            ty = y + PAD
            for linea in celdas[ci]:
                draw.text((x + PAD, ty), linea, font=F_CELDA, fill=TEXTO)
                ty += lh
            x += w
        y += alto_fila

    # Lineas de grilla (verticales)
    x = MARGEN
    for w in ancho_cols:
        draw.line([x, MARGEN + alto_titulo, x, y], fill=LINEA, width=1)
        x += w
    draw.line([x, MARGEN + alto_titulo, x, y], fill=LINEA, width=1)
    # borde
    draw.rectangle([MARGEN, MARGEN, MARGEN + ancho_tabla, y], outline=LINEA, width=1)

    img.save(out_path)
    return out_path


if __name__ == "__main__":
    import json, sys
    datos = json.loads(sys.argv[1]) if len(sys.argv) > 1 else []
    render_tabla(datos, "Boletín ANMAT — 05/08 (muestra)", "muestra.png")
    print("listo: muestra.png")

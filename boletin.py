#!/usr/bin/env python3
"""
Bot del Boletin de ANMAT (Helena - Productos Medicos) -> Telegram.

Lee el listado de productos medicos en https://helena.anmat.gob.ar/Boletin/
y avisa por Telegram cuando aparecen novedades. Ahora revisa LAS DOS solapas
de la pagina:

  - "Registros"      (tabla gvTramites)
  - "Notificaciones" (tabla gvDeclaraciones)  -> declaraciones juradas PM I-II

Cuando hay novedades, en vez de mandar fotos manda UN archivo Excel (.xlsx)
con UNA hoja que junta las dos solapas, con una columna "Tipo"
(Registro / Notificación) y ordenada por Razon Social (para ver juntas todas
las filas de una misma empresa). Cada fila trae: Tipo, Tramite, Fecha, Razon
Social, Nombre, Marca, Modelo/s (resumido a ~4 lineas), PM y Expediente.

Solo avisa lo nuevo; el estado ya avisado se guarda en 'vistos_boletin.json'.

Usa los mismos Secrets que el bot de alertas:
  - TELEGRAM_TOKEN
  - TELEGRAM_CHAT_ID
"""

import os
import re
import sys
import json
import time
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("America/Argentina/Buenos_Aires")
except Exception:
    TZ = None

URL = "https://helena.anmat.gob.ar/Boletin/"
# Las dos solapas de la pagina ya vienen ambas en el HTML inicial (son pestañas
# tipo Bootstrap: #tabRegistros y #tabDeclaraciones), no hace falta postback.
TABLA_REGISTROS = "ctl00_ContentPlaceHolder1_gvTramites"
TABLA_NOTIFICACIONES = "ctl00_ContentPlaceHolder1_gvDeclaraciones"

STATE_FILE = Path("vistos_boletin.json")

TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

# Columnas (mismo orden en las dos tablas para los indices 0..7):
#   0 Tramite | 1 Fecha Fin | 2 Razon Social | 3 Nombre | 4 Marca
#   5 Modelo/s | 6 PM | 7 Expediente | (8/9 links de documento, no se usan)
COL = {
    "tramite": 0,
    "fecha": 1,
    "razon": 2,
    "nombre": 3,
    "marca": 4,
    "modelo": 5,
    "pm": 6,
    "expediente": 7,
}


def _es_expediente(exp):
    """True si 'exp' parece un expediente ANMAT real (ej. 1-0047-3110-004096-26-3).

    Sirve para descartar las filas basura que la grilla ASP.NET agrega al final
    (el paginador '1 2 3 ... >>' y una fila fantasma de conteo), que si no se
    filtran entran como si fueran registros.
    """
    return exp.count("-") >= 3 and any(c.isdigit() for c in exp)


def parse_tabla(html_text, table_id, solapa):
    """Extrae los registros de una tabla (por id) y los etiqueta con la solapa."""
    soup = BeautifulSoup(html_text, "html.parser")
    tabla = soup.find("table", id=table_id)
    items = []
    if tabla is None:
        return items
    for tr in tabla.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 8:
            continue  # encabezado (th) o filas sin datos

        def celda(i):
            return tds[i].get_text(" ", strip=True) if i < len(tds) else ""

        item = {"solapa": solapa}
        for campo, idx in COL.items():
            item[campo] = celda(idx)

        if not _es_expediente(item["expediente"]):
            continue  # descarta paginador / pie de la grilla
        items.append(item)
    return items


class BoletinCaido(Exception):
    """La web de Helena no respondio o no trajo las tablas esperadas."""


# Nombres ASP.NET de las grillas (para pedir la pagina 2, 3... por postback).
_EVENT_TARGET = {
    "Registros": "ctl00$ContentPlaceHolder1$gvTramites",
    "Notificaciones": "ctl00$ContentPlaceHolder1$gvDeclaraciones",
}
_TABLA_ID = {"Registros": TABLA_REGISTROS, "Notificaciones": TABLA_NOTIFICACIONES}

# Cuantas paginas (de ~50 filas) leer como maximo por solapa. Normalmente
# alcanza con la 1; se leen mas solo si TODA la pagina es nueva (dias con
# muchas novedades o recuperacion despues de un dia en que la web fallo).
MAX_PAGINAS = 6

# Reintentos ante errores de red / web caida, antes de darla por caida.
REINTENTOS = 3
ESPERA_REINTENTO = (20, 60)   # segundos entre intentos

_UA = {"User-Agent": "alertas-anmat-bot/1.0 (+github actions)"}


def _con_reintentos(fn, *args, **kwargs):
    ultimo = None
    for intento in range(1, REINTENTOS + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:  # red, timeout, HTTP 5xx, pagina rota...
            ultimo = e
            print(f"Boletin: intento {intento}/{REINTENTOS} fallo: {e!r}")
            if intento < REINTENTOS:
                time.sleep(ESPERA_REINTENTO[min(intento - 1, len(ESPERA_REINTENTO) - 1)])
    raise BoletinCaido(f"Helena no respondio tras {REINTENTOS} intentos: {ultimo!r}")


def _campos_formulario(html):
    """Campos del formulario ASP.NET (VIEWSTATE, filtros...) para hacer postback."""
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form")
    if form is None:
        raise BoletinCaido("la pagina no trae el formulario")
    datos = {}
    for el in form.find_all("input"):
        nombre = el.get("name")
        tipo = (el.get("type") or "text").lower()
        if not nombre or tipo in ("submit", "button", "image", "file"):
            continue
        if tipo in ("checkbox", "radio") and not el.has_attr("checked"):
            continue
        datos[nombre] = el.get("value", "")
    for sel in form.find_all("select"):
        nombre = sel.get("name")
        if not nombre:
            continue
        op = sel.find("option", selected=True) or sel.find("option")
        datos[nombre] = op.get("value", "") if op else ""
    return datos


def _texto(r):
    """HTML como texto, en UTF-8 si el servidor no declara la codificacion
    (para no romper tildes como 'recolección')."""
    if "charset" not in r.headers.get("Content-Type", "").lower():
        r.encoding = "utf-8"
    return r.text


def _get_pagina1(sesion):
    r = sesion.get(URL, timeout=60, headers=_UA)
    r.raise_for_status()
    html = _texto(r)
    if "gvTramites" not in html and "gvDeclaraciones" not in html:
        raise RuntimeError("la pagina cargo pero sin las tablas del Boletin")
    return html


def _post_pagina(sesion, html_anterior, solapa, n):
    datos = _campos_formulario(html_anterior)
    datos["__EVENTTARGET"] = _EVENT_TARGET[solapa]
    datos["__EVENTARGUMENT"] = f"Page${n}"
    r = sesion.post(URL, data=datos, timeout=90, headers=_UA)
    r.raise_for_status()
    html = _texto(r)
    if _TABLA_ID[solapa] not in html:
        raise RuntimeError(f"pagina {n} de {solapa} sin tabla")
    return html


def _hay_pagina(html, solapa, n):
    """True si el paginador de la grilla ofrece la pagina n."""
    return f"'{_EVENT_TARGET[solapa]}','Page${n}'" in html or \
        f"&#39;{_EVENT_TARGET[solapa]}&#39;,&#39;Page${n}&#39;" in html


def _leer_solapa(sesion, html1, solapa, seen):
    """Lee la solapa empezando por la pagina 1 y sigue mientras TODO sea nuevo."""
    items = parse_tabla(html1, _TABLA_ID[solapa], solapa)
    if seen is None:
        return items                      # primera corrida: alcanza con la pag. 1
    if solapa == "Notificaciones" and not any(k.startswith("NOTIF|") for k in seen):
        return items                      # todavia sin estado de Notificaciones
    html, pagina, ultimos = html1, 1, items
    # La grilla viene de la mas nueva a la mas vieja: si las ULTIMAS filas de la
    # pagina (las mas viejas) todavia son nuevas, puede haber mas en la siguiente.
    # (Se miran las ultimas 5 y no solo 1, por si alguna fila suelta repite un
    # expediente ya avisado, ej. una revalida.)
    while (ultimos and any(clave(it) not in seen for it in ultimos[-5:])
           and pagina < MAX_PAGINAS and _hay_pagina(html, solapa, pagina + 1)):
        pagina += 1
        print(f"Boletin: el final de la pagina {pagina - 1} de {solapa} es nuevo -> leo la pagina {pagina}.")
        html = _con_reintentos(_post_pagina, sesion, html, solapa, pagina)
        ultimos = parse_tabla(html, _TABLA_ID[solapa], solapa)
        items.extend(ultimos)
    return items


def fetch_todo(seen=None):
    """Descarga el Boletin y devuelve (registros, notificaciones).

    Reintenta si la web no responde; si sigue caida lanza BoletinCaido.
    Lee paginas extra cuando hay mas novedades que las ~50 filas de la pagina 1.
    """
    sesion = requests.Session()
    html1 = _con_reintentos(_get_pagina1, sesion)
    registros = _leer_solapa(sesion, html1, "Registros", seen)
    notificaciones = _leer_solapa(sesion, html1, "Notificaciones", seen)
    if not registros and not notificaciones:
        raise BoletinCaido("la pagina cargo pero las dos tablas vinieron vacias")
    return registros, notificaciones


def clave(item):
    """Identificador unico e invisible para no repetir avisos.

    - Registros: se mantiene la clave = expediente (compatibilidad con el estado
      ya guardado en vistos_boletin.json; cambiarla reenviaria todo el historial).
    - Notificaciones: clave compuesta con prefijo NOTIF|, porque en esta solapa
      un mismo expediente puede repetirse para productos distintos (ej. 3083-10).
    """
    if item["solapa"] == "Notificaciones":
        return "NOTIF|" + "|".join([
            item["expediente"], item["nombre"], item["marca"], item["fecha"],
        ])
    if item["expediente"]:
        return item["expediente"]
    return "|".join([item["tramite"], item["nombre"], item["pm"], item["marca"]])


def dedup(items):
    """Colapsa filas repetidas (misma clave) conservando el orden."""
    vistos = set()
    out = []
    for it in items:
        k = clave(it)
        if k in vistos:
            continue
        vistos.add(k)
        out.append(it)
    return out


def load_seen():
    """Devuelve el set de claves ya avisadas, o None si es 'primera corrida'.

    Un archivo inexistente, vacio (`[]`) o ilegible se trata como primera corrida:
    en ese caso se muestran SOLO las novedades de hoy y se siembra el resto.
    """
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            return set(data) if data else None
        except Exception:
            return None
    return None


def _hoy():
    """Fecha de hoy en formato del Boletin (DD-MM-YYYY), hora Argentina."""
    ahora = datetime.now(TZ) if TZ else datetime.now()
    return ahora.strftime("%d-%m-%Y")


def save_seen(claves):
    STATE_FILE.write_text(
        json.dumps(sorted(claves), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

# Tipo va primero para distinguir de un vistazo Registro vs Notificación.
# La columna "modelo" (Modelo/s) es la unica con ajuste de texto.
_HEADERS = [
    ("tipo", "Tipo", 14),
    ("tramite", "Trámite", 34),
    ("fecha", "Fecha", 12),
    ("razon", "Razón Social", 36),
    ("nombre", "Nombre", 34),
    ("marca", "Marca", 24),
    ("modelo", "Modelo/s", 60),
    ("pm", "PM", 14),
    ("expediente", "Expediente", 24),
]

# Etiqueta legible de la columna Tipo segun la solapa de origen.
_TIPO = {"Registros": "Registro", "Notificaciones": "Notificación"}

# La columna Modelo/s puede tener 30+ items y hacer filas gigantes. RECORTAMOS
# el texto a ~4 lineas (por caracteres) y agregamos " […]" cuando se corta.
# El detalle completo esta siempre en el Boletin de ANMAT (con PM/Expediente).
# (Antes se probo limitar solo la ALTURA de la fila, pero Excel auto-expande las
#  celdas con ajuste de texto al abrir, asi que el recorte del texto es lo unico
#  que garantiza el resumen.)
MODELO_MAX_CHARS = 220    # ~4 lineas con la columna Modelo/s en ancho 60


def _recortar_modelo(txt, limite=MODELO_MAX_CHARS):
    """Recorta Modelo/s a ~4 lineas, cortando en un espacio y agregando ' […]'."""
    txt = " ".join((txt or "").split())  # colapsa espacios/saltos
    if len(txt) <= limite:
        return txt
    corte = txt[:limite]
    esp = corte.rfind(" ")
    if esp > limite * 0.6:
        corte = corte[:esp]
    return corte.rstrip(" ,;.") + " […]"


# Filas a RESALTAR en el Excel: todo lo que sea "tubos de recolección".
# Se compara sin tildes ni mayusculas y acepta "tubo(s) de/para recolección"
# (ANMAT lo escribe de las dos formas, ej. "Tubos para recolección de sangre")
# y la version en ingles "blood collection tube(s)".
PATRON_RESALTAR = re.compile(
    r"\btubos?\s+(?:de|para)\s+(?:la\s+)?recoleccion|\bblood\s+collection\s+tubes?\b"
)
CAMPOS_RESALTAR = ("nombre", "marca", "modelo", "tramite")
RESALTAR_FILL = PatternFill("solid", fgColor="FFD966")   # amarillo/ambar fuerte
RESALTAR_FONT = Font(bold=True, color="7F2704")          # bordo, negrita


def _normalizar(txt):
    txt = unicodedata.normalize("NFKD", txt or "")
    return "".join(c for c in txt if not unicodedata.combining(c)).lower()


def es_tubo_recoleccion(it):
    """True si la fila habla de tubos de recoleccion (en cualquiera de sus campos)."""
    return any(PATRON_RESALTAR.search(_normalizar(it.get(c, ""))) for c in CAMPOS_RESALTAR)


def _orden(it):
    """Ordena por Razón Social (A→Z, sin distinguir may/min), luego Tipo y Nombre."""
    return (
        (it.get("razon", "") or "").strip().upper(),
        it.get("tipo", ""),
        (it.get("nombre", "") or "").strip().upper(),
    )


def construir_excel(nuevos_reg, nuevos_notif, path):
    """Crea un .xlsx con UNA hoja que junta las dos solapas.

    Columna "Tipo" (Registro / Notificación) para diferenciarlas y ordenado
    por Razón Social, para ver juntas todas las filas de una misma empresa.
    """
    filas = []
    for it in nuevos_reg:
        filas.append({**it, "tipo": _TIPO["Registros"]})
    for it in nuevos_notif:
        filas.append({**it, "tipo": _TIPO["Notificaciones"]})
    filas.sort(key=_orden)

    wb = Workbook()
    ws = wb.active
    ws.title = "Novedades"

    header_fill = PatternFill("solid", fgColor="2E6DA4")
    header_font = Font(bold=True, color="FFFFFF")
    ws.append([h[1] for h in _HEADERS])
    for ci, (_k, _t, ancho) in enumerate(_HEADERS, 1):
        c = ws.cell(row=1, column=ci)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(ci)].width = ancho

    col_modelo = [h[0] for h in _HEADERS].index("modelo") + 1
    for it in filas:
        fila = []
        for (k, _t, _w) in _HEADERS:
            v = it.get(k, "")
            if k == "modelo":
                v = _recortar_modelo(v)   # ~4 lineas, con " […]" si se corta
            fila.append(v)
        ws.append(fila)
        r = ws.max_row
        # Ajuste de texto en Modelo/s (por nombre de encabezado, robusto al orden).
        ws.cell(row=r, column=col_modelo).alignment = Alignment(wrap_text=True, vertical="top")
        # Resaltar toda la fila si es "tubos de recolección" (se evalua sobre el
        # texto completo, antes del recorte de Modelo/s).
        if es_tubo_recoleccion(it):
            for ci in range(1, len(_HEADERS) + 1):
                c = ws.cell(row=r, column=ci)
                c.fill = RESALTAR_FILL
                c.font = RESALTAR_FONT

    ws.freeze_panes = "A2"
    ultima_col = get_column_letter(len(_HEADERS))
    ws.auto_filter.ref = f"A1:{ultima_col}{max(ws.max_row, 1)}"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------

def send_telegram(text):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    resp = requests.post(
        url,
        data={"chat_id": CHAT_ID, "text": text, "parse_mode": "HTML"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def send_document(path, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    with open(path, "rb") as f:
        resp = requests.post(
            url,
            data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "HTML"},
            files={"document": (os.path.basename(path), f)},
            timeout=120,
        )
    resp.raise_for_status()
    return resp.json()


def _fmt_fecha(f):
    """'18-11-2026' o '2026-11-18' -> '18/11'."""
    p = f.split("-")
    if len(p) == 3:
        return f"{p[2]}/{p[1]}" if len(p[0]) == 4 else f"{p[0]}/{p[1]}"
    return f


def enviar_excel(nuevos_reg, nuevos_notif, recuperado=None):
    """Arma el Excel con las novedades y lo manda como documento por Telegram.

    recuperado: lista de fechas (YYYY-MM-DD) en que la web fallo y cuyas
    novedades vienen incluidas en este envio.
    """
    hoy = datetime.now(TZ).strftime("%Y-%m-%d") if TZ else datetime.now().strftime("%Y-%m-%d")
    path = f"boletin_anmat_{hoy}.xlsx"
    construir_excel(nuevos_reg, nuevos_notif, path)
    total = len(nuevos_reg) + len(nuevos_notif)
    lineas = []
    if recuperado:
        dias = ", ".join(_fmt_fecha(d) for d in recuperado)
        lineas.append(f"🔄 <b>Volvió la web de Helena.</b> Incluye lo pendiente del {dias}.\n")
    lineas += [
        f"🆕 <b>Boletín ANMAT — {total} novedad(es)</b>",
        f"• Registros: {len(nuevos_reg)}",
        f"• Notificaciones: {len(nuevos_notif)}",
    ]
    por_fecha = Counter(it.get("fecha", "") for it in nuevos_reg + nuevos_notif)
    if len(por_fecha) > 1:
        def _k(f):
            p = f.split("-")
            return tuple(reversed(p)) if len(p) == 3 else (f,)
        detalle = " | ".join(f"{_fmt_fecha(f)}: {n}" for f, n in sorted(por_fecha.items(), key=lambda x: _k(x[0])))
        lineas.append(f"📅 Por fecha: {detalle}")
    n_tubos = sum(1 for it in nuevos_reg + nuevos_notif if es_tubo_recoleccion(it))
    if n_tubos:
        lineas.append(f"\n🩸 <b>Tubos de recolección: {n_tubos}</b> (resaltados en amarillo)")
    send_document(path, "\n".join(lineas))


# ---------------------------------------------------------------------------
# Logica principal
# ---------------------------------------------------------------------------

def revisar(recuperado=None):
    """Revisa las dos solapas, avisa lo nuevo (Excel) y devuelve (n_reg, n_notif).

    Si la web esta caida lanza BoletinCaido (sin enviar ni guardar nada), para
    que el orquestador avise y lo reintente despues.
    """
    seen = load_seen()
    hoy = _hoy()
    registros, notificaciones = fetch_todo(seen)
    print(
        f"Boletin: {len(registros)} registros y {len(notificaciones)} "
        f"notificaciones leidos."
    )

    # Primera corrida (sin estado / estado vacio): mostramos SOLO las novedades
    # de HOY en las dos solapas y sembramos el resto (dias anteriores) en
    # silencio, para no volcar el backlog historico.
    if seen is None:
        nuevos_reg = dedup([r for r in registros if r.get("fecha") == hoy])
        nuevos_notif = dedup([n for n in notificaciones if n.get("fecha") == hoy])
        if nuevos_reg or nuevos_notif:
            enviar_excel(nuevos_reg, nuevos_notif)
        else:
            send_telegram(
                "✅ <b>Aviso del Boletín ANMAT activado</b>\n\n"
                "Reviso las solapas <b>Registros</b> y <b>Notificaciones</b> y te "
                "aviso las novedades en un archivo Excel. Hoy todavía no hay "
                "novedades nuevas."
            )
        save_seen({clave(r) for r in registros} | {clave(n) for n in notificaciones})
        print(f"Boletin: primera corrida, {len(nuevos_reg)} registros y {len(nuevos_notif)} notificaciones de hoy.")
        return len(nuevos_reg), len(nuevos_notif)

    # ¿Ya venimos siguiendo Notificaciones? (la solapa se sumo despues).
    hay_estado_notif = any(k.startswith("NOTIF|") for k in seen)

    nuevos_reg = dedup([r for r in registros if clave(r) not in seen])

    if hay_estado_notif:
        nuevos_notif = dedup([n for n in notificaciones if clave(n) not in seen])
    else:
        # Primera vez con Notificaciones sobre un estado ya existente: mostramos
        # las de HOY y sembramos las de dias anteriores en silencio.
        nuevos_notif = dedup([n for n in notificaciones if n.get("fecha") == hoy])

    if nuevos_reg or nuevos_notif:
        enviar_excel(nuevos_reg, nuevos_notif, recuperado)

    # Guardar estado con las claves actuales de las dos solapas (siembra
    # Notificaciones en esta corrida y mantiene fresca la ventana de vistos).
    seen |= {clave(r) for r in registros} | {clave(n) for n in notificaciones}
    save_seen(seen)

    print(f"Boletin: {len(nuevos_reg)} registros nuevos, {len(nuevos_notif)} notificaciones nuevas.")
    return len(nuevos_reg), len(nuevos_notif)


def main():
    if not TOKEN or not CHAT_ID:
        print("ERROR: faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID.")
        sys.exit(1)
    revisar()


if __name__ == "__main__":
    main()

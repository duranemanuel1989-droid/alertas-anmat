#!/usr/bin/env python3
"""
Bot del Boletin de ANMAT (Helena - Productos Medicos) -> Telegram.

Lee el listado de tramites/registros de productos medicos en
https://helena.anmat.gob.ar/Boletin/ y avisa por Telegram (como foto tipo
grilla) cuando aparecen registros nuevos.

De cada registro muestra: Nombre, Empresa (Razon Social), Marca, Tramite,
PM y Modelo/s (resumido a las primeras palabras). El numero de Expediente se
usa solo por dentro para no repetir avisos; no se muestra.

Si hay muchos registros nuevos, se parte en varias fotos (12 por imagen).

Los registros ya avisados se guardan en 'vistos_boletin.json'.

Usa los mismos Secrets que el bot de alertas:
  - TELEGRAM_TOKEN
  - TELEGRAM_CHAT_ID
"""

import os
import sys
import json
from pathlib import Path

import requests
from bs4 import BeautifulSoup

import tabla_img

URL = "https://helena.anmat.gob.ar/Boletin/"
TABLE_ID = "ctl00_ContentPlaceHolder1_gvTramites"
STATE_FILE = Path("vistos_boletin.json")

TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

# Caracteres del campo Modelo/s que se muestran en la grilla (el resto se corta).
MODELO_LIMITE = 70
# Cuantos registros entran por foto (para que cada imagen se lea bien).
FILAS_POR_IMAGEN = 12


def parse_boletin(html_text):
    """Extrae la lista de registros de la tabla de tramites."""
    soup = BeautifulSoup(html_text, "html.parser")
    tabla = soup.find("table", id=TABLE_ID)
    if tabla is None:
        tabla = soup.find("table")  # fallback
    items = []
    if tabla is None:
        return items
    for tr in tabla.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 8:
            continue  # saltea encabezado (th) y filas raras
        def celda(i):
            return tds[i].get_text(" ", strip=True) if i < len(tds) else ""
        items.append({
            "tramite": celda(0),
            "razon": celda(2),
            "nombre": celda(3),
            "marca": celda(4),
            "modelo": celda(5),
            "pm": celda(6),
            "expediente": celda(7),
        })
    return items


def fetch_boletin():
    r = requests.get(
        URL,
        timeout=45,
        headers={"User-Agent": "alertas-anmat-bot/1.0 (+github actions)"},
    )
    r.raise_for_status()
    return parse_boletin(r.content)


def clave(item):
    """Identificador unico e invisible para no repetir avisos."""
    if item["expediente"]:
        return item["expediente"]
    return "|".join([item["tramite"], item["nombre"], item["pm"], item["marca"]])


def load_seen():
    if STATE_FILE.exists():
        try:
            return set(json.loads(STATE_FILE.read_text(encoding="utf-8")))
        except Exception:
            return set()
    return None


def save_seen(claves):
    STATE_FILE.write_text(
        json.dumps(sorted(claves), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def resumir_modelo(txt, limite=MODELO_LIMITE):
    txt = " ".join((txt or "").split())
    if len(txt) <= limite:
        return txt
    return txt[:limite].rstrip() + "…"


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    resp = requests.post(
        url,
        data={"chat_id": CHAT_ID, "text": text, "parse_mode": "HTML"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def send_photo(path, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendPhoto"
    with open(path, "rb") as f:
        resp = requests.post(
            url,
            data={"chat_id": CHAT_ID, "caption": caption},
            files={"photo": f},
            timeout=60,
        )
    resp.raise_for_status()
    return resp.json()


def enviar_digest(nuevos):
    """Arma una o varias fotos tipo grilla con los registros nuevos."""
    grupos = [nuevos[i:i + FILAS_POR_IMAGEN] for i in range(0, len(nuevos), FILAS_POR_IMAGEN)]
    total = len(grupos)
    for idx, grupo in enumerate(grupos, 1):
        filas = [{
            "nombre": r["nombre"],
            "razon": r["razon"],
            "marca": r["marca"],
            "tramite": r["tramite"],
            "pm": r["pm"],
            "modelo": resumir_modelo(r["modelo"]),
        } for r in grupo]
        parte = f" — parte {idx} de {total}" if total > 1 else ""
        titulo = f"Nuevos registros — Boletín ANMAT ({len(nuevos)} nuevos){parte}"
        salida = f"tabla_boletin_{idx}.png"
        tabla_img.render_tabla(filas, titulo, salida)
        send_photo(salida, f"🆕 Nuevos registros de productos médicos (Boletín ANMAT){parte}")


def main():
    if not TOKEN or not CHAT_ID:
        print("ERROR: faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID.")
        sys.exit(1)

    registros = fetch_boletin()
    print(f"Se encontraron {len(registros)} registros en la pagina del Boletin.")

    seen = load_seen()
    claves_actuales = {clave(r) for r in registros}

    # Primera ejecucion: guardamos el estado y avisamos que quedo activo,
    # sin mandar los registros que ya estan publicados.
    if seen is None:
        save_seen(claves_actuales)
        msg = (
            "✅ <b>Aviso de nuevos registros (Boletín ANMAT) activado</b>\n\n"
            "Te voy a avisar con una foto tipo grilla cuando se registren nuevos "
            "productos médicos.\n"
            f"Ahora mismo hay {len(registros)} registros recientes en la lista; "
            "a partir de acá solo te aviso los que sean nuevos."
        )
        send_telegram(msg)
        print("Primera ejecucion: estado inicial guardado y aviso enviado.")
        return

    nuevos = [r for r in registros if clave(r) not in seen]
    if nuevos:
        enviar_digest(nuevos)
        save_seen(seen | claves_actuales)

    print(f"Se avisaron {len(nuevos)} registros nuevos.")


if __name__ == "__main__":
    main()

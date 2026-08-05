#!/usr/bin/env python3
"""
Bot de alertas de ANMAT -> Telegram.

Revisa la pagina de alertas de ANMAT (https://www.argentina.gob.ar/anmat/alertas)
y, cada vez que aparece una alerta nueva, la envia a un chat de Telegram.

Las alertas que ya se avisaron se guardan en 'vistos.json' para no repetirlas.

Necesita dos variables de entorno (se cargan como Secrets en GitHub):
  - TELEGRAM_TOKEN   : el token del bot (de @BotFather)
  - TELEGRAM_CHAT_ID : el id del chat donde llegan los avisos
"""

import os
import sys
import json
import html as htmllib
from pathlib import Path

import requests
from bs4 import BeautifulSoup

BASE = "https://www.argentina.gob.ar"
ALERTS_URL = f"{BASE}/anmat/alertas"
STATE_FILE = Path("vistos.json")

TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "").strip()


def parse_alertas(html_text):
    """Extrae la lista de alertas del HTML de la pagina de ANMAT."""
    soup = BeautifulSoup(html_text, "html.parser")
    items = []
    vistos_en_pagina = set()
    for a in soup.select('a.panel[href*="/noticias/"]'):
        href = a.get("href", "").strip()
        if not href:
            continue
        link = href if href.startswith("http") else BASE + href
        if link in vistos_en_pagina:
            continue
        vistos_en_pagina.add(link)

        time_el = a.find("time")
        h3 = a.find("h3")
        fecha = time_el.get_text(strip=True) if time_el else ""
        titulo = h3.get_text(strip=True) if h3 else a.get_text(" ", strip=True)
        items.append({"url": link, "titulo": titulo, "fecha": fecha})
    return items


def fetch_alertas():
    """Descarga la pagina de alertas y devuelve la lista parseada."""
    r = requests.get(
        ALERTS_URL,
        timeout=30,
        headers={"User-Agent": "alertas-anmat-bot/1.0 (+github actions)"},
    )
    r.raise_for_status()
    return parse_alertas(r.text)


def load_seen():
    """Devuelve el set de URLs ya avisadas, o None si es la primera ejecucion."""
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            return set(data)
        except Exception:
            return set()
    return None


def save_seen(urls):
    STATE_FILE.write_text(
        json.dumps(sorted(urls), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def send_telegram(text):
    """Envia un mensaje al chat de Telegram configurado."""
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    resp = requests.post(
        url,
        data={
            "chat_id": CHAT_ID,
            "text": text,
            "parse_mode": "HTML",
            "disable_web_page_preview": "false",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def format_alerta(item):
    titulo = htmllib.escape(item["titulo"])
    fecha = htmllib.escape(item["fecha"])
    return (
        "\U0001F6A8 <b>Nueva alerta de ANMAT</b>\n\n"
        f"<b>{titulo}</b>\n"
        f"\U0001F4C5 {fecha}\n\n"
        f"\U0001F517 {item['url']}"
    )


def main():
    if not TOKEN or not CHAT_ID:
        print("ERROR: faltan las variables TELEGRAM_TOKEN o TELEGRAM_CHAT_ID.")
        sys.exit(1)

    alertas = fetch_alertas()
    print(f"Se encontraron {len(alertas)} alertas en la pagina de ANMAT.")

    seen = load_seen()
    current_urls = {a["url"] for a in alertas}

    # Primera ejecucion: no hay historial todavia.
    # Guardamos el estado actual y mandamos un unico mensaje de bienvenida
    # (con la ultima alerta de ejemplo) en vez de spamear todas las de golpe.
    if seen is None:
        save_seen(current_urls)
        if alertas:
            ultima = alertas[0]
            msg = (
                "✅ <b>Bot de alertas de ANMAT activado</b>\n\n"
                "A partir de ahora te voy a avisar por aca cada vez que ANMAT "
                "publique una nueva alerta.\n\n"
                "Como ejemplo, esta es la ultima alerta publicada:\n\n"
                f"<b>{htmllib.escape(ultima['titulo'])}</b>\n"
                f"\U0001F4C5 {htmllib.escape(ultima['fecha'])}\n"
                f"\U0001F517 {ultima['url']}"
            )
        else:
            msg = (
                "✅ <b>Bot de alertas de ANMAT activado</b>\n\n"
                "Por ahora no hay alertas para mostrar, pero ya estoy vigilando."
            )
        send_telegram(msg)
        print("Primera ejecucion: estado inicial guardado y bienvenida enviada.")
        return

    # Ejecuciones siguientes: avisar solo lo nuevo.
    nuevas = [a for a in alertas if a["url"] not in seen]

    # Enviar de la mas vieja a la mas nueva, para que lleguen en orden.
    for item in reversed(nuevas):
        send_telegram(format_alerta(item))
        print(f"Enviada alerta: {item['titulo']}")

    if nuevas:
        save_seen(seen | current_urls)

    print(f"Se enviaron {len(nuevas)} alertas nuevas.")


if __name__ == "__main__":
    main()

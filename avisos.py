#!/usr/bin/env python3
"""
Aviso diario unificado de ANMAT -> Telegram.

Una vez por dia (18:00 hora Argentina) revisa las dos fuentes:
  - Alertas de ANMAT  (texto)
  - Boletin de productos medicos  (foto tipo grilla)

Manda todo lo nuevo junto. Si no hay NADA nuevo en ninguna de las dos,
manda un unico mensaje: "No hay avisos de ANMAT por hoy".

Reutiliza la logica ya probada de alertas.py y boletin.py.
"""

import sys

import alertas
import boletin

TOKEN = boletin.TOKEN
CHAT_ID = boletin.CHAT_ID


def revisar_alertas():
    """Envia las alertas nuevas (texto) y devuelve cuantas fueron."""
    registros = alertas.fetch_alertas()
    seen = alertas.load_seen() or set()
    nuevas = [a for a in registros if a["url"] not in seen]
    for item in reversed(nuevas):  # de la mas vieja a la mas nueva
        alertas.send_telegram(alertas.format_alerta(item))
    if nuevas:
        alertas.save_seen(seen | {a["url"] for a in registros})
    return len(nuevas)


def revisar_boletin():
    """Envia los registros nuevos (foto grilla) y devuelve cuantos fueron."""
    registros = boletin.fetch_boletin()
    seen = boletin.load_seen() or set()
    nuevos = [r for r in registros if boletin.clave(r) not in seen]
    if nuevos:
        boletin.enviar_digest(nuevos)
        boletin.save_seen(seen | {boletin.clave(r) for r in registros})
    return len(nuevos)


def main():
    if not TOKEN or not CHAT_ID:
        print("ERROR: faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID.")
        sys.exit(1)

    n_alertas = revisar_alertas()
    n_boletin = revisar_boletin()
    print(f"Alertas nuevas: {n_alertas} | Registros nuevos del Boletin: {n_boletin}")

    if n_alertas == 0 and n_boletin == 0:
        boletin.send_telegram("📭 <b>No hay avisos de ANMAT por hoy.</b>")
        print("Sin novedades: se aviso 'No hay avisos por hoy'.")


if __name__ == "__main__":
    main()

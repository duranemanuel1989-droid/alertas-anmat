#!/usr/bin/env python3
"""
Aviso diario unificado de ANMAT -> Telegram.

Una vez por dia (18:00 hora Argentina) revisa las dos fuentes:
  - Alertas de ANMAT  (texto)
  - Boletin de productos medicos  (Excel: Registros + Notificaciones, columna Tipo,
    ordenado por Razon Social, con los "tubos de recolección" resaltados)

Manda todo lo nuevo junto. Si no hay NADA nuevo en ninguna de las dos,
manda un unico mensaje: "No hay avisos de ANMAT por hoy".

Si una web de ANMAT esta caida:
  - ese dia avisa por Telegram "⚠️ Hoy falló la web de ANMAT ..." (una sola vez
    por dia, aunque se reintente varias veces),
  - guarda la fecha en 'fallos.json',
  - y en la siguiente corrida que funcione (reintento o el aviso del otro dia)
    manda las novedades de TODOS los dias pendientes juntas, avisando que es
    una recuperacion.

MODO (variable de entorno, la pone el workflow):
  - "diario"    : el aviso de las 18:00 (o ejecucion manual). Revisa todo.
  - "reintento" : corridas extra. Solo revisan las fuentes que quedaron
                  pendientes por una caida; si no hay nada pendiente, no hacen
                  nada (no mandan ningun mensaje).

Reutiliza la logica ya probada de alertas.py y boletin.py.
"""

import os
import sys
import json
import time
from datetime import datetime
from pathlib import Path

import alertas
import boletin

TOKEN = boletin.TOKEN
CHAT_ID = boletin.CHAT_ID
MODO = os.environ.get("MODO", "diario").strip().lower() or "diario"

FALLOS_FILE = Path("fallos.json")
NOMBRE = {"boletin": "Boletín de productos médicos (Helena)", "alertas": "Alertas de ANMAT"}


# ---------------------------------------------------------------------------
# Estado de caidas: {"boletin": {"pendientes": ["2026-11-18"], "avisado": "2026-11-18"}, ...}
# ---------------------------------------------------------------------------

def _hoy_iso():
    ahora = datetime.now(boletin.TZ) if boletin.TZ else datetime.now()
    return ahora.strftime("%Y-%m-%d")


def cargar_fallos():
    try:
        data = json.loads(FALLOS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def guardar_fallos(fallos):
    limpio = {k: v for k, v in fallos.items() if v.get("pendientes")}
    FALLOS_FILE.write_text(json.dumps(limpio, ensure_ascii=False, indent=2), encoding="utf-8")


def pendientes(fallos, fuente):
    return list(fallos.get(fuente, {}).get("pendientes", []))


# ---------------------------------------------------------------------------
# Fuentes
# ---------------------------------------------------------------------------

def _fetch_alertas_con_reintentos():
    ultimo = None
    for intento in range(1, boletin.REINTENTOS + 1):
        try:
            return alertas.fetch_alertas()
        except Exception as e:
            ultimo = e
            print(f"Alertas: intento {intento}/{boletin.REINTENTOS} fallo: {e!r}")
            if intento < boletin.REINTENTOS:
                time.sleep(boletin.ESPERA_REINTENTO[min(intento - 1, len(boletin.ESPERA_REINTENTO) - 1)])
    raise RuntimeError(f"la web de alertas no respondio: {ultimo!r}")


def revisar_alertas(recuperado=None):
    """Envia las alertas nuevas (texto) y devuelve cuantas fueron."""
    registros = _fetch_alertas_con_reintentos()
    if not registros:
        raise RuntimeError("la pagina de alertas cargo vacia")
    seen = alertas.load_seen() or set()
    nuevas = [a for a in registros if a["url"] not in seen]
    if recuperado and nuevas:
        dias = ", ".join(boletin._fmt_fecha(d) for d in recuperado)
        boletin.send_telegram(
            f"🔄 <b>Volvió la web de Alertas de ANMAT.</b> Te mando lo pendiente del {dias}."
        )
    for item in reversed(nuevas):  # de la mas vieja a la mas nueva
        alertas.send_telegram(alertas.format_alerta(item))
    if nuevas:
        alertas.save_seen(seen | {a["url"] for a in registros})
    return len(nuevas)


def revisar_boletin(recuperado=None):
    """Envia las novedades del Boletin (Excel) y devuelve cuantas fueron en total."""
    n_reg, n_notif = boletin.revisar(recuperado=recuperado)
    return n_reg + n_notif


REVISAR = {"alertas": revisar_alertas, "boletin": revisar_boletin}


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def main():
    if not TOKEN or not CHAT_ID:
        print("ERROR: faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID.")
        sys.exit(1)

    hoy = _hoy_iso()
    fallos = cargar_fallos()
    print(f"Modo: {MODO} | pendientes: { {k: pendientes(fallos, k) for k in REVISAR} }")

    if MODO == "reintento":
        fuentes = [f for f in REVISAR if pendientes(fallos, f)]
        if not fuentes:
            print("Reintento: no hay ninguna caida pendiente, no hago nada.")
            return
    else:
        fuentes = list(REVISAR)

    resultados = {}        # fuente -> cantidad de novedades
    caidas = []            # fuentes que fallaron en esta corrida
    recuperadas = {}       # fuente -> dias que estaban pendientes y ya se mandaron

    for fuente in fuentes:
        pend = pendientes(fallos, fuente)
        try:
            resultados[fuente] = REVISAR[fuente](recuperado=pend or None)
            if pend:
                recuperadas[fuente] = pend
                fallos.pop(fuente, None)
        except Exception as e:
            print(f"FALLO {fuente}: {e!r}")
            caidas.append(fuente)
            info = fallos.setdefault(fuente, {"pendientes": []})
            if hoy not in info["pendientes"]:
                info["pendientes"].append(hoy)

    # Aviso de caida: UNA vez por dia y por fuente (los reintentos no repiten).
    nuevas_caidas = [f for f in caidas if fallos[f].get("avisado") != hoy]
    if nuevas_caidas:
        lineas = ["⚠️ <b>Hoy falló la web de ANMAT</b>"]
        for f in nuevas_caidas:
            dias = fallos[f]["pendientes"]
            extra = ""
            if len(dias) > 1:
                extra = f" — sigue caída desde el {boletin._fmt_fecha(dias[0])}"
            lineas.append(f"• {NOMBRE[f]}{extra}")
        lineas.append(
            "\nVuelvo a intentar más tarde y, cuando la página se restablezca, "
            "te mando las novedades de todos los días pendientes juntas."
        )
        boletin.send_telegram("\n".join(lineas))
        for f in nuevas_caidas:
            fallos[f]["avisado"] = hoy

    # Se recupero pero no habia nada nuevo: avisarlo igual para cerrar el tema.
    for f, dias in recuperadas.items():
        if resultados.get(f, 0) == 0:
            boletin.send_telegram(
                f"✅ <b>Volvió la web: {NOMBRE[f]}.</b>\n"
                f"No hubo novedades en lo pendiente del "
                f"{', '.join(boletin._fmt_fecha(d) for d in dias)}."
            )

    guardar_fallos(fallos)
    print(f"Resultados: {resultados} | caidas: {caidas} | recuperadas: {recuperadas}")

    if MODO != "reintento" and not caidas and not recuperadas and sum(resultados.values()) == 0:
        boletin.send_telegram("📭 <b>No hay avisos de ANMAT por hoy.</b>")
        print("Sin novedades: se aviso 'No hay avisos por hoy'.")


if __name__ == "__main__":
    main()

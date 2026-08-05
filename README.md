# alertas-anmat

Bot que revisa la página de **alertas de ANMAT**
(https://www.argentina.gob.ar/anmat/alertas) y avisa por **Telegram**
cada vez que aparece una alerta nueva (medicamentos, alimentos, productos
médicos, cosméticos, etc.).

Funciona solo, gratis, usando **GitHub Actions**: cada unas horas revisa la
página y, si hay algo nuevo, te manda el aviso al chat de Telegram.

## Cómo está armado

- `alertas.py` — el programa que revisa las alertas y las envía a Telegram.
- `requirements.txt` — las librerías que necesita (requests y beautifulsoup4).
- `.github/workflows/alertas.yml` — la tarea automática de GitHub que ejecuta
  el programa cada 6 horas.
- `vistos.json` — se crea solo. Guarda las alertas que ya se avisaron para no
  repetirlas.

## Configuración (ya hecha)

El bot necesita dos datos guardados como **Secrets** del repositorio
(en *Settings → Secrets and variables → Actions*):

- `TELEGRAM_TOKEN` — el token del bot de Telegram (de @BotFather).
- `TELEGRAM_CHAT_ID` — el id del chat donde llegan los avisos.

## Cómo probarlo a mano

En la pestaña **Actions**, elegí el workflow *"Alertas ANMAT"* y tocá
**Run workflow**. En la primera ejecución te llega un mensaje de bienvenida;
a partir de ahí solo te avisa cuando hay alertas nuevas.

## Notas

- El horario de la tarea automática está en UTC (Argentina es UTC-3).
- GitHub pausa las tareas automáticas si el repositorio queda 60 días sin
  actividad; con que entres cada tanto o corras el workflow a mano alcanza.
- Cambiar cada cuánto revisa: editá la línea `cron` en
  `.github/workflows/alertas.yml`.
# alertas-anmat

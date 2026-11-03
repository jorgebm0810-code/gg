# Predictor Real Madrid (LaLiga) – v2 con despliegue en Netlify

## Cómo funciona
`update.py --site` descarga resultados (football-data.co.uk, gratis y sin clave), calcula Elo + Poisson,
hace el backtest y genera la carpeta `site/` con `index.html` (dashboard), `data.json` y el Excel.
Netlify ejecuta ese comando en cada despliegue (`netlify.toml`) y una función programada
(`netlify/functions/actualizacion-diaria.mjs`) lanza un despliegue automático periódicamente.

## Puesta en marcha (una vez)
1. Sube esta carpeta a un repositorio de GitHub/GitLab.
2. En Netlify: *Add new site → Import an existing project* → elige el repositorio. Los ajustes salen de `netlify.toml`.
3. Crea un build hook: *Site configuration → Build & deploy → Build hooks → Add build hook* (rama principal) y copia la URL.
4. *Environment variables*: añade `BUILD_HOOK_URL` con esa URL. Opcional: `API_FOOTBALL_KEY` (lesiones y estadísticas individuales).
5. Haz un primer despliegue manual (*Deploys → Trigger deploy*). Cuando termine, tu panel estará en la URL del sitio.

## Frecuencia y coste
La función programada está en `0 6 */2 * *` (cada 2 días, 06:00 UTC). En los planes gratuitos nuevos de Netlify cada
despliegue de producción cuesta 15 créditos y se incluyen 300 al mes: cada 2 días son unos 225; a diario (`0 6 * * *`)
serían unos 450 y se pausaría el sitio. Las cuentas antiguas (plan "legacy") usan otro sistema; comprueba el tuyo en Netlify.
Cambia la frecuencia editando `schedule` en el `.mjs`.

## Lesiones y estado de forma
- **Manual (siempre funciona):** edita `bajas.csv` (Estado = Baja/Duda/Sanción, Importancia 0–1) y haz commit; se redespliega.
- **Automático opcional:** `API_FOOTBALL_KEY`. El plan gratuito puede no cubrir la temporada actual; si no, el script avisa y usa `bajas.csv`.

## Uso en local
```
pip install -r requirements.txt
python update.py --site        # genera ./site
python update.py               # solo el Excel
```

## Qué está verificado y qué no
- Verificado con prensa: victorias 2016-17 → 2025-26 y estado 2026-27 hasta la jornada 11 (9V-0E-2D, 27 pts).
- Probado con datos sintéticos: el script, el backtest, la web y el Excel generan sin errores.
  **Sin probar todavía** contra las descargas reales ni en un despliegue real de Netlify: si el primer build falla,
  copia el log de *Deploys* y lo corregimos. Si la descarga falla, la web se publica con los datos verificados y el aviso correspondiente.
- El ajuste por bajas no está calibrado (no hay histórico gratuito de lesiones). xG no incluido.
- El calendario gratuito solo da la jornada próxima; el resto de partidos se calculan por descarte y sin fecha.

# glp-data

Datos oficiales del **precio del GLP en Colombia**, refrescados a diario por GitHub Actions.
Fuente de la herramienta en vivo: **https://digiautom.com/glp** — Automatizaciones Digitales.

- `glp-data.json` — molécula del productor (precio de publicación oficial, por fuente) y, si hay llave EIA, Mont Belvieu.
- `update-glp-data.js` — el refrescador (la fuente nacional no requiere llave; EIA opcional vía secret `EIA_API_KEY`).
- La página lo lee vía `raw.githubusercontent.com` (público, CORS abierto). La TRM va en vivo aparte (Banrep).

Solo **datos públicos** de publicación oficial (CREG / Superservicios / EIA). No contiene información
confidencial ni material interno de ninguna empresa.

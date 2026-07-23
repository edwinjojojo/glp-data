# glp-data

Datos oficiales del **precio del GLP en Colombia**, refrescados a diario por GitHub Actions.
Fuente de la herramienta en vivo: **https://automatizaciones-digitales.web.app/glp** — Automatizaciones Digitales S.A.S.

- `glp-data.json` — molécula del productor (Ecopetrol PME-VPRECIOS, por fuente) y, si hay llave EIA, Mont Belvieu.
- `update-glp-data.js` — el refrescador (sin llave para Ecopetrol; EIA opcional vía secret `EIA_API_KEY`).
- La página lo lee vía `raw.githubusercontent.com` (público, CORS abierto). La TRM va en vivo aparte (Banrep).

Solo **datos públicos** (CREG/Superservicios/EIA/Ecopetrol OPC). No contiene información confidencial.

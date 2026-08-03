# glp-data

Datos del **precio del GLP en Colombia**, refrescados a diario por GitHub Actions.
Consumidos por la herramienta en vivo **https://digiautom.com/glp** — Automatizaciones Digitales.

## Qué contiene

- `glp-data.json` — dos bloques:
  - **Molécula del productor**, por fuente de producción. Sale del archivo de precios de
    publicación oficial que Ecopetrol publica abiertamente en su web (PME-VPRECIOS, `.xls`,
    sin credenciales ni registro).
  - **Mont Belvieu** (propano/butano) desde la API v2 de la EIA de Estados Unidos, solo si
    existe el secret `EIA_API_KEY`; si no, ese bloque se omite.
- `update-glp-data.js` — el refrescador. Es el único que hace red: descarga esas dos fuentes
  y nada más.

La página lo lee vía `raw.githubusercontent.com` (público, CORS abierto). La TRM se consulta
aparte y en vivo contra el Banco de la República.

## Sobre los datos

Todo lo que hay aquí es **información de publicación oficial y acceso público**: se descarga
sin autenticación y cualquiera puede obtener lo mismo desde la fuente. No hay material
interno, reservado ni confidencial de ninguna empresa, ni datos personales.

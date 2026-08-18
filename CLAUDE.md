# Salmones Austral — Activación de Campañas OnSign TV

Herramienta interna (español) para que encargados de área activen/detengan campañas de OnSign TV bajo demanda, y para que Filete reporte resultados de turno que se muestran en pantalla física vía Google Slides.

**Lee `ONBOARDING.md` para el contexto completo** (IDs de recursos, endpoints, decisiones, pendientes). Este archivo es solo lo esencial para trabajar en el código.

## Stack y estructura

Sitio estático sin build ni dependencias. Todo el frontend vive en un único `index.html` (HTML + CSS + JS inline).

```
index.html       # toda la página
assets/          # logo + tipografía Overpass
appscript_v2/    # backend en uso (Apps Script, desplegado con clasp)
appscript/       # backend viejo RETIRADO — no tocar, es solo referencia histórica
```

## Deploy

- **Frontend**: push a `main` → GitHub Pages publica solo en 1-2 min. No hay pipeline.
- **Backend**: desde `appscript_v2/`, `clasp push --force` y luego `clasp deploy --deploymentId AKfycbzwEuz_J3PMHXpj0HLjddFao8WxGAZGb8DSrWkIxDAjBj0ojmdKMFUBvQuieOsGX2li`. **Siempre el mismo deploymentId**, si no cambia la URL y hay que actualizarla en `index.html` y en OnSign.

## Convenciones

- Toda la UI y los mensajes al usuario van **en español**.
- Los links de OnSign, la URL del Apps Script y las claves son constantes al inicio del `<script>` en `index.html`. `APPS_SCRIPT_WRITE_KEY` (index.html) y `WRITE_KEY` (Code.js) deben coincidir siempre.
- El sync a Slides es genérico: copia cada celda con contenido al shape cuyo Alt-Text sea `R{fila}C{columna}`. Agregar una columna al Sheet no requiere tocar el código del sync, solo crear el shape tageado en la Slide.

## Trampas conocidas (ya nos mordieron)

- **Sheets autoconvierte texto que "parece" fecha**: valores como `10-11` (Rango HR) o `3/08/2026` se guardan como Date y se leen mal. Por eso `writeData_()` fuerza `setNumberFormat("@")` en las columnas de texto antes de escribir. Si se agregan campos de texto nuevos, incluirlos en ese rango.
- **Apps Script no puede llamar URLs externas sin el scope `script.external_request`**, y en esta cuenta de Workspace ese permiso no se pudo autorizar. Por eso todo el sync vive dentro del mismo script (nada de `UrlFetchApp` a otro Apps Script).
- Los `fetch` a OnSign van con `mode: 'no-cors'` — no se puede leer la respuesta, se asume éxito optimista.
- La primera ejecución del Apps Script en una cuenta nueva **requiere autorización manual** en script.google.com (no se puede automatizar).

## Al probar

Este proyecto escribe en **recursos de producción reales** (el Sheet y la Slide que se ven en la planta). Si se prueba un guardado con valores ficticios, **restaurar los valores reales inmediatamente después** — y tener presente que cada guardado deja una fila en la pestaña `Log`.

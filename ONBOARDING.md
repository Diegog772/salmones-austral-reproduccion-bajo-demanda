# Salmones Austral — Activación de Campañas OnSign TV

Herramienta interna para que los encargados de área activen/detengan campañas de OnSign TV bajo demanda, y para que Filete reporte sus resultados de turno (que terminan mostrándose en una pantalla física vía una presentación de Google Slides).

## Acceso a la página

- **URL pública**: https://plantasps.getsystem.io (dominio propio) — fallback: https://diegog772.github.io/salmones-austral-reproduccion-bajo-demanda/
- **Clave de acceso a la página** (gate simple, no es seguridad real — está en el código fuente): `sa2026`
- Sesión se guarda en `sessionStorage`; hay botón "Cerrar sesión".

## Repositorio

- **GitHub**: https://github.com/Diegog772/salmones-austral-reproduccion-bajo-demanda (rama `main`, publicada directo como GitHub Pages)
- Deploy es automático: cualquier push a `main` se refleja en el sitio en 1-2 minutos.
- **Pablo necesita ser agregado como colaborador** en este repo (Settings → Collaborators) para poder pushear.

## Estructura del proyecto

```
index.html          # toda la página (HTML+CSS+JS en un solo archivo, sin build)
assets/              # logo y tipografía (Overpass)
CNAME                # dominio propio para GitHub Pages
appscript_v2/         # backend ACTUAL en uso (Code.js + appsscript.json)
appscript/            # backend viejo, YA NO SE USA (retirado, se deja como referencia histórica)
```

## Módulo 1: Activación de Ejercicios

Un botón Activar/Detener por área (Filete, Empaque, Porcionado, Lavado), cada uno pega directo a un link de OnSign TV (`/play/...` o `/stop/...`) con `fetch(..., {mode:'no-cors'})`. Los links están hardcodeados como constantes en el `<script>` de `index.html`. Si OnSign genera links nuevos, solo hay que reemplazar esas constantes.

## Módulo 2: Resultados por Área (Filete)

Formulario con **Línea 1** y **Línea 2**, cada una con 6 campos: Piezas, Rango HR, Trim, Calibre, Cliente, Acumulado. Al guardar:

1. La página llama al **backend de Apps Script** (`appscript_v2/Code.js`), que escribe esos valores en un Google Sheet propio.
2. Ese mismo script, en el mismo request, sincroniza los valores hacia una presentación de **Google Slides** (buscando en la Slide los cuadros de texto cuyo *Alt-Text/Título* sea `R{fila}C{columna}` y reemplazando su texto).
3. La página dispara el link de `play` de OnSign de esa composición, para forzar que la pantalla se refresque ya (sin esperar el ciclo de refresh de OnSign).

### Recursos de Google (Drive) — Pablo necesita acceso de Editor a estos 3

| Recurso | ID | Notas |
|---|---|---|
| Google Sheet (datos) | `1hkzQJJnTtBi_3k9LhlieLu8msHnoorS3Ikf5RsI8vgo` | Pestaña `Hoja 1` = valores actuales (B2:G3, fila 2=L1, fila 3=L2, columnas B..G = Piezas/Rango HR/Trim/Calibre/Cliente/Acumulado; B5 = fecha, fórmula `=HOY()`). Pestaña `Log` = historial de cada guardado. |
| Google Slides (pantalla) | `1SUnpb0vz5XmA5QDpOX2D5dU3UdPKa9CoE9KqP1ez8wo` | Copia propia (de la que hizo Pablo originalmente). OnSign reproduce esta Slide directo. |
| Apps Script (`PSP_Filete_Unificado`) | `1ff296hN7OpW8c1LIGV6MOw-jYNhi59HqcXamcmpZVCuojg6Mi6ocMLM8` | Deployment activo: `AKfycbzwEuz_J3PMHXpj0HLjddFao8WxGAZGb8DSrWkIxDAjBj0ojmdKMFUBvQuieOsGX2li` |

### Cómo editar y desplegar el backend (Apps Script)

Se usa `clasp` (CLI de Google) para no depender del editor web:

```bash
cd appscript_v2
clasp login          # una vez, si no está ya autenticado en esta máquina
clasp pull           # trae el código actual (por si se editó desde el navegador)
# ... editar Code.js ...
clasp push --force
clasp deploy --deploymentId AKfycbzwEuz_J3PMHXpj0HLjddFao8WxGAZGb8DSrWkIxDAjBj0ojmdKMFUBvQuieOsGX2li --description "descripción del cambio"
```

Importante: usar siempre el **mismo `--deploymentId`** al redeployar, así la URL (`APPS_SCRIPT_URL` en `index.html`) no cambia.

**Primera vez en una cuenta/máquina nueva**: la primera ejecución de cualquier función necesita autorización manual (no se puede hacer por API) — abrir el proyecto en script.google.com, elegir una función en el dropdown (ej. `doGet`), Ejecutar, y aceptar los permisos (Avanzado → Ir a [proyecto] → Permitir).

### Endpoint del backend (`APPS_SCRIPT_URL` en index.html)

`https://script.google.com/macros/s/AKfycbzwEuz_J3PMHXpj0HLjddFao8WxGAZGb8DSrWkIxDAjBj0ojmdKMFUBvQuieOsGX2li/exec`

- `?action=read` → lee valores actuales (público, sin key)
- `?action=update&key=sa-resultados-2026&l1_piezas=..&l1_rangoHr=..&l1_trim=..&l1_calibre=..&l1_cliente=..&l1_acumulado=..&l2_...` → guarda y sincroniza la Slide
- `?action=setup&key=sa-resultados-2026` → reescribe encabezados de `Hoja 1` y `Log` (usar si se agregan/renombran columnas)

`sa-resultados-2026` es la clave de escritura — está hardcodeada en `index.html` (`APPS_SCRIPT_WRITE_KEY`) y en `Code.js` (`WRITE_KEY`), deben coincidir siempre.

## Pendiente / próximos pasos conocidos

- **En la Slide todavía faltan los cuadros de texto tageados** para los 5 campos nuevos (Rango HR, Trim, Calibre, Cliente, Acumulado) de cada línea — hoy solo existen los de Piezas y Fecha. Faltan crear (tageados como Alt-Text): `R2C3`..`R2C7` (Línea 1) y `R3C3`..`R3C7` (Línea 2). El sync ya los toma automáticamente en cuanto existan, sin tocar código.
- La pestaña `Log` acumuló varias filas de pruebas de desarrollo (valores como 8080/9090, 4141/5151) — pendiente decidir si se limpian o se dejan como parte del historial.
- Hoy solo está armado el módulo de resultados para **Filete**. Empaque/Porcionado/Lavado solo tienen el módulo de Activación.

## Seguridad (nivel aceptado para esta herramienta interna)

La clave de acceso a la página y la clave de escritura del backend viven en el código fuente (visibles para cualquiera que inspeccione la página) — es una barrera contra accesos casuales, no seguridad real. Aceptado dado el bajo riesgo de la operación. Si en algún momento se necesita seguridad real, habría que mover la lógica a un backend con autenticación de verdad.

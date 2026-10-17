# Seguimiento Obra Autovías — Bacheo con fresado

Panel de costo y utilidad diaria. Frontend en GitHub Pages + backend Apps Script/Google Sheets. Acceso por PIN o desde la app INGECO.

## Qué hace

- **Carga diaria** de toneladas colocadas (y viajes de mezcla si difieren del auto-cálculo).
- **Combustible real** leído de `ENTREGAS DE COMBUSTIBLE - LEANDRO CASARES`, filtrando por la obra. Días sin carga usan el estimado y quedan marcados.
- **Recambios de puntas** para calibrar la vida útil real.
- **Panel**: costo/t del día y promedio de obra, utilidad del día / promedio / acumulada, punto de equilibrio, gráficos de evolución, estructura del costo.
- **Costos editables desde la app** (pestaña Costos): parámetros de CONFIG y lista de EQUIPOS. La planilla es solo almacenamiento; nadie necesita abrirla.
- **Combustible por día con prioridad**: monto manual cargado en el día → planilla de Leandro → estimado.

## Arquitectura

- **Frontend**: `docs/` servido por GitHub Pages → https://marcoskatz-cmd.github.io/obra-autovias/ (PWA instalable, logo INGECO).
- **Backend**: Apps Script vinculado a la planilla, expuesto como API JSON por `doPost` (body `text/plain` `{ accion, pin | sesion, nombre, payload }`). Acciones: `login`, `datos`, `guardarDia` (acepta `comb` en $ para el día), `guardarPuntas`, `borrarDia`, `config`, `guardarConfig`, `obras` (diagnóstico de nombres de obra en la planilla de Leandro).
- **Acceso**: PIN de obra (fila `PIN` de CONFIG, hoy 1234) + nombre de quien carga, guardado en el celular. Desde la app INGECO entra sin PIN (guard.js, módulo `AUTOVIAS`, validado contra app-ingeco). El `/exec` del script solo redirige a Pages.

## Instalación (una sola vez)

1. `clasp login` con la cuenta de INGECO. `clasp create --type sheets --title "SEGUIMIENTO OBRA AUTOVIAS" --rootDir .` crea planilla + script (ojo: pisa `appsscript.json`, restaurarlo con git). `clasp push -f`.
2. Ejecutar **`setup`** una vez desde el editor (crea CONFIG, EQUIPOS, PRODUCCION, PUNTAS). `clasp run` no funciona sin proyecto GCP estándar.
3. Editar `CONFIG`: `PIN`, `OBRA_COMBUSTIBLE` (texto que usa Leandro en OBRA GENERAL o LUGAR ENTREGA; el filtro es *contiene*, sin distinguir mayúsculas ni tildes), precios.
4. `clasp deploy` como aplicación web: ejecutar como **Yo**, acceso **Cualquier usuario (anónimo)** — el control real es el PIN. Pegar la URL `/exec` en `API` dentro de `docs/index.html`.
5. GitHub Pages: Settings → Pages → rama `master`, carpeta `/docs`.
6. En el celular: abrir la URL de Pages en Chrome → menú ⋮ → "Instalar app" (iPhone: Safari → Compartir → "Agregar a inicio").

## Actualizar

- Backend: `clasp push -f && clasp deploy -i <deploymentId> -d "cambio"` (ver sección Deploy).
- Frontend: commit en `master` y `git push`; Pages publica solo en ~1 min (caché HTTP de 10 min).
- Nunca pegar código en el editor online.

## Modelo de costo (por día cargado)

```
fijos      = (Σ alquileres USD/mes ÷ DIAS_MES) + MO_USD + SENAL_USD
             cada equipo suma solo desde su fecha DESDE (col. F de EQUIPOS; vacía = OBRA_INICIO).
             Autovías: todo menos la fresadora cuenta desde el 2026-10-13; los días previos van sin costo.
             (La columna DESDE se crea sola la primera vez que corre el backend nuevo sobre una planilla vieja.)
combustible= manual del día | real de Leandro | COMB_ESTIMADO_ARS   (÷ (1+IVA si COMB_CON_IVA) ÷ TC)
flete      = viajes × FLETE_ARS ÷ TC        viajes = ceil(t ÷ CAP_CAMION) salvo carga manual
puntas     = t × (PUNTAS_JUEGO_ARS ÷ 1,21 ÷ TC ÷ vida) × SOBREFRESADO
             vida = PUNTAS_VIDA_T hasta tener ≥2 recambios; después t fresadas reales entre recambios
costo      = fijos + combustible + flete + puntas
utilidad   = t × PRECIO_T − costo          (sin gastos generales, sin IVA)
```

## Auto-diagnóstico

- Si la planilla de Leandro cambia de nombre de columna (FECHA, CANTIDAD (L), PRECIO TOTAL, OBRA GENERAL), el panel avisa cuál falta en vez de fallar en silencio.
- Si el filtro de obra no encuentra registros, avisa en el panel.
- Si el precio del combustible de Leandro incluye IVA (precio de surtidor), dejar `COMB_CON_IVA = 1`.

## Archivos

- `Codigo.gs` — backend: lectura de config, combustible, cálculo, guardado.
- `docs/` — frontend mobile-first (index.html con Chart.js desde cdnjs, manifest.json, sw.js, logo e íconos) publicado en GitHub Pages.
- `appsscript.json` — manifiesto (zona horaria, web app).

## Deploy

- **scriptId**: `1TJmbbw5nDvbi57LmSYyKfeOzV5HNT62OaotdAYdnGv_YsJzfctO1_CTJ`
- **Planilla**: https://docs.google.com/spreadsheets/d/1oXBh5DDSIeCPsNxv9agst84b5KSOTdTpb3vNDIxKOl4/edit
- **Editor**: https://script.google.com/d/1TJmbbw5nDvbi57LmSYyKfeOzV5HNT62OaotdAYdnGv_YsJzfctO1_CTJ/edit
- **deploymentId**: `AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg`
- **URL de la app (Pages)**: https://marcoskatz-cmd.github.io/obra-autovias/
- **URL de la API (/exec)**: https://script.google.com/macros/s/AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg/exec
- **Repo**: https://github.com/marcoskatz-cmd/obra-autovias (público: GitHub Pages lo requiere en cuenta Free; no contiene secretos)

Actualizar (desde esta carpeta, con `.clasp.json` presente):

```bash
clasp push -f && clasp deploy -i AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg -d "cambio"
```



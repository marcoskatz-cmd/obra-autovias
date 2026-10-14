# Seguimiento Obra Autovías — Bacheo con fresado

Panel de costo y utilidad diaria. Apps Script + Google Sheets. Acceso restringido por mail.

## Qué hace

- **Carga diaria** de toneladas colocadas (y viajes de mezcla si difieren del auto-cálculo).
- **Combustible real** leído de `ENTREGAS DE COMBUSTIBLE - LEANDRO CASARES`, filtrando por la obra. Días sin carga usan el estimado y quedan marcados.
- **Recambios de puntas** para calibrar la vida útil real.
- **Panel**: costo/t del día y promedio de obra, utilidad del día / promedio / acumulada, punto de equilibrio, gráficos de evolución, estructura del costo.
- Parámetros (TC, alquileres, precio pactado, flete) viven en la planilla: pestañas `CONFIG` y `EQUIPOS`.

## Instalación (una sola vez)

1. Crear una Google Sheet nueva en el Shared Drive de INGECO: **"SEGUIMIENTO OBRA AUTOVIAS"**. Compartirla solo con Marcos y Agustín.
2. Extensiones → Apps Script. Copiar el `scriptId` (Configuración del proyecto).
3. En esta carpeta:
   ```bash
   npm i -g @google/clasp
   clasp login
   cp .clasp.json.example .clasp.json   # pegar el scriptId
   clasp push
   ```
4. En el editor de Apps Script ejecutar **`setup`** una vez (crea CONFIG, EQUIPOS, PRODUCCION, PUNTAS). Aceptar permisos: lee Sheets (la planilla de Leandro) y usa Session.
5. Editar `CONFIG`:
   - `ACCESO`: mails con permiso, separados por coma. Verificar el mail exacto de Agustín.
   - `OBRA_COMBUSTIBLE`: el texto con que Leandro va a identificar esta obra en la columna **OBRA GENERAL** (o en LUGAR ENTREGA). **Esto hay que acordarlo con Leandro antes de que empiece a cargar** — si él escribe "Autovias" sin tilde y acá dice "Autovías", no matchea. El filtro es *contiene*, sin distinguir mayúsculas, pero sí tildes.
6. Implementar → Nueva implementación → Aplicación web:
   - Ejecutar como: **Yo** (Marcos). La app lee la planilla de Leandro con tu permiso; Agustín no necesita acceso a esa planilla.
   - Acceso: **Cualquier usuario de grupoingeco.com.ar**. El filtro real es la lista `ACCESO`: cualquier otra cuenta del dominio ve "Sin acceso".
7. Pasarle la URL a Agustín. En el celular: abrir en Chrome → menú → "Agregar a pantalla de inicio".

## Actualizar

`clasp push` y luego Implementar → Administrar implementaciones → editar → Nueva versión. Nunca pegar código en el editor online.

## Modelo de costo (por día cargado)

```
fijos      = (Σ alquileres USD/mes ÷ DIAS_MES) + MO_USD + SENAL_USD      [× 0 si jornada = parada]
combustible= real de Leandro ÷ (1+IVA si COMB_CON_IVA) ÷ TC   |   COMB_ESTIMADO_ARS ÷ TC si no hay carga
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
- `index.html` — frontend mobile-first (Chart.js desde cdnjs).
- `appsscript.json` — manifiesto (zona horaria, web app).

## Deploy

- **scriptId**: `1TJmbbw5nDvbi57LmSYyKfeOzV5HNT62OaotdAYdnGv_YsJzfctO1_CTJ`
- **Planilla**: https://docs.google.com/spreadsheets/d/1oXBh5DDSIeCPsNxv9agst84b5KSOTdTpb3vNDIxKOl4/edit
- **Editor**: https://script.google.com/d/1TJmbbw5nDvbi57LmSYyKfeOzV5HNT62OaotdAYdnGv_YsJzfctO1_CTJ/edit
- **deploymentId**: `AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg`
- **URL de la app**: https://script.google.com/macros/s/AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg/exec

Actualizar (desde esta carpeta, con `.clasp.json` presente):

```bash
clasp push -f && clasp deploy -i AKfycbwWwSGkx8kvPjYXGQvMTAc_7jYKqcOV47mDeWXOHm_FHOZcZ9v0RIXr1TYEo4JEPOqywg -d "cambio"
```

Nota: `clasp run` no funciona hasta asociar un proyecto GCP estándar al script; `setup` se ejecutó desde el editor.

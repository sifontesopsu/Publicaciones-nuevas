# App Gestión de Publicaciones Pendientes - Aurora V5 Cloud

Versión preparada para:

```text
GitHub + Streamlit Cloud + Google Sheets/AppScript
```

## Arquitectura

```text
GitHub
├── app.py
├── requirements.txt
├── appscript.gs
└── data/
    ├── maestro_sku_ean.xlsx
    ├── publicaciones_mercado_libre.xlsx
    └── packs.xlsx

Streamlit Cloud
└── Interfaz usada por 3 personas

Google Sheets + Apps Script
└── Base central:
    ├── estado_actual
    ├── inventario_actual
    ├── inventario_cargas
    ├── historial_cambios
    └── sync_log
```

## Regla clave

```text
Si un SKU aparece en LibroInventario.xlsx y no está en maestro_sku_ean.xlsx,
la app lo agrega como PRODUCTO NUEVO.
```

## Estados

```text
SIN STOCK
LLEGÓ STOCK
PENDIENTE VERIFICAR FÍSICO
STOCK FÍSICO CONFIRMADO
EN PROCESO DE PUBLICACIÓN
PUBLICADO
NO ENCONTRADO FÍSICO
NO PUBLICABLE
FALTA FOTO
FALTA INFORMACIÓN
CUBIERTO POR PACK
CUBIERTO POR UNIDAD
PRODUCTO NUEVO CON STOCK
PRODUCTO NUEVO SIN STOCK
PRODUCTO NUEVO PUBLICADO
```

## Paso 1: preparar Google Sheets

1. Crea un Google Sheet nuevo.
2. Copia el ID del Sheet desde la URL.

Ejemplo de URL:

```text
https://docs.google.com/spreadsheets/d/ID_DEL_SHEET/edit
```

El ID es lo que está entre `/d/` y `/edit`.

## Paso 2: instalar Apps Script

1. En el Google Sheet, abre:
   - Extensiones > Apps Script
2. Borra el código existente.
3. Pega el contenido de `appscript.gs`.
4. Guarda el proyecto.

## Paso 3: configurar Script Properties

En Apps Script:

```text
Project Settings > Script Properties
```

Agrega:

```text
SPREADSHEET_ID = ID_DE_TU_GOOGLE_SHEET
SECRET_TOKEN = aurora_publicaciones_2026
```

Puedes cambiar el token, pero debe ser igual al de Streamlit.

## Paso 4: publicar Apps Script

En Apps Script:

```text
Deploy > New deployment > Web app
```

Configura:

```text
Execute as: Me
Who has access: Anyone with the link
```

Copia la URL del Web App.

## Paso 5: subir proyecto a GitHub

Sube estos archivos a tu repo:

```text
app.py
requirements.txt
appscript.gs
README_CLOUD.md
data/maestro_sku_ean.xlsx
data/publicaciones_mercado_libre.xlsx
data/packs.xlsx
```

## Paso 6: crear app en Streamlit Cloud

1. Entra a Streamlit Cloud.
2. New app.
3. Selecciona tu repo.
4. Branch: main.
5. Main file path:

```text
app.py
```

## Paso 7: configurar Secrets en Streamlit Cloud

En la configuración de la app, pega:

```toml
APP_SCRIPT_URL = "https://script.google.com/macros/s/XXXX/exec"
APP_SCRIPT_TOKEN = "aurora_publicaciones_2026"
```

## Uso diario

1. Un usuario entra a la app.
2. Escribe su nombre en "Responsable carga inventario".
3. Sube `LibroInventario.xlsx`.
4. Presiona "Procesar y guardar inventario".
5. La app guarda el inventario central en Google Sheets.
6. Los 3 usuarios ven la misma cola.
7. Cada usuario marca estados desde "Cola Marketing".
8. Los cambios quedan en:
   - `estado_actual`
   - `historial_cambios`

## Importante para 3 usuarios

No se instala la app en 3 computadoras.

Los 3 usuarios entran al mismo link de Streamlit Cloud.

## Archivos que se actualizan

Desde GitHub:

```text
maestro_sku_ean.xlsx
publicaciones_mercado_libre.xlsx
packs.xlsx
```

Desde Streamlit:

```text
LibroInventario.xlsx
```

Desde la app:

```text
estado_actual
inventario_actual
historial_cambios
```

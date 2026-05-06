# Dashboard KPIs Panier

Dashboard web para visualizar los KPIs del proyecto. Lee los datos desde **Google Sheets** en una carpeta compartida de Drive.

## Arquitectura

- `build-data.js` se conecta a Drive con un service account, lista todos los Google Sheets de la carpeta configurada y genera `dashboard/data.json` + `dashboard/data.js` con la estructura normalizada.
- El dashboard (`dashboard/index.html`) es 100% estático y consume `data.js`. Se abre con doble clic o desde un servidor estático.

```
.
├── build-data.js          ← ingesta desde Google Drive
├── service-account.json   ← credenciales (NO se commitea)
├── .env                   ← config (NO se commitea)
└── dashboard/
    ├── index.html
    ├── app.js
    ├── styles.css
    ├── data.js            ← autogenerado
    └── data.json          ← autogenerado
```

## Setup (primera vez)

### 1. Crear el service account en Google Cloud

1. Entrá a [console.cloud.google.com](https://console.cloud.google.com) y creá (o seleccioná) un proyecto.
2. Habilitá las APIs:
   - **Google Drive API**
   - **Google Sheets API**
3. En `IAM y administración` → `Cuentas de servicio` → `Crear cuenta de servicio`.
4. Dale un nombre (ej: `panier-dashboard-reader`). No hace falta asignar roles.
5. Una vez creada, entrá a la cuenta → pestaña `Claves` → `Agregar clave` → `Crear clave` → formato **JSON**. Descargá el archivo.
6. Renombralo a `service-account.json` y copialo a la raíz del proyecto (`Panier/`).
7. Copiá el email de la cuenta de servicio (algo como `panier-dashboard-reader@<proyecto>.iam.gserviceaccount.com`).

### 2. Compartir la carpeta de Drive

1. Abrí la carpeta de Drive con los Sheets.
2. Click en `Compartir` → pegá el email del service account → rol **Lector** → `Listo`.

### 3. Configurar el proyecto

1. Instalá dependencias (una sola vez):
   ```bash
   npm install
   ```
2. Copiá el ejemplo de configuración:
   ```bash
   cp .env.example .env
   ```
3. Editá `.env` y poné el ID de tu carpeta (la parte final de la URL de Drive).

## Uso diario

```bash
node build-data.js
```

Regenera los datos leyendo todos los Google Sheets de la carpeta. Después recargá el dashboard en el navegador.

## Hojas reconocidas

El parser identifica pestañas de cada Sheet por nombre (no importa el orden):

| Nombre contiene | Se interpreta como |
|---|---|
| `Mensual General` | KPIs mensuales globales |
| `x cliente` / `por cliente` | Detalle por cliente |
| `tipo de cliente` | Clasificación por tipo |
| `Unidades` + `Kg` | Detalle por producto |
| `Resumen` | Resumen anual |

Los encabezados de columnas se detectan automáticamente: `Uni Ene`, `Ped Feb`, `Fact Mar`, `Marzo Uni`, `Enero Kg`, etc.

El año se extrae del nombre del Sheet (ej: `KPIs Panier 2026`).

## Funciones del dashboard

- **Resumen**: KPIs del año con comparación contra el año anterior y gráficos mensuales.
- **Mensual**: Tabla con el registro mensual general.
- **Clientes**: Top N por facturación / unidades / pedidos, con búsqueda y filtro por tipo.
- **Tipos de cliente**: Donut y tabla por categoría.
- **Productos**: Ranking por unidades o kg, con filtro por categoría.

## Troubleshooting

- **`No se encontraron Google Sheets en esa carpeta`** → verificá que compartiste la carpeta con el email del service account.
- **`The caller does not have permission`** → el service account no tiene acceso a esa carpeta o a algún Sheet dentro.
- **Faltan hojas en el dashboard** → el nombre de la pestaña en el Sheet no matchea ninguna regla de la tabla de arriba. Renombrala o ajustá `classifySheet` en `build-data.js`.

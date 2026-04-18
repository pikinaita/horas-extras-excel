# Horas Extraordinarias — Dropbox + Excel

[![Deploy to Netlify](https://www.netlify.com/img/deploy/button.svg)](https://app.netlify.com/start/deploy?repository=https://github.com/pikinaita/horas-extras-excel)

App web para registrar horas extraordinarias directamente en un archivo Excel alojado en Dropbox. **Sin instalaciones, sin backend** — todo funciona en el navegador (PC, móvil, iPhone).

## Funcionalidades

- Autenticación OAuth 2.0 con Dropbox (token guardado en `localStorage`)
- Descarga y manipula el `.xlsx` original directamente desde Dropbox
- Detecta automáticamente el **mes y día actual**
- Permite registrar turno (Mañana / Tarde / Noche) y tipo de cobertura
- Escribe directamente en la celda del día correspondiente
- Sube el archivo modificado a Dropbox (modo overwrite)
- Diseño optimizado para móvil con colores naranja/negro
- Compatible con iPhone (añadir a pantalla de inicio como PWA)

## Estructura del Excel esperada

El archivo debe contener **una hoja por mes** (ej. `ENERO`, `FEBRERO`, etc.) con:

- **Columna B (2):** Días del mes (números del 1 al 31)
- **Columnas C, D, E (3, 4, 5):** Turnos (Mañana, Tarde, Noche)
- **Columna P (16):** Cobertura (ej. "Sustitución Operador")

La app busca la fila correspondiente al día seleccionado y marca con `X` el turno elegido.

## Ejecutar desde el iPhone (recomendado)

### Opción A — Desplegar en Netlify (gratis y permanente)

#### 1. Conectar el repositorio a Netlify

1. Ve a [netlify.com](https://netlify.com) y crea una cuenta gratuita
2. **Add new site** → **Import an existing project** → **GitHub**
3. Selecciona el repositorio `horas-extras-excel`
4. Netlify detecta automáticamente `netlify.toml` — solo dale clic a **Deploy site**
5. Te asignarán una URL tipo `https://amazing-name-123.netlify.app`

#### 2. Configurar Dropbox OAuth

Ya tienes un **Client ID** hardcodeado en `index.html` (`gp93qf4f6ozo0oh`). Si es tu app de Dropbox:

1. Ve a [dropbox.com/developers/apps](https://www.dropbox.com/developers/apps)
2. Abre tu app y en **OAuth 2 → Redirect URIs** añade:
   ```
   https://TU-SITIO.netlify.app
   ```
3. Guarda cambios

> Si no tienes app en Dropbox, crea una nueva con **Scoped access** → **Full Dropbox** y usa el nuevo Client ID editando `index.html` línea 11.

#### 3. Añadir al iPhone como app

1. Abre **Safari** en el iPhone y ve a `https://TU-SITIO.netlify.app`
2. Pulsa el botón **Compartir** (icono ⎋)
3. Selecciona **"Añadir a pantalla de inicio"**
4. Ponle un nombre (ej. "Horas Extra") y toca **Añadir**

La app aparecerá en tu pantalla de inicio como cualquier otra app.

#### 4. Flujo de uso

1. Abre la app desde el icono
2. Toca **Conectar con Dropbox** → autorizas
3. Vuelves a la app automáticamente (token guardado)
4. Introduce la ruta del archivo Excel (ej. `/Sg/Partes2026.xlsx`)
5. Pulsa **Cargar y Vincular**
6. Selecciona mes, día, turno y cobertura
7. **Guardar en Archivo Original** → el Excel se actualiza en Dropbox instantáneamente

---

### Opción B — Prueba rápida en red local (WiFi)

Si tu iPhone y tu PC están en la misma red WiFi:

```bash
# En el PC, abre una terminal en la carpeta del repo
python -m http.server 8000
# O con Node.js
npx serve .
```

Luego en el iPhone, abre Safari y ve a `http://IP-DE-TU-PC:8000`.

> **Limitación:** Solo funciona mientras el PC esté encendido y en la misma red. No es permanente.

---

## Estructura del proyecto

```
horas-extras-excel/
├── index.html          # App completa (HTML + CSS + JS)
├── netlify.toml        # Configuración de Netlify
├── main.py             # Versión Python (obsoleta para web)
├── requirements.txt    # Dependencias Python (no usadas en web)
└── README.md
```

## Tecnologías

- HTML5 + JavaScript puro (sin frameworks)
- [Dropbox SDK](https://www.dropbox.com/developers/documentation/http/overview) — OAuth + API de ficheros
- [ExcelJS](https://github.com/exceljs/exceljs) — Lectura y escritura de `.xlsx` en el navegador
- PWA optimizada para iPhone con meta tags Apple

## Notas

- El **token de Dropbox no caduca** con OAuth Implicit Grant (se guarda en localStorage)
- La app modifica **directamente** el archivo original en Dropbox (modo `overwrite`)
- Funciona **sin conexión** una vez cargado el HTML (pero necesita internet para leer/escribir Dropbox)

## Licencia

MIT

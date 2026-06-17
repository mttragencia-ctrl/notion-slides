# 🧁 Notion → Slides

App web que convierte un calendario de contenidos de **Notion** en una
presentación **.pptx** lista para abrir en **Google Slides**, con el diseño
cálido de panadería de La Porteña.

Pensada para la agencia: una sola persona la configura una vez y después todas
entran con un link y una contraseña compartida.

---

## ¿Cómo se usa? (uso diario)

1. Entrá al link de la app.
2. Poné la **contraseña de la agencia**.
3. Pegá el **link de la base de Notion** del mes (la vista de tabla o calendario).
4. Clic en **Generar presentación** → se descarga un archivo `.pptx`.
5. Subí ese archivo a [Google Drive](https://drive.google.com) → clic derecho →
   **Abrir con → Presentaciones de Google**. ¡Listo!

---

## Puesta en marcha (una sola vez)

### Paso 1 · Crear la integración de Notion (para que la app pueda leer)

1. Entrá a <https://www.notion.so/my-integrations> → **New integration**.
2. Nombre: `Slides Agencia`. Tipo: **Internal**. Asociala a tu workspace.
3. Copiá el **Internal Integration Secret** (empieza con `ntn_…`). Es el `NOTION_TOKEN`.
4. **Compartí tus bases con la integración**: abrí en Notion la página que
   contiene los calendarios (por ej. *La Porteña - MKT*) → botón `•••`
   (arriba a la derecha) → **Conexiones / Connections** → agregá `Slides Agencia`.
   Las sub-páginas y bases heredan el acceso.

### Paso 2 · Subir el código a GitHub

> Si Claude ya lo hizo por vos, salteá este paso.

- Creá un repositorio nuevo en <https://github.com> y subí esta carpeta.

### Paso 3 · Publicar en Render (gratis)

1. Creá una cuenta en <https://render.com> (podés entrar con GitHub).
2. **New → Web Service** → conectá el repositorio.
3. Render detecta el archivo `render.yaml` automáticamente. Si pide datos a mano:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
4. En **Environment** agregá las variables:
   - `NOTION_TOKEN` → el secreto del Paso 1.
   - `APP_PASSWORD` → la contraseña que van a compartir en la agencia.
   - `SECRET_KEY` → cualquier texto largo y aleatorio (o dejá que Render lo genere).
5. **Create Web Service**. En 1-2 minutos tenés tu link:
   `https://notion-slides-xxxx.onrender.com`

> ℹ️ En el plan gratuito, si nadie la usa por un rato, la primera carga puede
> tardar ~50 segundos en "despertar". Es normal.

---

## Correr en tu computadora (opcional, para probar)

```bash
pip install -r requirements.txt
cp .env.example .env        # completá NOTION_TOKEN y APP_PASSWORD
# cargá las variables y arrancá:
export $(grep -v '^#' .env | xargs)
python app.py
# abrí http://localhost:5000
```

Para probar solo el diseño (sin Notion):

```bash
python test_generate.py     # genera test_output.pptx
```

---

## Estructura

```
app.py                 Servidor Flask (login + generación + descarga)
lib/parse_url.py       Extrae el ID de la base desde el link de Notion
lib/notion_fetch.py    Lee la base, mapea propiedades y contenido
lib/slides.py          Arma el .pptx con el diseño de panadería
templates/index.html   Interfaz
static/style.css       Estilos
render.yaml            Configuración de despliegue en Render
```

## Notas

- Las bases de cada mes deben usar el **mismo template** que "Planificación
  Julio" (mismas propiedades: Nombre, Fecha, Formato, Pilar, Tipo, Objetivo,
  Estado). El mapeo es tolerante, pero nombres muy distintos requieren ajuste.
- La "Red social" se deduce del Formato (Historias → Stories; Carrusel/Posteo →
  Feed; Reel → Instagram).
- Las páginas sin guión ni copy aparecen como slide **"Contenido pendiente"**.

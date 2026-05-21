# Memoria Naval — La Real Armada Española del siglo XVIII

Web pública navegable construida automáticamente a partir de cinco hojas de
cálculo Excel. Cada vez que actualices los Excels y subas los cambios, la
web se reconstruye sola en unos minutos.

---

## Cómo funciona, en una frase

Tú mantienes los datos en cinco Excels en este repositorio. Un robot
(GitHub Actions) los lee con un script (`extract.py`), genera un único
archivo HTML autosuficiente (`index.html`) y lo publica en GitHub Pages.
La web no tiene servidor ni base de datos: es un solo archivo que se
abre directo en el navegador.

```
   5 Excels  ─►  extract.py  ─►  index.html  ─►  GitHub Pages
   (tú)         (automático)    (automático)    (visible al público)
```

---

## Primera vez: poner la web en marcha

Hazlo una sola vez. Después solo tendrás que editar Excels.

### 1. Crear una cuenta en GitHub

Si no la tienes: ve a <https://github.com/signup> y crea una cuenta
gratuita. Anota tu nombre de usuario (te hará falta más adelante; lo
llamaremos `TUUSUARIO`).

### 2. Crear un repositorio nuevo

1. Una vez con sesión iniciada, pulsa el botón verde **"New"** en la
   página de inicio (o ve a <https://github.com/new>).
2. **Repository name**: ponle el nombre que quieras (por ejemplo
   `armada-xviii`). Será parte de la URL pública de la web.
3. Déjalo en **Public** (público). Eso permite que GitHub Pages publique
   la web gratis.
4. Marca **"Add a README file"** (no importa lo que diga; lo
   sobreescribiremos).
5. Pulsa **Create repository**.

### 3. Descargar GitHub Desktop

GitHub Desktop es una aplicación gráfica que te evita usar terminal.

1. Descarga desde <https://desktop.github.com> e instala.
2. Ábrela e inicia sesión con tu cuenta de GitHub.
3. Menú **File → Clone repository → GitHub.com**. Selecciona el
   repositorio que acabas de crear y elige una carpeta local donde
   guardarlo (por ejemplo `Documentos/armada-xviii`). Pulsa **Clone**.

Ahora tienes una carpeta en tu ordenador que es el "espejo" del
repositorio en GitHub.

### 4. Copiar los archivos del paquete

Copia todos los archivos de este paquete (lo que te entregué) dentro de
esa carpeta. Debe quedarte así:

```
armada-xviii/
├── .github/
│   └── workflows/
│       └── build.yml
├── BUQUE_CLAUDE.xlsx
├── PERSONAL_CLAUDE.xlsx
├── LUGAR_CLAUDE.xlsx
├── GRADO_CLAUDE.xlsx
├── lat_claude.xlsx
├── extract.py
├── template.html
├── requirements.txt
├── .gitignore
└── README.md
```

> El README.md (este archivo) sobrescribirá el que GitHub creó por defecto.

### 5. Subir los archivos a GitHub

Vuelve a **GitHub Desktop**:

1. Verás una lista de archivos en azul (los nuevos que has copiado).
2. Abajo a la izquierda, donde dice **"Summary"**, escribe algo como
   `Versión inicial de la web`.
3. Pulsa **Commit to main**.
4. Arriba a la derecha pulsa **Push origin**.

En unos segundos los archivos estarán en GitHub.

### 6. Activar GitHub Pages

1. En tu navegador, abre el repositorio en GitHub
   (`https://github.com/TUUSUARIO/armada-xviii`).
2. Pulsa la pestaña **Settings** (arriba a la derecha).
3. En el menú izquierdo, pulsa **Pages**.
4. En **"Build and deployment"** → **Source**, elige **GitHub Actions**
   (no "Deploy from a branch"). Esto es importante.

No hace falta guardar nada: GitHub Pages se queda activado en cuanto
seleccionas esa opción.

### 7. Esperar a que se construya la web

1. Vuelve a la pestaña **Actions** del repositorio.
2. Verás un workflow en marcha llamado *"Construir y publicar la web"*.
3. Tarda 2–4 minutos. Cuando termine, verás un check verde.
4. La URL pública de la web aparece dentro del job **publicar**, o en
   **Settings → Pages**. Suele ser:
   `https://TUUSUARIO.github.io/armada-xviii/`

¡Ya está! Esa URL la puedes compartir con quien quieras.

---

## Mantenimiento diario: actualizar los datos

Cuando añadas o corrijas algo en los Excels:

1. **Edita el Excel** en tu ordenador (en la carpeta del repositorio).
   Guárdalo y ciérralo.
2. **Abre GitHub Desktop**. Te aparecerá el Excel modificado en azul.
3. En **Summary** escribe un resumen breve (ej. *"Añado el navío
   San Telmo y dos rutas más"*).
4. Pulsa **Commit to main** y luego **Push origin**.
5. Espera 2–4 minutos: la web se reconstruye automáticamente con los
   nuevos datos.

Si te equivocas, no pasa nada: corriges el Excel, vuelves a hacer
commit y push, y al siguiente build queda arreglado.

---

## Qué hace cada archivo

| Archivo | Para qué sirve | ¿Lo editas tú? |
|---|---|---|
| Los 5 Excels | Tus datos | **Sí, son tu fuente de verdad** |
| `extract.py` | Script que lee los Excels y genera el dataset | No, salvo cambios técnicos |
| `template.html` | Plantilla visual de la web | No, salvo cambios técnicos |
| `.github/workflows/build.yml` | Receta del robot de construcción | No |
| `requirements.txt` | Lista de librerías de Python necesarias | No |
| `.gitignore` | Qué archivos no debe subir GitHub Desktop | No |

Y **se generan automáticamente** (no aparecen en tu carpeta local,
solo en la web publicada):

| Archivo | Qué es |
|---|---|
| `index.html` | La web pública, con todos los datos dentro |
| `data.json` | Los datos en formato JSON (útil para depuración) |
| `incidencias.txt` | Lista de avisos del extractor (fechas mal formadas, IDs huérfanos, etc.) |

---

## Revisar las incidencias

Cada vez que el robot construye la web, escribe un archivo
`incidencias.txt` con los avisos: por ejemplo, *"CARGA: id_ruta='BUQ00028'
no existe en su catálogo (registro id=CARG000033)"*. Esto te ayuda a
encontrar errores tipográficos o referencias rotas.

Para verlo, abre `https://TUUSUARIO.github.io/armada-xviii/incidencias.txt`
en el navegador.

---

## El modelo de datos en pocas palabras

- **Buques** son las naves. Cada buque tiene **rutas** (cada ruta es
  una travesía: parte de un puerto y llega a otro en fechas concretas).
- **Personal** son las personas. Cada una tiene **ascensos**, **destinos
  a bordo** (a qué buque está asignado y en qué fechas) y **destinos
  en tierra**.
- **Cargamentos**, **caudales**, **pasajeros** y **transporte de
  tropas** son lo que viaja en los buques. Si rellenas en el Excel
  los campos `id_lugar_carga` e `id_lugar_descarga` (o `id_lugar_embarque`
  e `id_lugar_desembarque`), la web entiende automáticamente que ese
  cargamento atraviesa varias rutas consecutivas, y marca cada parada
  con un badge:

  - 🟢 *embarca aquí* — sube a bordo en este tramo
  - 🔵 *en tránsito* — ya estaba a bordo, sigue
  - 🟠 *desembarca aquí* — baja a bordo en este tramo
  - 🟤 *trayecto completo* — embarca y desembarca en el mismo tramo

- **Posiciones** (`lat_claude.xlsx`) son posiciones diarias. Si una
  posición no tiene `id_buque` pero sí `id_escuadra`, se aplica a
  todos los buques que en esa fecha pertenecen a la escuadra.

---

## ¿Algo no funciona?

- **El check de Actions sale rojo**: pulsa sobre él para ver qué
  pasó. Casi siempre es un error en los Excels (una columna renombrada,
  una hoja borrada, un Excel corrupto). El log de errores suele decirlo
  con bastante claridad.
- **La web se publicó pero faltan datos**: revisa `incidencias.txt`.
  Probablemente hay IDs que apuntan a algo que no existe.
- **No sé qué columnas debe tener cada hoja**: mira un Excel que sí
  funciona y respeta los nombres exactos de las columnas (sensibles a
  mayúsculas y acentos).

---

## Créditos

Investigación y datos: investigador del proyecto.
Código de extracción y plantilla web: trabajo en colaboración con Claude (Anthropic), 2025-2026.

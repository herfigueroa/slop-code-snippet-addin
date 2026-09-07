# Snippets de Código — Add-in para Word y PowerPoint

Add-in de tipo **task pane** que permite insertar fragmentos de código con
**resaltado de sintaxis** y **fuente monoespaciada**, tanto en **Word** como
en **PowerPoint**, desde el mismo panel.

## ¿Qué hace exactamente?

- Panel lateral con un editor de código, selector de lenguaje, tamaño de
  fuente, **selector de tipografía** (con opción de fuente personalizada),
  **selector de esquema de color** (Dark+, Monokai, Dracula, GitHub Light),
  un toggle para **título del snippet** (desactivado por defecto) y un
  toggle para **fondo/tarjeta** (pensado especialmente para Word).
- **Ajuste especial para PHP**: las variables `$foo`, las etiquetas
  `<?php ?>` y los `namespace/use` reciben colores propios (ver
  `CONFIG.languageOverrides.php` en `config.js`), independientemente del
  esquema de color elegido.
- **PHP + HTML (embebido)**: opción alternativa en el selector de lenguaje
  que, además de resaltar el PHP, también resalta como HTML real (tags,
  atributos, comentarios) todo el marcado que quede fuera de las etiquetas
  `<?php ?>` — el mismo tipo de archivo mixto que trabaja la extensión de
  VS Code [format-html-in-php](https://github.com/RiFi2k/format-html-in-php),
  solo que aquí es para resaltado en lugar de formateo.
- **Pseudocódigo de PSeInt**: opción de lenguaje con una gramática escrita a
  mano (PSeInt no es un lenguaje real de Prism.js) a partir del diccionario
  de palabras clave en español de PSeInt y su sintaxis de Vim. Reconoce
  `Proceso`/`FinProceso`, `Si`/`Entonces`/`Sino Si`/`Sino`/`FinSi`,
  `Mientras`/`Para`/`Repetir` y sus cierres, `Definir ... Como <tipo>`,
  funciones incorporadas (`Abs`, `Longitud`, `ConvertirANumero`, etc.),
  cadenas, comentarios `//` y los operadores propios de PSeInt (`<-`, `←`,
  `Y`/`O`/`No`, `MOD`, etc.).
- **Funciona sin conexión a internet** (salvo la carga inicial de Office.js):
  Prism.js y las gramáticas de los lenguajes más comunes están vendorizadas
  en `vendor/prism/`, así que no dependen de un CDN. Ver
  [Uso sin conexión](#uso-sin-conexión-offline) más abajo.
- Vista previa en vivo con el mismo resaltado que se insertará.
- Al pulsar **Insertar**:
  - **En Word** → inserta una tabla HTML nativa (editable como cualquier
    tabla de Word: puedes seleccionar el texto, cambiar colores a mano,
    copiar/pegar, etc.). El color de cada token se aplica como formato de
    caracter real, no como imagen.
  - **En PowerPoint** → como PowerPoint no soporta insertar texto con
    múltiples colores por caracter vía la API de Office.js, el snippet se
    dibuja en un `<canvas>` (con el mismo resaltado y tipografía) y se
    inserta como imagen PNG en alta resolución. Esto da fidelidad visual
    perfecta y consistente entre plantillas.
- **Editar un snippet ya insertado**: selecciona el bloque insertado (la
  tabla en Word, o la imagen en PowerPoint), ajusta el código/opciones en el
  panel y vuelve a pulsar "Insertar": como Office.js inserta sobre la
  selección actual, reemplaza automáticamente lo seleccionado.

## Estructura del proyecto

```
code-snippet-addin/
├─ manifest.xml                 ← define el add-in para Word y PowerPoint
├─ package.json                 ← scripts para levantar el servidor local
├─ assets/                      ← íconos del add-in
└─ src/taskpane/
   ├─ taskpane.html             ← UI del panel (carga Prism.js desde vendor/)
   ├─ taskpane.css              ← estilos del panel
   ├─ taskpane.js               ← lógica de UI + llamadas a Office.js
   ├─ config.js                 ← 🎨 TODA la personalización vive aquí
   ├─ highlighter.js            ← tokenización con Prism.js → HTML con color inline
   ├─ renderer.js               ← genera el HTML (Word) / imagen canvas (PowerPoint)
   └─ vendor/prism/             ← Prism.js vendorizado (ver "Uso sin conexión")
      ├─ prism-core-bundle.min.js     (core + markup/HTML + css + clike + javascript)
      ├─ prism-markup-templating.min.js
      ├─ prism-php.min.js
      ├─ prism-python.min.js
      ├─ prism-typescript.min.js
      ├─ prism-json.min.js
      ├─ prism-bash.min.js
      ├─ prism-csharp.min.js
      ├─ prism-c.min.js               (también cubre "C / C++" en el selector)
      ├─ prism-java.min.js
      └─ prism-pseint.js              (gramática propia, no oficial de Prism)
```

## Cómo personalizarlo fácilmente

Casi cualquier cambio visual o de comportamiento se hace en
**`src/taskpane/config.js`**, sin tocar el resto del código:

| Quiero cambiar…                              | Edita…                                   |
|------------------------------------------------|-------------------------------------------|
| Las fuentes disponibles en el selector          | `CONFIG.fonts` (agrega `{id, label, family}`) |
| La fuente/tema seleccionados por defecto        | `CONFIG.defaultFontId` / `CONFIG.defaultThemeId` |
| Los esquemas de color disponibles               | `CONFIG.themes` (copia un bloque completo y ajusta colores/`card`/`title`) |
| Ajustes de color específicos de un lenguaje     | `CONFIG.languageOverrides` (ver el ejemplo de `php`) |
| El padding/radio/grosor de borde de la tarjeta  | `CONFIG.card` |
| El peso/tamaño de la barra de título            | `CONFIG.title` |
| Los lenguajes disponibles en el selector        | `CONFIG.languages` (usa nombres de [Prism](https://prismjs.com/#supported-languages)) |
| El tamaño de fuente por defecto                 | `CONFIG.fontSize` |
| Las palabras clave del pseudocódigo PSeInt      | `vendor/prism/prism-pseint.js` (arreglos `multiWordKeywords`, `singleWordKeywords`, `types`, `builtinFunctions`) |

### Agregar un nuevo esquema de color

Duplica cualquier bloque de `CONFIG.themes` (por ejemplo `dark-plus`), ponle
un id nuevo (por ejemplo `"solarized-dark"`), cámbiale el `label` (aparece
tal cual en el `<select>` del panel) y ajusta los colores. No hace falta
tocar `taskpane.js`: el selector se puebla automáticamente a partir de las
claves de `CONFIG.themes`.

### Agregar un ajuste de estilo para otro lenguaje

Sigue el patrón usado para PHP en `CONFIG.languageOverrides`:

```js
languageOverrides: {
  php: { variable: "#ff9d5c", delimiter: "#808080", package: "#4ec9b0" },
  python: { decorator: "#dcdcaa", "triple-quoted-string": "#e6a56e" },
}
```

Las claves deben coincidir con los nombres de token que genera la gramática
de Prism.js para ese lenguaje (revisa el archivo `prism-<lenguaje>.js`
correspondiente en el [repositorio de Prism](https://github.com/PrismJS/prism/tree/master/components)
si no estás seguro de los nombres exactos).

## Puesta en marcha (desarrollo local)

Los add-ins de Office deben servirse por **HTTPS**, incluso en desarrollo.

### 1. Instalar certificados de desarrollo confiables

```bash
npx --yes office-addin-dev-certs install
```

Esto instala un certificado autofirmado válido para `localhost` y lo agrega
a tu almacén de certificados de confianza (te pedirá confirmación).

### 2. Levantar el servidor local

```bash
npm start
```

Esto sirve la carpeta del proyecto en `https://localhost:3000` usando el
certificado generado en el paso 1 (ver el script `start` en `package.json`).

> **Windows / PowerShell**: si el script de `npm start` no resuelve `$HOME`
> correctamente, ejecuta directamente:
> ```powershell
> npx http-server . -p 3000 -S -C "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" -K "$env:USERPROFILE\.office-addin-dev-certs\localhost.key"
> ```

Verifica que `https://localhost:3000/manifest.xml` cargue sin advertencias
de certificado antes de continuar.

### 3. Sideload del add-in

**Opción A — Word/PowerPoint de escritorio (Windows/Mac):**
1. Abre Word o PowerPoint.
2. Ve a `Insertar` → `Complementos` → `Mis complementos` → `Subir mi
   complemento` (o el ícono de engranaje/"Cargar complemento personal") y
   selecciona `manifest.xml`.

**Opción B — Office en la web:**
1. Sube `manifest.xml` a una carpeta compartida (OneDrive/SharePoint)
   configurada como "catálogo de complementos" en el Centro de
   administración de Microsoft 365, o usa `Insertar` → `Agregar
   complementos` → `Cargar mi complemento` en Office en la web.

Una vez cargado, aparecerá el botón **"Insertar código"** en la pestaña
Inicio, tanto en Word como en PowerPoint (usando el mismo `manifest.xml`).

### 4. (Opcional) Validar el manifest

```bash
npm run validate
```

## Uso sin conexión (offline)

Prism.js y las gramáticas de los lenguajes de abajo están vendorizadas en
`vendor/prism/` (archivos locales, no un CDN), así que el resaltado funciona
sin internet. Lo único que sigue requiriendo red es la carga de **Office.js**
(`https://appsforoffice.microsoft.com/...` en `taskpane.html`) — eso es
inevitable: Microsoft no recomienda auto-hospedar Office.js porque recibe
parches de seguridad constantes, y además tu propio servidor local
(`npm start`) también tiene que seguir corriendo para que Office pueda pedirle
los archivos del panel por HTTPS.

**Lenguajes incluidos offline:** JavaScript, TypeScript, Python, Java, C#,
C / C++ (ver nota abajo), PHP, PHP + HTML, JSON, Bash/Shell, HTML/XML, CSS y
Pseudocódigo (PSeInt).

**Lenguajes que necesitan un archivo adicional** (Ruby, Go, Rust, Kotlin,
Swift, SQL, PowerShell, YAML): mientras no agregues su componente, el
selector los muestra igual, pero el código se inserta sin colorear (texto
plano) — no rompe nada, simplemente no hay gramática cargada para ese
lenguaje.

### Agregar más lenguajes

1. Con internet, descarga el componente minificado de Prism.js 1.29.0 para
   el lenguaje que quieras, por ejemplo:
   `https://cdn.jsdelivr.net/npm/prismjs@1.29.0/components/prism-ruby.min.js`
   (cambia `ruby` por el nombre del lenguaje; la lista completa está en
   [prismjs.com/#supported-languages](https://prismjs.com/#supported-languages)).
2. Guarda el archivo descargado en `vendor/prism/`.
3. Agrega un `<script>` para ese archivo en `taskpane.html`, después de
   `prism-core-bundle.min.js` (y después de cualquier otro lenguaje del que
   dependa — revisa el encabezado del archivo descargado o la
   [tabla de dependencias de Prism](https://github.com/PrismJS/prism/blob/master/components.json)
   si no estás seguro).
4. Si el lenguaje no está ya en `CONFIG.languages` (en `config.js`), agrega
   una entrada `{ label: "...", value: "..." }`.

**Nota sobre C / C++**: solo se vendorizó la gramática `c` (no `cpp`) para
mantener el paquete offline liviano; cubre bien C y la mayoría de C++ básico,
pero no resalta sintaxis exclusiva de C++ (clases, plantillas, namespaces).
Si la necesitas, seguí los pasos de arriba con `prism-cpp.min.js` y cambiá
`prismLang: "c"` por `prismLang: "cpp"` en la entrada correspondiente de
`CONFIG.languages`.

**Volver al modo "CDN + autoloader"** (todos los +290 lenguajes de Prism,
pero requiere internet siempre): en `taskpane.html`, reemplaza los `<script>`
de `vendor/prism/` por:
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/prism/1.29.0/prism.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/prism/1.29.0/plugins/autoloader/prism-autoloader.min.js"></script>
```
No hace falta tocar `highlighter.js`: `ensureLanguageLoaded()` ya detecta si
existe `Prism.plugins.autoloader` y, si existe, le pide la carga de la
gramática (y de `extraDeps`, como `markup` para "PHP + HTML") bajo demanda.
La única excepción es **PSeInt**, que al ser una gramática propia (no un
componente oficial de Prism) solo funciona si su `<script>` sigue apuntando
a `vendor/prism/prism-pseint.js`.

## Notas técnicas importantes

- **Por qué Word usa HTML y PowerPoint usa imagen**: `Office.CoercionType.Html`
  solo está soportado en Word y Outlook. PowerPoint únicamente admite
  `Office.CoercionType.Image` para contenido con múltiples colores/estilos en
  una sola inserción, por eso el snippet se renderiza a un canvas y se
  inserta como PNG en las diapositivas.
- **Por qué se ejecutan los hooks de Prism a mano**: el add-in tokeniza el
  código con `Prism.tokenize()` en vez de `Prism.highlightElement()`, para
  poder generar HTML propio y dibujar en canvas. El problema es que algunas
  gramáticas (como `php`, que usa `markup-templating` para resaltar el HTML
  fuera de `<?php ?>` como HTML real) dependen de los hooks
  `before-tokenize`/`after-tokenize` para funcionar, y esos hooks normalmente
  solo se disparan dentro de `Prism.highlight()`. Por eso
  `Highlighter.highlightToLines()` en `highlighter.js` ejecuta esos hooks
  manualmente alrededor de `Prism.tokenize()`, replicando lo que hace
  `Prism.highlight()` internamente. Sin esto, "PHP + HTML" tokenizaría el PHP
  correctamente pero el HTML alrededor quedaría sin resaltar.
- **Por qué "PHP + HTML" es una opción aparte de "PHP"**: en Prism.js, la
  gramática `php` solo carga automáticamente `markup-templating` (necesaria
  para poder ubicar los bloques `<?php ?>`), pero **no** carga `markup`
  (HTML) por defecto — es una dependencia opcional. Sin `markup` cargado, el
  HTML que rodea al PHP se muestra sin resaltar. La opción "PHP + HTML"
  simplemente asegura que `markup` también esté disponible (en el modo
  offline ya viene incluido en `prism-core-bundle.min.js`; en el modo CDN se
  pide explícitamente vía `extraDeps: ["markup"]`), y con eso el HTML
  embebido se resalta con tokens reales (`tag`, `attr-name`, `attr-value`,
  `comment`, etc. — colores que ya están definidos en todos los temas).
- **PSeInt no es un lenguaje de Prism.js**: `vendor/prism/prism-pseint.js`
  es una gramática escrita a mano para este proyecto (ver los comentarios
  dentro del archivo). Todos los tipos de token que usa (`keyword`,
  `class-name`, `boolean`, `constant`, `function`, `operator`, `string`,
  `number`, `comment`, `punctuation`) ya existen en `CONFIG.themes`, así que
  hereda automáticamente los colores de cualquier esquema sin necesitar
  overrides adicionales.
- **Antes de publicar**: cambia el `<Id>` (GUID) en `manifest.xml` por uno
  propio, y reemplaza `https://localhost:3000` por el dominio HTTPS real
  donde alojes los archivos (Azure Static Web Apps, GitHub Pages con
  dominio propio + HTTPS, IIS, etc.).

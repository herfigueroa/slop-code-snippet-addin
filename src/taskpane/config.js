/**
 * config.js
 * ---------------------------------------------------------------------------
 * TODA la configuración visual y de comportamiento del add-in vive aquí.
 * Este archivo está pensado para que puedas personalizar el add-in sin tocar
 * la lógica en taskpane.js ni highlighter.js.
 *
 * Cosas que puedes cambiar sin miedo:
 *  - CONFIG.fonts             -> tipografías monoespaciadas disponibles en el selector
 *  - CONFIG.defaultFontId     -> fuente seleccionada por defecto
 *  - CONFIG.themes            -> esquemas de color (paleta por tipo de token + tarjeta + título)
 *  - CONFIG.defaultThemeId    -> esquema de color seleccionado por defecto
 *  - CONFIG.languageOverrides -> ajustes de color específicos por lenguaje (ver ejemplo de PHP)
 *  - CONFIG.card              -> padding/radio/grosor de borde (no-color) de la tarjeta
 *  - CONFIG.title             -> peso/tamaño (no-color) de la barra de título opcional
 *  - CONFIG.languages         -> idiomas disponibles en el <select> del panel
 *  - CONFIG.fontSize          -> tamaño de fuente por defecto (en puntos)
 * ---------------------------------------------------------------------------
 */

const CONFIG = {
  // Tamaño de fuente por defecto (pt). El usuario puede ajustarlo en el panel.
  fontSize: 14,

  // Interlineado relativo (1.4 = 140% del tamaño de fuente).
  lineHeight: 1.5,

  // ---------------------------------------------------------------------
  // FUENTES DISPONIBLES
  // ---------------------------------------------------------------------
  // El usuario elige una desde el panel (o escribe una personalizada).
  // "family" es el valor CSS font-family completo, con fallbacks monoespaciados.
  // Nota: la fuente debe estar instalada en la máquina que ABRE el documento
  // para que se vea igual; por eso siempre incluimos fallbacks seguros.
  fonts: [
    { id: "cascadia", label: "Cascadia Code", family: "'Cascadia Code', 'Cascadia Mono', Consolas, monospace" },
    { id: "consolas", label: "Consolas", family: "Consolas, 'Courier New', monospace" },
    { id: "fira", label: "Fira Code", family: "'Fira Code', Consolas, monospace" },
    { id: "jetbrains", label: "JetBrains Mono", family: "'JetBrains Mono', Consolas, monospace" },
    { id: "source-code-pro", label: "Source Code Pro", family: "'Source Code Pro', Consolas, monospace" },
    { id: "courier", label: "Courier New", family: "'Courier New', monospace" },
    { id: "custom", label: "Personalizada…", family: null }, // el valor lo escribe el usuario en el panel
  ],
  defaultFontId: "cascadia",

  // ---------------------------------------------------------------------
  // ESQUEMAS DE COLOR (TEMAS)
  // ---------------------------------------------------------------------
  // Cada tema define: colores por tipo de token (Prism), color de la
  // tarjeta/fondo y color de la barra de título. Para agregar un tema nuevo,
  // copia un bloque completo, cámbiale el id/label y ajusta los colores.
  themes: {
    "dark-plus": {
      label: "Dark+ (VS Code)",
      card: { backgroundColor: "#1e1e1e", borderColor: "#3c3c3c" },
      title: { backgroundColor: "#333333", textColor: "#e6e6e6" },
      colors: {
        comment: "#6a9955", prolog: "#6a9955", doctype: "#6a9955", cdata: "#6a9955",
        punctuation: "#d4d4d4", operator: "#d4d4d4",
        property: "#9cdcfe", tag: "#569cd6", boolean: "#569cd6", number: "#b5cea8",
        constant: "#4fc1ff", symbol: "#b5cea8", deleted: "#ce9178",
        selector: "#d7ba7d", "attr-name": "#9cdcfe", string: "#ce9178", char: "#ce9178",
        builtin: "#4ec9b0", inserted: "#b5cea8", variable: "#9cdcfe", "attr-value": "#ce9178",
        keyword: "#569cd6", regex: "#d16969", important: "#569cd6",
        function: "#dcdcaa", "class-name": "#4ec9b0", package: "#4ec9b0", delimiter: "#808080",
        plain: "#d4d4d4",
      },
    },
    monokai: {
      label: "Monokai",
      card: { backgroundColor: "#272822", borderColor: "#49483e" },
      title: { backgroundColor: "#3e3d32", textColor: "#f8f8f2" },
      colors: {
        comment: "#75715e", prolog: "#75715e", doctype: "#75715e", cdata: "#75715e",
        punctuation: "#f8f8f2", operator: "#f92672",
        property: "#66d9ef", tag: "#f92672", boolean: "#ae81ff", number: "#ae81ff",
        constant: "#ae81ff", symbol: "#ae81ff", deleted: "#f92672",
        selector: "#a6e22e", "attr-name": "#a6e22e", string: "#e6db74", char: "#e6db74",
        builtin: "#66d9ef", inserted: "#a6e22e", variable: "#f8f8f2", "attr-value": "#e6db74",
        keyword: "#f92672", regex: "#e6db74", important: "#f92672",
        function: "#a6e22e", "class-name": "#a6e22e", package: "#a6e22e", delimiter: "#75715e",
        plain: "#f8f8f2",
      },
    },
    dracula: {
      label: "Dracula",
      card: { backgroundColor: "#282a36", borderColor: "#44475a" },
      title: { backgroundColor: "#343746", textColor: "#f8f8f2" },
      colors: {
        comment: "#6272a4", prolog: "#6272a4", doctype: "#6272a4", cdata: "#6272a4",
        punctuation: "#f8f8f2", operator: "#ff79c6",
        property: "#8be9fd", tag: "#ff79c6", boolean: "#bd93f9", number: "#bd93f9",
        constant: "#bd93f9", symbol: "#bd93f9", deleted: "#ff5555",
        selector: "#50fa7b", "attr-name": "#50fa7b", string: "#f1fa8c", char: "#f1fa8c",
        builtin: "#8be9fd", inserted: "#50fa7b", variable: "#f8f8f2", "attr-value": "#f1fa8c",
        keyword: "#ff79c6", regex: "#f1fa8c", important: "#ff79c6",
        function: "#50fa7b", "class-name": "#8be9fd", package: "#8be9fd", delimiter: "#6272a4",
        plain: "#f8f8f2",
      },
    },
    "a11y-dark": {
      label: "A11y Dark",
      card: { backgroundColor: "#2b2b2b", borderColor: "#444444" },
      title: { backgroundColor: "#363636", textColor: "#f8f8f2" },
      colors: {
        comment: "#d4d0ab", prolog: "#d4d0ab", doctype: "#d4d0ab", cdata: "#d4d0ab",
        punctuation: "#f8f8f2", operator: "#f8f8f2",
        property: "#abe338", tag: "#dcc6e0", boolean: "#ffa07a", number: "#dcc6e0",
        constant: "#00e0e0", symbol: "#dcc6e0", deleted: "#ffd700",
        selector: "#ffd700", "attr-name": "#abe338", string: "#ffd700", char: "#ffd700",
        builtin: "#ffd700", inserted: "#dcc6e0", variable: "#00e0e0", "attr-value": "#ffd700",
        keyword: "#ffa07a", regex: "#ffa07a", important: "#ffa07a",
        function: "#ffd700", "class-name": "#ffd700", package: "#ffd700", delimiter: "#d4d0ab",
        plain: "#f8f8f2",
      },
    },
    "github-light": {
      label: "GitHub Light",
      card: { backgroundColor: "#f6f8fa", borderColor: "#d0d7de" },
      title: { backgroundColor: "#eaeef2", textColor: "#24292f" },
      colors: {
        comment: "#6e7781", prolog: "#6e7781", doctype: "#6e7781", cdata: "#6e7781",
        punctuation: "#24292f", operator: "#0550ae",
        property: "#0550ae", tag: "#116329", boolean: "#0550ae", number: "#0550ae",
        constant: "#0550ae", symbol: "#0550ae", deleted: "#82071e",
        selector: "#116329", "attr-name": "#953800", string: "#0a3069", char: "#0a3069",
        builtin: "#8250df", inserted: "#116329", variable: "#953800", "attr-value": "#0a3069",
        keyword: "#cf222e", regex: "#0a3069", important: "#cf222e",
        function: "#8250df", "class-name": "#953800", package: "#953800", delimiter: "#6e7781",
        plain: "#24292f",
      },
    },
  },
  defaultThemeId: "dark-plus",

  // ---------------------------------------------------------------------
  // AJUSTES DE ESTILO POR LENGUAJE
  // ---------------------------------------------------------------------
  // Se aplican ENCIMA del esquema de color elegido, solo cuando el lenguaje
  // seleccionado coincide. Útil para resaltar particularidades de un
  // lenguaje sin tener que duplicar los 4 temas completos.
  //
  // Ajuste para PHP: Prism tokeniza el PHP con tipos propios que no existen
  // en otros lenguajes: 'delimiter' (las etiquetas <?php y ?>), 'variable'
  // (las variables $foo) y 'package' (namespace/use). Aquí les damos un
  // tratamiento visual distinto al del resto del código:
  //   - las variables $foo resaltan en un tono cálido, fácil de detectar,
  //   - las etiquetas <?php / ?> se atenúan (gris), como en editores tipo PhpStorm,
  //   - namespace/use quedan igual que los nombres de clase.
  languageOverrides: {
    php: {
      variable: "#ff9d5c",
      delimiter: "#808080",
      package: "#4ec9b0",
    },
  },

  // Estructura (no-color) de la tarjeta/fondo del snippet.
  card: {
    padding: 16, // px
    borderRadius: 8, // px (Word puede ignorar el radio, PowerPoint lo respeta al 100%)
    borderWidth: 1, // px, 0 = sin borde
  },

  // Estructura (no-color) de la barra de título opcional.
  title: {
    fontWeight: "600",
    fontSize: 12, // pt
  },

  // Ancho (px) de la imagen renderizada del snippet (usado por PowerPoint
  // siempre, y por Word cuando se activa "Insertar como imagen"). El
  // usuario lo puede ajustar desde el panel; "default" es el valor inicial
  // y "min"/"max" limitan el campo numérico. Si el código no entra en este
  // ancho, se ajusta línea por línea (word-wrap) asumiendo fuente
  // monoespaciada.
  imageWidth: {
    default: 920,
    min: 320,
    max: 2400,
  },

  // Escala de renderizado del canvas (2 = imagen a doble resolución para que
  // se vea nítida al escalarla dentro de la diapositiva).
  canvasScale: 2,

  // Idiomas disponibles en el selector. "value" es lo que ve/elige el usuario
  // y lo que se usa como clave de CONFIG.languageOverrides y para el <select>.
  // "prismLang" es el nombre de gramática real de Prism.js a cargar/tokenizar
  // (si se omite, se asume igual a "value"). "extraDeps" son componentes de
  // Prism adicionales que se cargan junto con "prismLang" (por ejemplo,
  // 'markup' para que el HTML embebido en PHP también se resalte).
  languages: [
    { label: "JavaScript", value: "javascript" },
    { label: "TypeScript", value: "typescript" },
    { label: "Python", value: "python" },
    { label: "Java", value: "java" },
    { label: "C#", value: "csharp" },
    {
      label: "C / C++",
      value: "cpp",
      // Solo vendorizamos la gramática 'c' (no 'cpp') para mantener el modo
      // offline liviano. 'c' cubre la enorme mayoría de la sintaxis de C++
      // básico (funciones, control de flujo, tipos, comentarios); lo que NO
      // resalta correctamente es sintaxis exclusiva de C++ (clases, plantillas,
      // namespaces). Si necesitas eso, agrega vendor/prism/prism-cpp.min.js
      // (ver README → "Agregar más lenguajes") y cambia esto a prismLang: "cpp".
      prismLang: "c",
    },
    { label: "PHP", value: "php" },
    {
      label: "PHP + HTML (embebido)",
      value: "php-html",
      prismLang: "php",
      // 'markup' NO se carga automáticamente junto con 'php' (es una
      // dependencia opcional, no obligatoria, en Prism.js). Al agregarlo
      // aquí, el HTML que quede fuera de las etiquetas <?php ?> también se
      // resalta como HTML real (tags, atributos, comentarios, etc.).
      extraDeps: ["markup"],
    },
    { label: "Ruby", value: "ruby" },
    { label: "Go", value: "go" },
    { label: "Rust", value: "rust" },
    { label: "Kotlin", value: "kotlin" },
    { label: "Swift", value: "swift" },
    { label: "SQL", value: "sql" },
    { label: "Bash / Shell", value: "bash" },
    { label: "PowerShell", value: "powershell" },
    { label: "JSON", value: "json" },
    { label: "YAML", value: "yaml" },
    { label: "HTML / XML", value: "markup" },
    { label: "CSS", value: "css" },
    {
      // Gramática escrita a mano (no es un lenguaje oficial de Prism.js) en
      // vendor/prism/prism-pseint.js, a partir del diccionario de palabras
      // clave de PSeInt (es.json) y su sintaxis de Vim (pseint.vim). Ver ese
      // archivo para agregar variantes de palabras clave o soporte de otro
      // idioma.
      label: "Pseudocódigo (PSeInt)",
      value: "pseint",
    },
    { label: "Texto plano (sin resaltado)", value: "none" },
  ],
};

// Se expone en window para que taskpane.js y highlighter.js puedan usarlo
// tanto si se cargan como <script> clásicos como si se migran a módulos ES.
window.CONFIG = CONFIG;

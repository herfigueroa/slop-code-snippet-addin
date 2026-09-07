/**
 * highlighter.js
 * ---------------------------------------------------------------------------
 * Encargado de convertir código fuente + lenguaje -> tokens coloreados.
 *
 * Usa Prism.js (cargado vía CDN en taskpane.html) para tokenizar, pero en
 * lugar de depender de una hoja de estilos externa (que Word/PowerPoint no
 * conservan al insertar contenido), generamos HTML con estilos EN LÍNEA
 * (inline style="color:...") tomados de CONFIG.theme. Esto garantiza que el
 * color se mantenga sin importar dónde se pegue el HTML.
 *
 * También exponemos `highlightToLines()` que aplana el árbol de tokens en un
 * arreglo de líneas -> segmentos {text, color}, útil para dibujar el código
 * carácter por carácter sobre un <canvas> (necesario para insertarlo como
 * imagen en PowerPoint).
 * ---------------------------------------------------------------------------
 */

const Highlighter = (() => {
  /**
   * Busca la entrada de CONFIG.languages correspondiente a un "value" del
   * selector (ej. "php-html") y devuelve siempre un objeto con al menos
   * { value, prismLang, extraDeps }. Si el value no tiene entrada explícita
   * en CONFIG.languages (no debería pasar, pero por robustez), se asume que
   * "value" y "prismLang" son lo mismo y sin dependencias extra.
   */
  function getLanguageEntry(langValue) {
    const list = (window.CONFIG && window.CONFIG.languages) || [];
    const found = list.find((l) => l.value === langValue);
    if (found) {
      return {
        value: found.value,
        prismLang: found.prismLang || found.value,
        extraDeps: found.extraDeps || [],
      };
    }
    return { value: langValue, prismLang: langValue, extraDeps: [] };
  }

  /**
   * Se asegura de que la gramática de Prism necesaria (y sus dependencias
   * extra, ej. 'markup' para PHP+HTML) estén cargadas antes de tokenizar.
   *
   * IMPORTANTE: el plugin prism-autoloader normalmente se activa mediante el
   * hook 'complete', que solo se dispara al usar Prism.highlightElement().
   * Como este add-in tokeniza manualmente con Prism.tokenize() (para poder
   * generar HTML propio / dibujar en canvas), ese hook nunca se dispara por
   * sí solo. Por eso llamamos explícitamente a
   * Prism.plugins.autoloader.loadLanguages(), que es la API pública
   * documentada para pedir la carga de uno o más lenguajes bajo demanda.
   *
   * `callback` se invoca siempre (éxito o error) para no bloquear la UI; si
   * la carga falla, resolveGrammar simplemente no encontrará la gramática y
   * el snippet se mostrará como texto plano en vez de romper el add-in.
   */
  function ensureLanguageLoaded(langValue, callback) {
    if (!langValue || langValue === "none") {
      callback();
      return;
    }
    const entry = getLanguageEntry(langValue);
    const toLoad = [entry.prismLang].concat(entry.extraDeps);
    const autoloader = window.Prism && window.Prism.plugins && window.Prism.plugins.autoloader;
    if (!autoloader) {
      // El plugin no cargó (por ejemplo, sin conexión a internet). Se sigue
      // adelante: resolveGrammar caerá a texto plano si falta la gramática.
      callback();
      return;
    }
    autoloader.loadLanguages(toLoad, callback, callback);
  }

  /**
   * Determina si Prism reconoce el lenguaje solicitado. Si no lo reconoce
   * (o el usuario eligió "Texto plano"), se trata todo como texto simple.
   */
  function resolveGrammar(langValue) {
    if (!langValue || langValue === "none") return null;
    const entry = getLanguageEntry(langValue);
    return window.Prism && window.Prism.languages ? window.Prism.languages[entry.prismLang] : null;
  }

  /**
   * Recorre recursivamente el árbol de tokens de Prism y produce un arreglo
   * plano de segmentos { text, className }.
   * Prism.tokenize devuelve un arreglo mezclado de strings y objetos Token
   * (que a su vez pueden contener sub-arreglos anidados).
   */
  function flattenTokens(tokens, className, out) {
    for (const token of tokens) {
      if (typeof token === "string") {
        if (token.length > 0) out.push({ text: token, className: className || "plain" });
      } else {
        const cls = token.type || className || "plain";
        if (Array.isArray(token.content)) {
          flattenTokens(token.content, cls, out);
        } else if (typeof token.content === "string") {
          out.push({ text: token.content, className: cls });
        } else {
          // Token anidado con contenido tipo Token único
          flattenTokens([token.content], cls, out);
        }
      }
    }
    return out;
  }

  /**
   * Tokeniza el código y devuelve un arreglo plano de segmentos coloreados
   * (sin dividir por líneas todavía).
   *
   * IMPORTANTE: no llamamos a Prism.tokenize() a secas. Gramáticas basadas
   * en "templating" (como php, que usa markup-templating para resaltar el
   * HTML fuera de <?php ?> como HTML real) dependen de los hooks
   * 'before-tokenize' y 'after-tokenize' para funcionar: ahí es donde php.js
   * reemplaza temporalmente los bloques <?php ?> por placeholders, cambia la
   * gramática activa a 'markup' para tokenizar el HTML circundante, y luego
   * reinserta el PHP ya tokenizado en su lugar. Esos hooks normalmente los
   * dispara Prism.highlight(), no Prism.tokenize(); como este add-in nunca
   * llama a Prism.highlight() (tokenizamos manualmente para poder generar
   * HTML propio y dibujar en canvas), replicamos aquí el mismo patrón que
   * usa Prism.highlight() internamente para no perder ese comportamiento.
   */
  function tokenizeFlat(code, langValue) {
    const entry = getLanguageEntry(langValue);
    const grammar = resolveGrammar(langValue);
    if (!grammar) {
      return [{ text: code, className: "plain" }];
    }
    try {
      const env = { code: code, grammar: grammar, language: entry.prismLang };
      window.Prism.hooks.run("before-tokenize", env);
      env.tokens = window.Prism.tokenize(env.code, env.grammar);
      window.Prism.hooks.run("after-tokenize", env);
      return flattenTokens(env.tokens, "plain", []);
    } catch (e) {
      // Si algo falla en la tokenización, degradamos a texto plano en vez
      // de romper la inserción del snippet.
      console.warn("Prism no pudo tokenizar, usando texto plano:", e);
      return [{ text: code, className: "plain" }];
    }
  }

  /**
   * Divide un arreglo plano de segmentos {text, className} en líneas,
   * respetando los saltos de línea dentro de un mismo token.
   * Devuelve: [ [ {text, className}, ... ], [ ... ] ]  (un arreglo por línea)
   */
  function segmentsToLines(segments) {
    const lines = [[]];
    for (const seg of segments) {
      const parts = seg.text.split("\n");
      parts.forEach((part, i) => {
        if (i > 0) lines.push([]);
        if (part.length > 0) {
          lines[lines.length - 1].push({ text: part, className: seg.className });
        }
      });
    }
    return lines;
  }

  /**
   * Devuelve las líneas ya coloreadas y listas para:
   *  a) generar HTML (ver highlightToHtml)
   *  b) dibujar en canvas (ver renderer.js)
   */
  function highlightToLines(code, lang) {
    return segmentsToLines(tokenizeFlat(code, lang));
  }

  /**
   * Combina los colores del tema elegido con los overrides específicos del
   * lenguaje activo (ver CONFIG.languageOverrides), si existen. El override
   * siempre gana sobre el color base del tema. Se busca por "prismLang" (no
   * por el "value" del selector) para que variantes como "PHP + HTML"
   * hereden automáticamente el mismo ajuste que "PHP".
   */
  function resolveColors(themeId, langValue) {
    const cfg = window.CONFIG;
    const theme = cfg.themes[themeId] || cfg.themes[cfg.defaultThemeId];
    const entry = getLanguageEntry(langValue);
    const overrides = (cfg.languageOverrides && cfg.languageOverrides[entry.prismLang]) || {};
    return Object.assign({}, theme.colors, overrides);
  }

  function colorFor(className, colors) {
    const map = colors || window.CONFIG.themes[window.CONFIG.defaultThemeId].colors;
    return map[className] || map.plain || "#d4d4d4";
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /**
   * Genera HTML con <span style="color:..."> por cada token, uniendo líneas
   * con <br/> (más confiable que \n dentro de elementos insertados vía
   * Office.CoercionType.Html en Word).
   *
   * `colors` es el mapa ya resuelto por resolveColors(themeId, lang). Si se
   * omite, se usa el tema por defecto sin overrides de lenguaje.
   */
  function highlightToHtml(code, lang, colors) {
    const lines = highlightToLines(code, lang);
    const colorMap = colors || resolveColors(window.CONFIG.defaultThemeId, lang);
    const htmlLines = lines.map((line) => {
      if (line.length === 0) return "&nbsp;";
      return line
        .map((seg) => `<span style="color:${colorFor(seg.className, colorMap)};">${escapeHtml(seg.text)}</span>`)
        .join("");
    });
    return htmlLines.join("<br/>");
  }

  return {
    getLanguageEntry,
    ensureLanguageLoaded,
    highlightToLines,
    highlightToHtml,
    resolveColors,
    colorFor,
    escapeHtml,
  };
})();

window.Highlighter = Highlighter;

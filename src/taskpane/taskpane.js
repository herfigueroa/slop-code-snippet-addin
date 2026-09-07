/**
 * taskpane.js
 * ---------------------------------------------------------------------------
 * Conecta la interfaz del panel de tareas con:
 *   - highlighter.js  (tokenizado / colores)
 *   - renderer.js     (genera HTML para Word / imagen para PowerPoint)
 *   - Office.js       (inserta el resultado en el documento activo)
 *
 * Flujo:
 *   1. Office.onReady detecta si estamos en Word o PowerPoint.
 *   2. Se puebla el <select> de lenguajes desde CONFIG.languages.
 *   3. Cualquier cambio en los campos actualiza la vista previa en vivo.
 *   4. Al pulsar "Insertar", se genera el contenido final según el host
 *      (HTML nativo en Word / imagen PNG en PowerPoint) y se inserta en la
 *      posición actual del cursor / selección (reemplazándola si había algo
 *      seleccionado, lo que permite "editar" un snippet ya insertado).
 * ---------------------------------------------------------------------------
 */

let currentHost = null; // Office.HostType.Word | Office.HostType.PowerPoint

const els = {}; // referencias a elementos del DOM, llenadas en initDom()

Office.onReady((info) => {
  currentHost = info.host;
  initDom();
  populateLanguages();
  populateFonts();
  populateThemes();
  restoreDefaults();
  attachEventListeners();
  updateHostUI();
  refreshPreview(); // carga la gramática del lenguaje por defecto y luego renderiza
});

function initDom() {
  els.hostLabel = document.getElementById("hostLabel");
  els.codeInput = document.getElementById("codeInput");
  els.languageSelect = document.getElementById("languageSelect");
  els.fontSizeInput = document.getElementById("fontSizeInput");
  els.fontFamilySelect = document.getElementById("fontFamilySelect");
  els.customFontInput = document.getElementById("customFontInput");
  els.themeSelect = document.getElementById("themeSelect");
  els.showTitleToggle = document.getElementById("showTitleToggle");
  els.titleInput = document.getElementById("titleInput");
  els.backgroundSection = document.getElementById("backgroundSection");
  els.showBackgroundToggle = document.getElementById("showBackgroundToggle");
  els.backgroundHint = document.getElementById("backgroundHint");
  els.previewWrapper = document.getElementById("previewWrapper");
  els.previewTitle = document.getElementById("previewTitle");
  els.previewCode = document.getElementById("previewCode");
  els.previewCodeInner = document.getElementById("previewCodeInner");
  els.insertButton = document.getElementById("insertButton");
  els.statusMessage = document.getElementById("statusMessage");
}

function populateLanguages() {
  els.languageSelect.innerHTML = "";
  window.CONFIG.languages.forEach((lang) => {
    const opt = document.createElement("option");
    opt.value = lang.value;
    opt.textContent = lang.label;
    els.languageSelect.appendChild(opt);
  });
  els.languageSelect.value = "javascript";
}

function populateFonts() {
  els.fontFamilySelect.innerHTML = "";
  window.CONFIG.fonts.forEach((font) => {
    const opt = document.createElement("option");
    opt.value = font.id;
    opt.textContent = font.label;
    els.fontFamilySelect.appendChild(opt);
  });
}

function populateThemes() {
  els.themeSelect.innerHTML = "";
  Object.keys(window.CONFIG.themes).forEach((themeId) => {
    const opt = document.createElement("option");
    opt.value = themeId;
    opt.textContent = window.CONFIG.themes[themeId].label;
    els.themeSelect.appendChild(opt);
  });
}

function restoreDefaults() {
  els.fontSizeInput.value = window.CONFIG.fontSize;
  els.fontFamilySelect.value = window.CONFIG.defaultFontId;
  els.customFontInput.hidden = window.CONFIG.defaultFontId !== "custom";
  els.themeSelect.value = window.CONFIG.defaultThemeId;
  els.showTitleToggle.checked = false; // desactivado por defecto, según lo solicitado
  els.titleInput.disabled = true;
  els.showBackgroundToggle.checked = false; // desactivado por defecto
  els.codeInput.value = "";
}

/**
 * Devuelve el valor CSS font-family final, resolviendo la fuente
 * "Personalizada…" a partir de lo escrito en #customFontInput.
 */
function getResolvedFontFamily() {
  const selected = window.CONFIG.fonts.find((f) => f.id === els.fontFamilySelect.value);
  if (!selected) return window.CONFIG.fonts[0].family;
  if (selected.id === "custom") {
    const custom = els.customFontInput.value.trim();
    return custom.length > 0 ? `${custom}, monospace` : "monospace";
  }
  return selected.family;
}

/**
 * En PowerPoint el "fondo" es solo un fill del canvas (siempre disponible y
 * barato), pero el usuario pidió específicamente que la opción de fondo sea
 * relevante "en caso de agregarlo en un documento de Word". Para reflejar
 * esa distinción sin quitarle la opción a PowerPoint, ajustamos únicamente
 * el texto de ayuda según el host detectado.
 */
function updateHostUI() {
  const isWord = currentHost === Office.HostType.Word;
  const isPpt = currentHost === Office.HostType.PowerPoint;

  if (isWord) {
    els.hostLabel.textContent = "Word detectado — inserción como contenido nativo editable";
    els.backgroundHint.textContent =
      "Envuelve el código en una tabla con fondo oscuro, borde y padding (formato nativo de Word).";
  } else if (isPpt) {
    els.hostLabel.textContent = "PowerPoint detectado — inserción como imagen de alta resolución";
    els.backgroundHint.textContent =
      "Dibuja una tarjeta con fondo detrás del código en la imagen insertada en la diapositiva.";
  } else {
    els.hostLabel.textContent = "Aplicación no reconocida (se requiere Word o PowerPoint)";
    els.insertButton.disabled = true;
  }
}

function attachEventListeners() {
  els.codeInput.addEventListener("input", renderPreview);
  els.languageSelect.addEventListener("change", refreshPreview);
  els.fontSizeInput.addEventListener("input", renderPreview);

  els.showTitleToggle.addEventListener("change", () => {
    els.titleInput.disabled = !els.showTitleToggle.checked;
    renderPreview();
  });
  els.titleInput.addEventListener("input", renderPreview);

  els.showBackgroundToggle.addEventListener("change", renderPreview);

  els.fontFamilySelect.addEventListener("change", () => {
    els.customFontInput.hidden = els.fontFamilySelect.value !== "custom";
    renderPreview();
  });
  els.customFontInput.addEventListener("input", renderPreview);
  els.themeSelect.addEventListener("change", renderPreview);

  els.insertButton.addEventListener("click", handleInsert);
}

function getCurrentOptions() {
  return {
    code: els.codeInput.value,
    lang: els.languageSelect.value,
    title: els.showTitleToggle.checked ? els.titleInput.value : null,
    showBackground: els.showBackgroundToggle.checked,
    fontSize: parseInt(els.fontSizeInput.value, 10) || window.CONFIG.fontSize,
    fontFamily: getResolvedFontFamily(),
    themeId: els.themeSelect.value || window.CONFIG.defaultThemeId,
  };
}

// ------------------------------------------------------------ VISTA PREVIA

/**
 * Se asegura de que la gramática del lenguaje seleccionado (y sus
 * dependencias extra, ej. 'markup' para "PHP + HTML") esté cargada antes de
 * volver a renderizar. Se usa al iniciar el panel y cada vez que el usuario
 * cambia de lenguaje. El resto de los cambios (código, tema, fuente, título,
 * fondo) no requieren recargar nada y llaman a renderPreview() directamente.
 */
function refreshPreview() {
  const lang = els.languageSelect.value;
  setStatus("Cargando resaltado…", "");
  window.Highlighter.ensureLanguageLoaded(lang, () => {
    setStatus("", "");
    renderPreview();
  });
}

function ensureLanguageLoadedAsync(lang) {
  return new Promise((resolve) => {
    window.Highlighter.ensureLanguageLoaded(lang, resolve);
  });
}

function renderPreview() {
  const opts = getCurrentOptions();

  // Título
  if (opts.title && opts.title.trim().length > 0) {
    els.previewTitle.hidden = false;
    els.previewTitle.textContent = opts.title;
  } else {
    els.previewTitle.hidden = true;
  }

  // La vista previa siempre refleja el tema y la tarjeta activos, tal como
  // se verán al insertar (con fondo). El toggle "Agregar fondo" solo decide
  // si esa tarjeta viaja o no al documento final.
  const theme = window.CONFIG.themes[opts.themeId] || window.CONFIG.themes[window.CONFIG.defaultThemeId];
  els.previewWrapper.style.background = theme.card.backgroundColor;
  els.previewWrapper.style.borderColor = theme.card.borderColor;
  els.previewTitle.style.background = theme.title.backgroundColor;
  els.previewTitle.style.color = theme.title.textColor;

  const colors = window.Highlighter.resolveColors(opts.themeId, opts.lang);
  const html = window.Highlighter.highlightToHtml(
    opts.code || "// Escribe código para ver el resaltado…",
    opts.lang,
    colors
  );
  els.previewCodeInner.innerHTML = html;
  els.previewCodeInner.style.fontSize = `${opts.fontSize}px`;
  els.previewCodeInner.style.color = colors.plain;
  els.previewCode.style.fontFamily = opts.fontFamily;
}

// --------------------------------------------------------------- INSERCIÓN
async function handleInsert() {
  const opts = getCurrentOptions();

  if (!opts.code || opts.code.trim().length === 0) {
    setStatus("Escribe algo de código antes de insertar.", "error");
    return;
  }

  setStatus("Insertando…", "");
  els.insertButton.disabled = true;

  try {
    await ensureLanguageLoadedAsync(opts.lang);
    if (currentHost === Office.HostType.Word) {
      await insertIntoWord(opts);
    } else if (currentHost === Office.HostType.PowerPoint) {
      await insertIntoPowerPoint(opts);
    } else {
      throw new Error("Host no soportado.");
    }
    setStatus("Snippet insertado correctamente.", "success");
  } catch (err) {
    console.error(err);
    setStatus("No se pudo insertar el snippet: " + (err && err.message ? err.message : err), "error");
  } finally {
    els.insertButton.disabled = false;
  }
}

function insertIntoWord(opts) {
  const html = window.Renderer.renderWordHtml(opts);
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
      } else {
        resolve();
      }
    });
  });
}

function insertIntoPowerPoint(opts) {
  const dataUrl = window.Renderer.renderPptImage(opts);
  const base64 = dataUrl.split(",")[1]; // quitar el prefijo "data:image/png;base64,"
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
      } else {
        resolve();
      }
    });
  });
}

function setStatus(message, kind) {
  els.statusMessage.textContent = message;
  els.statusMessage.className = "status-message" + (kind ? " " + kind : "");
}

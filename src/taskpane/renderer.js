/**
 * renderer.js
 * ---------------------------------------------------------------------------
 * Dos formas de "materializar" un snippet, según destino:
 *
 *  - Word (renderWordHtml):
 *      Genera HTML real (tabla con una celda) que Word convierte en
 *      contenido NATIVO y editable al insertarlo vía
 *      Office.CoercionType.Html. El color, la fuente monoespaciada y el
 *      fondo/tarjeta quedan como formato de Word normal.
 *
 *  - PowerPoint (renderPptImage):
 *      PowerPoint (a través de Office.js) NO soporta insertar texto con
 *      múltiples colores por caracter de forma nativa, así que dibujamos el
 *      snippet completo sobre un <canvas> (con la misma paleta de colores
 *      que la vista previa) y lo insertamos como imagen PNG vía
 *      Office.CoercionType.Image. Esto garantiza fidelidad visual exacta.
 * ---------------------------------------------------------------------------
 */

const Renderer = (() => {
  // ------------------------------------------------------------------ WORD
  /**
   * options:
   *  - code, lang
   *  - title: string | null
   *  - showBackground: boolean
   *  - fontSize: number (pt)
   *  - fontFamily: string (valor CSS font-family ya resuelto)
   *  - themeId: string (clave dentro de CONFIG.themes)
   */
  function renderWordHtml(options) {
    const { code, lang, title, showBackground, fontSize, fontFamily, themeId } = options;
    const cfg = window.CONFIG;
    const theme = cfg.themes[themeId] || cfg.themes[cfg.defaultThemeId];
    const colors = window.Highlighter.resolveColors(themeId, lang);
    const codeHtml = window.Highlighter.highlightToHtml(code, lang, colors);

    const cellBg = showBackground ? theme.card.backgroundColor : "transparent";
    const border = showBackground
      ? `border:${cfg.card.borderWidth}px solid ${theme.card.borderColor};`
      : "border:none;";
    const textColor = showBackground ? colors.plain : "#000000";

    const titleRow =
      title && title.trim().length > 0
        ? `<tr>
             <td style="background-color:${theme.title.backgroundColor};
                        color:${theme.title.textColor};
                        font-family:${fontFamily};
                        font-weight:${cfg.title.fontWeight};
                        font-size:${cfg.title.fontSize}pt;
                        padding:6px ${cfg.card.padding}px;
                        border-top-left-radius:${cfg.card.borderRadius}px;
                        border-top-right-radius:${cfg.card.borderRadius}px;">
               ${window.Highlighter.escapeHtml(title)}
             </td>
           </tr>`
        : "";

    // Se usa una tabla de una sola celda: es la forma más confiable de que
    // Word conserve el color de fondo (shading) y el padding al importar
    // HTML mediante Office.CoercionType.Html.
    return `
      <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;${border}">
        ${titleRow}
        <tr>
          <td style="background-color:${cellBg};
                     padding:${cfg.card.padding}px;
                     font-family:${fontFamily};
                     font-size:${fontSize}pt;
                     line-height:${cfg.lineHeight};
                     color:${textColor};
                     white-space:pre;">
            ${codeHtml}
          </td>
        </tr>
      </table>
    `;
  }

  // ------------------------------------------------------------- POWERPOINT
  /**
   * Dibuja el snippet en un canvas y devuelve un dataURL base64 PNG.
   * options: igual que renderWordHtml, más:
   *  - showBackground: en PowerPoint el fondo/tarjeta SIEMPRE ayuda a la
   *    legibilidad sobre la diapositiva, pero respetamos la preferencia del
   *    usuario igualmente.
   */
  function renderPptImage(options) {
    const { code, lang, title, showBackground, fontSize, fontFamily, themeId } = options;
    const cfg = window.CONFIG;
    const theme = cfg.themes[themeId] || cfg.themes[cfg.defaultThemeId];
    const colors = window.Highlighter.resolveColors(themeId, lang);
    const scale = cfg.canvasScale || 2;
    const lines = window.Highlighter.highlightToLines(code, lang);

    const padding = cfg.card.padding;
    const lineHeightPx = Math.round(fontSize * cfg.lineHeight * (96 / 72)); // pt -> px approx
    const fontPx = Math.round(fontSize * (96 / 72));
    const titleHeightPx = title && title.trim().length > 0 ? Math.round(cfg.title.fontSize * (96 / 72) * 2.4) : 0;

    // Canvas temporal solo para medir anchos de texto.
    const measureCanvas = document.createElement("canvas");
    const measureCtx = measureCanvas.getContext("2d");
    measureCtx.font = `${fontPx}px ${fontFamily}`;

    let maxLineWidth = 0;
    lines.forEach((line) => {
      const width = line.reduce((acc, seg) => acc + measureCtx.measureText(seg.text).width, 0);
      if (width > maxLineWidth) maxLineWidth = width;
    });
    if (title) {
      measureCtx.font = `${Math.round(cfg.title.fontSize * (96 / 72))}px ${fontFamily}`;
      maxLineWidth = Math.max(maxLineWidth, measureCtx.measureText(title).width);
    }

    const contentWidth = Math.min(maxLineWidth + padding * 2, cfg.maxCanvasWidth);
    const contentHeight = titleHeightPx + lines.length * lineHeightPx + padding * 2;

    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(contentWidth * scale);
    canvas.height = Math.ceil(contentHeight * scale);
    const ctx = canvas.getContext("2d");
    ctx.scale(scale, scale);

    // Fondo general (tarjeta) - si el usuario desactiva el fondo, usamos
    // transparente para que la imagen se mezcle con el diseño de la slide.
    if (showBackground) {
      ctx.fillStyle = theme.card.backgroundColor;
      roundRect(ctx, 0, 0, contentWidth, contentHeight, cfg.card.borderRadius);
      ctx.fill();
      if (cfg.card.borderWidth > 0) {
        ctx.lineWidth = cfg.card.borderWidth;
        ctx.strokeStyle = theme.card.borderColor;
        roundRect(ctx, 0.5, 0.5, contentWidth - 1, contentHeight - 1, cfg.card.borderRadius);
        ctx.stroke();
      }
    } else {
      ctx.clearRect(0, 0, contentWidth, contentHeight);
    }

    let cursorY = padding;

    // Barra de título opcional
    if (title && title.trim().length > 0) {
      ctx.fillStyle = theme.title.backgroundColor;
      ctx.fillRect(0, 0, contentWidth, titleHeightPx);
      ctx.fillStyle = theme.title.textColor;
      ctx.font = `${cfg.title.fontWeight} ${Math.round(cfg.title.fontSize * (96 / 72))}px ${fontFamily}`;
      ctx.textBaseline = "middle";
      ctx.fillText(title, padding, titleHeightPx / 2);
      cursorY = titleHeightPx + padding;
    }

    // Código, línea por línea, segmento por segmento (coloreado)
    ctx.font = `${fontPx}px ${fontFamily}`;
    ctx.textBaseline = "top";
    lines.forEach((line, idx) => {
      let cursorX = padding;
      const y = cursorY + idx * lineHeightPx;
      if (line.length === 0) return;
      line.forEach((seg) => {
        ctx.fillStyle = window.Highlighter.colorFor(seg.className, colors);
        ctx.fillText(seg.text, cursorX, y);
        cursorX += ctx.measureText(seg.text).width;
      });
    });

    return canvas.toDataURL("image/png");
  }

  function roundRect(ctx, x, y, width, height, radius) {
    if (radius === 0) {
      ctx.rect(x, y, width, height);
      return;
    }
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.arcTo(x + width, y, x + width, y + height, radius);
    ctx.arcTo(x + width, y + height, x, y + height, radius);
    ctx.arcTo(x, y + height, x, y, radius);
    ctx.arcTo(x, y, x + width, y, radius);
    ctx.closePath();
  }

  return { renderWordHtml, renderPptImage };
})();

window.Renderer = Renderer;

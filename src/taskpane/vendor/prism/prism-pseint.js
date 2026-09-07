/**
 * prism-pseint.js
 * ---------------------------------------------------------------------------
 * PSeInt no es un lenguaje soportado oficialmente por Prism.js, así que esta
 * gramática se escribió a mano combinando dos fuentes:
 *
 *  - es.json: el diccionario oficial de palabras clave en español de PSeInt
 *    (el mismo formato que usa PSeInt para sus paquetes de idioma).
 *  - pseint.vim: el archivo de sintaxis de Vim para PSeInt, que confirma
 *    detalles de comportamiento como que las palabras clave son
 *    insensibles a mayúsculas/minúsculas y cómo se arma cada patrón
 *    (comentarios con //, cadenas con comillas dobles, asignación con
 *    "<-" o "=", etc.).
 *
 * IMPORTANTE — cómo modificar esta gramática:
 *  - Todas las palabras clave están en los arreglos de abajo, agrupadas por
 *    categoría. Para agregar una variante de escritura (o soportar otro
 *    idioma de PSeInt), solo hay que añadir el string al arreglo
 *    correspondiente; no hace falta tocar el resto del archivo.
 *  - Las frases de más de una palabra (ej. "Sino Si", "Hasta Que") DEBEN
 *    declararse en `multiWordKeywords`, no en `singleWordKeywords`, y se
 *    procesan primero para que Prism no las corte a la mitad.
 *  - Los tipos de token usados (keyword, class-name, boolean, constant,
 *    function, operator, string, number, comment, punctuation) son los
 *    mismos que ya usan el resto de lenguajes del add-in, así que este
 *    pseudocódigo automáticamente hereda los colores de cualquier esquema
 *    de color (CONFIG.themes) sin necesitar overrides adicionales.
 * ---------------------------------------------------------------------------
 */
(function (Prism) {
  // Palabras clave de más de una palabra (deben ir antes que las de una sola
  // palabra para que, por ejemplo, "Hasta Que" no quede cortado en "Hasta").
  var multiWordKeywords = [
    "Sino\\s+Si",
    "De\\s+Otro\\s+Modo",
    "Mientras\\s+Que",
    "Con\\s+Paso",
    "Hasta\\s+Que",
    "Por\\s+Referencia",
    "Por\\s+Valor",
    "Escribir\\s+Sin\\s+Saltar",
    "Mostrar\\s+Sin\\s+Saltar",
    "Limpiar\\s+Pantalla",
    "Borrar\\s+Pantalla",
    "Esperar\\s+Tecla",
  ];

  // Palabras clave de control de flujo, declaración e I/O (una sola palabra).
  var singleWordKeywords = [
    // Programa
    "Proceso", "Algoritmo", "FinProceso", "FinAlgoritmo",
    // Declaraciones
    "Definir", "Como", "Constante", "Dimension",
    // Si / Segun
    "Si", "Entonces", "Sino", "FinSi",
    "Segun", "FinSegun",
    // Mientras / Para / Repetir
    "Mientras", "Hacer", "FinMientras",
    "Para", "Hasta", "Paso", "FinPara",
    "Repetir", "Romper", "Continuar",
    // Subprocesos / funciones
    "SubProceso", "SubAlgoritmo", "Procedimiento",
    "FinSubProceso", "FinSubAlgoritmo", "FinProcedimiento",
    "Funcion", "FinFuncion", "Retornar",
    // Entrada / salida
    "Escribir", "Mostrar", "Imprimir", "Leer", "Esperar",
  ];

  // Tipos de dato -> se resaltan como "class-name" para distinguirlos
  // visualmente de las palabras clave de control de flujo.
  var types = [
    "Entero", "Enteros",
    "Real", "Reales",
    "Cadena", "Cadenas",
    "Caracter", "Caracteres",
    "Texto",
    "Logico", "Logicos",
  ];

  // Funciones incorporadas (matemáticas, de texto, etc.) -> "function".
  var builtinFunctions = [
    "Abs", "RC", "Raiz", "Ln", "Exp",
    "Sen", "Cos", "Tan", "ASen", "ACos", "ATan",
    "Trunc", "Truncar", "Redon", "Redondear",
    "Azar", "Aleatorio",
    "Longitud", "Mayusculas", "Minusculas",
    "Subcadena", "Concatenar",
    "ConvertirANumero", "ConvertirATexto", "ConvertirACadena",
  ];

  function alternation(words) {
    return words.join("|");
  }

  Prism.languages.pseint = {
    // Comentario: // hasta el final de la línea.
    comment: {
      pattern: /\/\/.*/,
      greedy: true,
    },
    // Cadenas entre comillas dobles.
    string: {
      pattern: /"(?:[^"\\]|\\.)*"/,
      greedy: true,
    },
    keyword: [
      {
        pattern: RegExp("\\b(?:" + alternation(multiWordKeywords) + ")\\b", "i"),
        greedy: true,
      },
      {
        pattern: RegExp("\\b(?:" + alternation(singleWordKeywords) + ")\\b", "i"),
      },
    ],
    "class-name": {
      pattern: RegExp("\\b(?:" + alternation(types) + ")\\b", "i"),
    },
    boolean: /\b(?:Verdadero|Falso)\b/i,
    // La constante PI.
    constant: /\bPI\b/i,
    // Solo se resalta como función si va seguida de paréntesis, para no
    // chocar con nombres de variable que casualmente coincidan (ej. "Sen").
    function: RegExp("\\b(?:" + alternation(builtinFunctions) + ")\\b(?=\\s*\\()", "i"),
    number: /\b\d+(?:\.\d+)?\b/,
    operator: [
      // Operadores lógicos y aritméticos escritos como palabra.
      { pattern: /\b(?:Y|O|No|MOD|DIV)\b/i },
      // Asignación, comparación y operadores aritméticos simbólicos.
      { pattern: /<-|←|<>|!=|≠|<=|≤|>=|≥|\*\*|[=<>+\-*/^]/ },
    ],
    punctuation: /[()[\],;:]/,
  };
})(Prism);

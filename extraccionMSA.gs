/**
 * Función principal para extraer enlaces
 * @param {string} docUrl - URL del documento de Google Docs (ej: "https://docs.google.com/document/d/abc123...")
 * @return {Array} Array con los enlaces encontrados
 * Ejemplo de uso: extraerMSA("https://docs.google.com/document/d/tu-id-aqui/edit")
 */
function extraerMSA(docUrl) {
  try {
    if (!docUrl || docUrl.trim() === "") {
      Logger.log("URL no válida");
      return [];
    }

    const docId = extraerIdDesdeUrl(docUrl);
    if (!docId) {
      Logger.log("No se pudo extraer el ID del documento desde la URL");
      return [];
    }

    // Abrir el documento por ID
    const doc = DocumentApp.openById(docId);
    if (!doc) {
      Logger.log("No se pudo abrir el documento con ID: " + docId);
      return [];
    }

    Logger.log("Documento MSD: " + doc.getName());

    // Obtener el cuerpo del documento
    const body = doc.getBody();

    // Obtener todas las tablas del documento
    const tablas = body.getTables();
    //Logger.log("Número de tablas encontradas: " + tablas.length);

    const textoBuscar = "Modelo de Solución de Arquitectura Local";
    const enlaces = [];
    let encontrado = false;

    // Recorrer cada tabla
    for (let i = 0; i < tablas.length && !encontrado; i++) {
      const tabla = tablas[i];
      const numFilas = tabla.getNumRows();

      // Recorrer filas de la tabla
      for (let j = 0; j < numFilas && !encontrado; j++) {
        const fila = tabla.getRow(j);
        const numCeldas = fila.getNumCells();

        // Verificar si hay al menos 2 celdas
        if (numCeldas >= 2) {
          const celdaIzquierda = fila.getCell(0);
          const textoCelda = celdaIzquierda.getText().trim();

          if (textoCelda === textoBuscar) {
            encontrado = true;

            // Obtener celda derecha
            const celdaDerecha = fila.getCell(1);
            const textoCeldaDerecha = celdaDerecha.getText();
            Logger.log("Contenido de celda derecha: " + textoCeldaDerecha.substring(0, 100) + (textoCeldaDerecha.length > 100 ? "..." : ""));

            // Método 1: Usar método recursivo para extraer enlaces
            const linksEncontrados = extraerLinksRecursivamente(celdaDerecha);
            Logger.log("Enlaces encontrados con método recursivo: " + linksEncontrados.length);

            // Método 2: Buscar URLs en el contenido usando expresiones regulares
            const urlsEnRegex = buscarUrlsEnTexto(textoCeldaDerecha);
            Logger.log("Enlaces encontrados con regex: " + urlsEnRegex.length);

            // Combinar los resultados
            for (const url of linksEncontrados) {
              if (!enlaces.includes(url)) {
                enlaces.push(url);
              }
            }

            for (const url of urlsEnRegex) {
              if (!enlaces.includes(url)) {
                enlaces.push(url);
              }
            }
          }
        }
      }
    }

    // Mostrar resultados en el log
    if (enlaces.length === 0) {
      Logger.log("No se encontraron enlaces asociados a 'Modelo de Solución de Arquitectura Local'");
    } else {
      Logger.log("Enlaces encontrados (" + enlaces.length + ") para 'Modelo de Solución de Arquitectura Local'.");
      enlaces.forEach((enlace, index) => { Logger.log((index + 1) + ". " + enlace); });
    }

    /*Logger.log("\n--- RESULTADO PARA COPIAR ---");
    if (enlaces.length > 0) {
      Logger.log(enlaces.join("\n"));
    } else {
      Logger.log("No se encontraron enlaces");
    }*/

    return enlaces;

  } catch (error) {
    Logger.log("Error: " + error.message);
    return [];
  }
}

/**
 * Extrae el ID del documento desde la URL
 * @param {string} url - URL del documento de Google Docs
 * @return {string|null} ID del documento o null si no se encuentra
 */
function extraerIdDesdeUrl(url) {
  try {
    // Patrones comunes para URLs de Google Docs
    const patrones = [
      /\/d\/([a-zA-Z0-9-_]+)/,                      // /d/DOCID
      /id=([a-zA-Z0-9-_]+)/,                        // id=DOCID
      /docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/ // docs.google.com/document/d/DOCID
    ];

    for (const patron of patrones) {
      const coincidencia = url.match(patron);
      if (coincidencia && coincidencia[1]) {
        return coincidencia[1];
      }
    }

    return null;
  } catch (error) {
    Logger.log("Error al extraer ID: " + error.message);
    return null;
  }
}

/**
 * Extrae enlaces de manera recursiva de un elemento y sus hijos
 * @param {Element} elemento - Elemento de Google Docs (celda, párrafo, etc.)
 * @return {Array} Array de URLs encontrados
 */
function extraerLinksRecursivamente(elemento) {
  const enlaces = [];

  try {
    // Intentar extraer enlaces del elemento actual
    try {
      // Si es un elemento de texto, intentar obtener URLs
      if (elemento.getType && elemento.getType() === DocumentApp.ElementType.TEXT) {
        const texto = elemento.asText();
        const contenido = texto.getText();

        // Revisar cada posición del texto para encontrar enlaces
        for (let i = 0; i < contenido.length; i++) {
          try {
            const url = texto.getLinkUrl(i);
            if (url && !enlaces.includes(url)) {
              enlaces.push(url);
            }
          } catch (e) {
            // Ignorar errores en posiciones específicas
          }
        }
      }
    } catch (errorTexto) {
      // Ignorar si no es un elemento de texto
    }

    // Recursivamente buscar en los hijos del elemento
    try {
      if (typeof elemento.getNumChildren === 'function') {
        const numHijos = elemento.getNumChildren();

        for (let i = 0; i < numHijos; i++) {
          try {
            const hijo = elemento.getChild(i);
            const linksHijo = extraerLinksRecursivamente(hijo);

            // Agregar los enlaces encontrados en los hijos
            for (const link of linksHijo) {
              if (!enlaces.includes(link)) {
                enlaces.push(link);
              }
            }
          } catch (errorHijo) {
            // Ignorar errores al procesar hijos individuales
          }
        }
      }
    } catch (errorHijos) {
      // Ignorar si no se pueden procesar los hijos
    }

    return enlaces;
  } catch (error) {
    Logger.log("Error en extracción recursiva: " + error.message);
    return enlaces;
  }
}

/**
 * Busca URLs en un texto usando expresiones regulares
 * @param {string} texto - Texto donde buscar URLs
 * @return {Array} Array de URLs encontradas
 */
function buscarUrlsEnTexto(texto) {
  const enlaces = [];

  try {
    // Expresión regular para encontrar URLs en texto
    // Esta regex busca URLs que comienzan con http:// o https://
    const regex = /(https?:\/\/[^\s\)\"]+)/g;
    let coincidencia;

    // Encontrar todas las coincidencias
    while ((coincidencia = regex.exec(texto)) !== null) {
      const url = coincidencia[0];

      // Limpiar la URL de posibles caracteres no deseados al final
      const urlLimpia = url.replace(/[.,;:\)\"]$/, '');

      if (!enlaces.includes(urlLimpia)) {
        enlaces.push(urlLimpia);
      }
    }
  } catch (error) {
    Logger.log("Error buscando URLs con regex: " + error.message);
  }

  return enlaces;
}

/**
 * Función para probar el extractor con una URL específica
 */
function probarExtractor() {
  // Reemplaza esta URL con la del documento que quieres analizar
  //const urlDocumento = "https://docs.google.com/document/d/1QkRvsPHDTwd164R0COn-WYXmc2AhwW9323ykatHGBCQ/edit?tab=t.0";
  const urlDocumento = "https://docs.google.com/document/d/1DcXG_ypDq6otBR0R1_7tb5Pmm0UISpt6Xsh_CCTXmLU/edit?tab=t.0"

  // Llamar a la función principal
  const enlaces = extraerMSA(urlDocumento);

  // Los resultados aparecerán en el Log (Ver > Logs o Ctrl+Enter)
  return enlaces
}
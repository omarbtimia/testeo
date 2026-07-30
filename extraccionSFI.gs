/**
 * Función para verificar si un spreadsheet es accesible antes de intentar trabajar con él
 * @param {string} spreadsheetId - ID del Google Sheet
 * @return {boolean} - True si es accesible, False si no lo es
 */
function esSpreadsheetAccesible(spreadsheetId) {
  try {
    // Intenta acceder de manera limitada usando Drive API en lugar de SpreadsheetApp
    // Esto es menos propenso a errores de servicio
    const file = DriveApp.getFileById(spreadsheetId);
    const nombre = file.getName();

    // Si llegamos aquí, el archivo existe y es accesible
    return true;
  } catch (e) {
    // El archivo no existe o no es accesible
    return false;
  }
}

/**
 * Función principal que integra la extracción de enlaces y valores de columnas
 * @param {string} docUrl - URL del documento de Google Docs
 * @param {string} sheetName - Nombre de la hoja en Google Sheets (por defecto: "Tablas")
 * @param {string[]} headerNames - Nombres de las cabeceras a buscar
 * @return {string} Un string JSON con la respuesta estandarizada
 */
function procesarDocumentoYHoja(docUrl, sheetName = "Tablas", headerNames = ["Nombre objeto Raw", "Nombre objeto Master"]) {
  // Objeto de respuesta estandarizada
  let respuesta = {
    success: false,
    message: "",
    data: {
      enlaceEncontrado: "",
      valores: {}
    }
  };

  try {
    // Paso 1: Extraer el enlace de "Solicitud Fuentes de Información"
    const resultadoEnlaces = extraerSolicitudFuentes(docUrl);

    if (!resultadoEnlaces.success || resultadoEnlaces.data.enlaces.length === 0) {
      respuesta.message = resultadoEnlaces.message || "No se encontraron enlaces en el documento";
      return JSON.stringify(respuesta);
    }

    // Tomar el primer enlace encontrado
    const enlaceSheet = resultadoEnlaces.data.enlaces[0];
    respuesta.data.enlaceEncontrado = enlaceSheet;

    // Extraer el ID del spreadsheet de la URL
    let spreadsheetId = "";
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = enlaceSheet.match(regex);

    if (!match || !match[1]) {
      respuesta.message = "El enlace encontrado no tiene un formato válido de Google Sheets";
      return JSON.stringify(respuesta);
    }

    spreadsheetId = match[1];

    // Verificar accesibilidad ANTES de intentar abrir el spreadsheet
    if (!esSpreadsheetAccesible(spreadsheetId)) {
      respuesta.message = "Se encontró el enlace, pero el documento de Google Sheets no es accesible. Verifica permisos o si el documento existe.";
      return JSON.stringify(respuesta);
    }

    // Si llegó aquí, sabemos que el spreadsheet es accesible, ahora sí procedemos
    const resultadoColumnas = extraerValoresDesdeColumnas(enlaceSheet, sheetName, headerNames);
    const datosColumnas = JSON.parse(resultadoColumnas);

    if (!datosColumnas.success) {
      respuesta.message = "Se encontró el enlace, pero hubo un problema al extraer los valores: " + datosColumnas.message;
      return JSON.stringify(respuesta);
    }

    // Todo ha ido bien, actualizar respuesta
    respuesta.success = true;
    respuesta.message = "Proceso completado correctamente";
    respuesta.data.valores = datosColumnas.data;

  } catch (error) {
    respuesta.message = "Error en el proceso: " + error.message;
  }

  return JSON.stringify(respuesta, null, 2);
}

/**
 * Función para extraer enlaces de "Solicitud Fuentes de Información" de un documento
 * @param {string} docUrl - URL del documento de Google Docs
 * @return {Object} Objeto con la respuesta estandarizada
 */
function extraerSolicitudFuentes(docUrl) {
  try {
    if (!docUrl || docUrl.trim() === "") {
      return {
        success: false,
        message: "URL no válida",
        data: {
          enlaces: []
        }
      };
    }

    const docId = extraerIdDesdeUrl(docUrl);
    if (!docId) {
      return {
        success: false,
        message: "No se pudo extraer el ID del documento desde la URL",
        data: {
          enlaces: []
        }
      };
    }

    // Abrir el documento por ID
    const doc = DocumentApp.openById(docId);
    if (!doc) {
      return {
        success: false,
        message: "No se pudo abrir el documento con ID: " + docId,
        data: {
          enlaces: []
        }
      };
    }

    Logger.log("Documento: " + doc.getName());

    // Obtener el cuerpo del documento
    const body = doc.getBody();

    // Obtener todas las tablas del documento
    const tablas = body.getTables();

    const textoBuscar = "Solicitud Fuentes de Información";
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

            // Método 1: Usar método recursivo para extraer enlaces
            const linksEncontrados = extraerSFILinksRecursivamente(celdaDerecha);

            // Método 2: Buscar URLs en el contenido usando expresiones regulares
            const urlsEnRegex = buscarUrlsSFIEnTexto(textoCeldaDerecha);

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

    // Preparar respuesta
    if (enlaces.length === 0) {
      return {
        success: true,
        message: "No se encontraron enlaces asociados a 'Solicitud Fuentes de Información'",
        data: {
          enlaces: []
        }
      };
    } else {
      return {
        success: true,
        message: "Enlaces encontrados para 'Solicitud Fuentes de Información'",
        data: {
          enlaces: enlaces
        }
      };
    }

  } catch (error) {
    return {
      success: false,
      message: "Error: " + error.message,
      data: {
        enlaces: []
      }
    };
  }
}


/**
 * Extrae enlaces de manera recursiva de un elemento y sus hijos
 * @param {Element} elemento - Elemento de Google Docs (celda, párrafo, etc.)
 * @return {Array} Array de URLs encontrados
 */
function extraerSFILinksRecursivamente(elemento) {
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
function buscarUrlsSFIEnTexto(texto) {
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
 * Función para probar el flujo completo
 */
function probarProcesoCompleto() {
  // Reemplaza esta URL con la del documento que quieres analizar
  const urlDocumento = "https://docs.google.com/document/d/1QkRvsPHDTwd164R0COn-WYXmc2AhwW9323ykatHGBCQ/edit?pli=1&tab=t.0"//"https://docs.google.com/document/d/1KSGImAb-ZGATfOUe7lLPXE8U03q_wPNl0b-MEmffvnY/edit?tab=t.0";

  // Llamar a la función principal integrada
  const resultado = procesarDocumentoYHoja(urlDocumento);

  // Los resultados aparecerán en el Log (Ver > Logs o Ctrl+Enter)
  Logger.log(resultado);

  return resultado;
}
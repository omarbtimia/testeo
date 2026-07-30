
/**
 * Función para extraer valores de columnas específicas de una hoja de Google Sheets
 * @param {string} spreadsheetUrl - URL del Google Sheet
 * @param {string} sheetName - Nombre de la hoja (por ejemplo, "Tablas")
 * @param {string[]} headerNames - Nombres de las cabeceras a buscar
 * @return {string} Un string JSON con la respuesta estandarizada
 */
function extraerValoresDesdeColumnas(spreadsheetUrl, sheetName, headerNames) {
  // Objeto de respuesta estandarizada
  let respuesta = {
    success: false,
    message: "",
    data: {}
  };

  try {
    // Inicializar el resultado con arrays vacíos para cada cabecera
    headerNames.forEach(header => {
      respuesta.data[header] = [];
    });

    // Extraer el ID del spreadsheet de la URL
    let spreadsheetId = "";
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = spreadsheetUrl.match(regex);

    if (!match || !match[1]) {
      respuesta.message = "URL de spreadsheet inválida";
      return JSON.stringify(respuesta);
    }

    spreadsheetId = match[1];

    // Verificar accesibilidad ANTES de intentar abrir el spreadsheet
    if (!esSpreadsheetAccesible(spreadsheetId)) {
      respuesta.message = "El documento de Google Sheets no es accesible. Verifica permisos o si el documento existe.";
      return JSON.stringify(respuesta);
    }

    // Ahora sí, intentar abrir el spreadsheet por ID
    let ss = SpreadsheetApp.openById(spreadsheetId);

    // Obtener la hoja específica por nombre
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      respuesta.message = `No se encontró la hoja "${sheetName}" en el spreadsheet.`;
      return JSON.stringify(respuesta);
    }

    // Obtener todos los datos de la hoja
    let data = sheet.getDataRange().getValues();

    // Buscar las cabeceras en cualquier fila
    let columnIndexes = {};
    let headerFound = false;
    let headersFoundCount = 0;
    let headersNotFound = [];

    // Crear una lista de cabeceras que faltan por encontrar
    let pendingHeaders = [...headerNames];

    // Buscar en cada fila hasta encontrar las cabeceras
    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      // Buscar las cabeceras en esta fila
      for (let i = pendingHeaders.length - 1; i >= 0; i--) {
        let header = pendingHeaders[i];
        let colIndex = data[rowIndex].indexOf(header);
        if (colIndex !== -1) {
          columnIndexes[header] = colIndex;
          headerFound = true;
          headersFoundCount++;
          // Eliminar la cabecera encontrada de la lista pendiente
          pendingHeaders.splice(i, 1);
        }
      }

      // Si encontramos al menos una cabecera, procesar los datos
      if (headerFound && (rowIndex === data.length - 1 || Object.keys(columnIndexes).length === headerNames.length)) {
        // Extraer los valores de las columnas a partir de la siguiente fila
        for (let i = rowIndex + 1; i < data.length; i++) {
          Object.keys(columnIndexes).forEach(header => {
            let value = data[i][columnIndexes[header]];

            // Solo incluir valores que:
            // 1. No sean nulos, undefined o vacíos
            // 2. No sean "N/A"
            // 3. Sean strings que empiecen con "t_"
            if (value !== null && value !== undefined && value !== "" &&
                value !== "N/A" &&
                !(typeof value === 'string' && value.trim().toUpperCase() === "N/A")) {

                // Verificar si el valor es un string y empieza con "t_"
                if (typeof value === 'string' && value.trim().startsWith("t_")) {
                    respuesta.data[header].push(value);
                }
            }
          });
        }

        // Si no encontramos todas las cabeceras, añadir mensaje
        if (headersFoundCount < headerNames.length) {
          headersNotFound = headerNames.filter(h => !Object.keys(columnIndexes).includes(h));
          respuesta.message = `Se encontraron ${headersFoundCount} de ${headerNames.length} cabeceras. No se encontraron: ${headersNotFound.join(", ")}`;
        } else {
          respuesta.success = true;
          respuesta.message = "Datos extraídos correctamente";
        }

        break;
      }
    }

    // Si no se encontró ninguna cabecera
    if (!headerFound) {
      respuesta.message = `No se encontró ninguna de las cabeceras: ${headerNames.join(", ")}`;
      return JSON.stringify(respuesta);
    }

    // Si hemos llegado hasta aquí, al menos algunas cabeceras fueron encontradas
    respuesta.success = true;

    // Si el mensaje está vacío (porque encontramos todas las cabeceras)
    if (respuesta.message === "") {
      respuesta.message = "Datos extraídos correctamente";
    }

  } catch (error) {
    respuesta.message = `Error inesperado: ${error.message}`;
  }

  return JSON.stringify(respuesta);
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


/**
 * Procesa una lista de URLs de documentos para extraer la fecha más reciente de las tablas de control de versiones
 * @param {Array} listaUrls - Array con URLs de los documentos a procesar
 * @return {Array} Array de objetos con la información simplificada en formato JSON (solo la fecha más reciente por documento)
 */
function procesarDocumentos(listaUrls) {
  // Array para los resultados en formato JSON simple
  const resultadosJSON = [];

  // Para cada URL en la lista
  for (let i = 0; i < listaUrls.length; i++) {
    const url = listaUrls[i];
    Logger.log("Procesando documento " + (i + 1) + "/" + listaUrls.length + ": " + url);

    try {
      // Inicializar un objeto para este documento (con valores por defecto)
      const documentoResultado = {
        url: url,
        titulo: "not found",
        fecha: "not found",
        descripcion: "not found"
      };

      // Si es un documento de Google Docs
      if (url.includes("docs.google.com/document")) {
        // Intentar obtener el título del documento
        try {
          const docId = extraerIdDesdeUrl(url);
          if (docId) {
            const doc = DocumentApp.openById(docId);
            if (doc) {
              documentoResultado.titulo = doc.getName();
            }
          }
        } catch (errorTitulo) {
          Logger.log("Error obteniendo título: " + errorTitulo.message);
          // Mantener el título como "not found"
        }

        // Procesar el documento para obtener fechas
        const resultado = procesarGoogleDoc(url);

        // Si hay fechas encontradas, usar la más reciente
        if (!resultado.error && resultado.fechas.length > 0) {
          // Encontrar la fecha más reciente
          const fechaMasReciente = obtenerFechaMasReciente(resultado.fechas);

          if (fechaMasReciente) {
            // Actualizar los campos del resultado
            documentoResultado.fecha = fechaMasReciente.fecha;
            documentoResultado.descripcion = fechaMasReciente.descripcion;
          }
        } else if (resultado.error) {
          Logger.log("Error en documento: " + resultado.error);
          // Mantener los valores por defecto
        }
      }
      // Si es un PDF (no podemos procesarlo directamente con Apps Script)
      else if (url.toLowerCase().endsWith('.pdf')) {
        Logger.log("Los archivos PDF no pueden ser procesados directamente con Apps Script");
        // Mantener los valores por defecto
      }
      // Otro tipo de URL
      else {
        Logger.log("URL no reconocida como documento de Google Docs o PDF");
      }

      resultadosJSON.push(documentoResultado);

    } catch (error) {
      Logger.log("Error general procesando " + url + ": " + error.message);
      // En caso de error, agregar el documento con valores por defecto
      resultadosJSON.push({
        url: url,
        titulo: "not found",
        fecha: "not found",
        descripcion: "not found"
      });
    }
  }

  return JSON.stringify(resultadosJSON);
}

/**
 * Determina cuál de las fechas es la más reciente
 * @param {Array} fechas - Array de objetos con fechas encontradas
 * @return {Object|null} Objeto con la fecha más reciente o null si no hay fechas
 */
function obtenerFechaMasReciente(fechas) {
  if (!fechas || fechas.length === 0) {
    return null;
  }

  // Si solo hay una fecha, es la más reciente por definición
  if (fechas.length === 1) {
    return fechas[0];
  }

  // Preparar array para conversión de fechas
  const fechasConvertidas = [];

  // Primero, convertir todas las fechas de texto a objetos Date para poder compararlas
  for (let i = 0; i < fechas.length; i++) {
    const fechaTexto = fechas[i].fecha;
    const fechaObj = convertirTextoAFecha(fechaTexto);

    if (fechaObj) {
      fechasConvertidas.push({
        indice: i,
        fecha: fechaObj,
        textoOriginal: fechas[i].fecha,
        descripcion: fechas[i].descripcion
      });
    }
  }

  // Si no pudimos convertir ninguna fecha, devolver la primera
  if (fechasConvertidas.length === 0) {
    return fechas[0];
  }

  // Ordenar las fechas de más reciente a más antigua
  fechasConvertidas.sort((a, b) => b.fecha - a.fecha);

  // La primera fecha del array ordenado es la más reciente
  const indiceMasReciente = fechasConvertidas[0].indice;
  return fechas[indiceMasReciente];
}

/**
 * Convierte un texto de fecha en formato dd/MM/yyyy a objeto Date
 * @param {string} textoFecha - Texto que contiene la fecha en formato dd/MM/yyyy
 * @return {Date|null} Objeto Date o null si no se pudo convertir
 */
function convertirTextoAFecha(textoFecha) {
  try {
    // Limpiar el texto de la fecha
    textoFecha = textoFecha.trim();

    // SOLO considerar el formato dd/MM/yyyy
    const patronFecha = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;
    const coincidencia = textoFecha.match(patronFecha);

    if (coincidencia) {
      const dia = parseInt(coincidencia[1], 10);
      const mes = parseInt(coincidencia[2], 10) - 1; // Restar 1 porque en JS los meses van de 0-11
      const año = parseInt(coincidencia[3], 10);

      return new Date(año, mes, dia);
    }

    // Si no coincide con el formato esperado, devolver null
    Logger.log("No se pudo convertir la fecha (no tiene formato dd/MM/yyyy): " + textoFecha);
    return null;

  } catch (error) {
    Logger.log("Error convirtiendo texto a fecha: " + error.message);
    return null;
  }
}

/**
 * Procesa un documento de Google Docs para extraer fechas de la tabla de control de versiones
 * @param {string} url - URL del documento de Google Docs
 * @return {Object} Resultado del procesamiento
 */
function procesarGoogleDoc(url) {
  try {
    // Extraer ID del documento de la URL
    const docId = extraerIdDesdeUrl(url);
    if (!docId) {
      return {
        error: "No se pudo extraer el ID del documento desde la URL",
        fechas: []
      };
    }

    // Abrir el documento
    let doc;
    try {
      doc = DocumentApp.openById(docId);
      if (!doc) {
        return {
          error: "No se pudo abrir el documento",
          fechas: []
        };
      }
    } catch (errorDoc) {
      return {
        error: "Error abriendo documento: " + errorDoc.message,
        fechas: []
      };
    }

    Logger.log(" ****** Documento MSA: " + doc.getName());

    // Buscar la sección "Control de versiones" o "Versions control"
    const body = doc.getBody();
    const texto = body.getText();

    // Verificar si contiene la sección de control de versiones (en español o inglés)
    const tieneSeccionControlVersiones = texto.includes("Control de versiones") || texto.includes("Versions control");

    // Si no tiene la sección en el texto, no fallaremos inmediatamente - buscaremos la tabla directamente
    if (!tieneSeccionControlVersiones) {
      Logger.log("No se encontró la sección 'Control de versiones' o 'Versions control', buscando tabla por encabezados...");
    }

    // Buscar tablas que puedan contener el control de versiones
    const tablas = body.getTables();

    // Almacenar las fechas encontradas
    const fechasEncontradas = [];

    // Recorrer las tablas buscando la que tiene el formato esperado de control de versiones
    let tablaVersionesEncontrada = false;

    for (let i = 0; i < tablas.length; i++) {
      const tabla = tablas[i];
      const numFilas = tabla.getNumRows();

      // Una tabla de control de versiones debe tener al menos 2 filas (cabecera + datos)
      if (numFilas < 2) continue;

      // Verificar si esta tabla parece ser de control de versiones
      // Método 1: Verificar por los encabezados específicos
      let primeraFila;
      try {
        primeraFila = tabla.getRow(0);
        if (primeraFila.getNumCells() < 4) continue; // Necesitamos al menos 4 columnas
      } catch (errorFila) {
        Logger.log("Error accediendo a la primera fila de la tabla " + (i+1) + ": " + errorFila.message);
        continue;
      }

      let textoCabeceras;
      try {
        textoCabeceras = primeraFila.getText().toLowerCase();
      } catch (errorTexto) {
        Logger.log("Error obteniendo texto de cabeceras: " + errorTexto.message);
        continue;
      }

      // Verificar si las cabeceras coinciden con los encabezados específicos (español o inglés)
      const tieneEncabezadosEspecificos =
          (textoCabeceras.includes("versión") || textoCabeceras.includes("version")) &&
          (textoCabeceras.includes("fecha creación") || textoCabeceras.includes("fecha de creación") ||
           textoCabeceras.includes("creation date") || textoCabeceras.includes("date of creation")) &&
          (textoCabeceras.includes("responsable") || textoCabeceras.includes("responsible")) &&
          (textoCabeceras.includes("descripción") || textoCabeceras.includes("descripcion") ||
           textoCabeceras.includes("description"));

      // Método 2: Verificar por encabezados genéricos (español o inglés)
      const tieneEncabezadosGenericos =
          (textoCabeceras.includes("versión") || textoCabeceras.includes("version")) &&
          (textoCabeceras.includes("fecha") || textoCabeceras.includes("date")) &&
          (textoCabeceras.includes("descripción") || textoCabeceras.includes("descripcion") ||
           textoCabeceras.includes("description"));

      if (tieneEncabezadosEspecificos || tieneEncabezadosGenericos) {

        Logger.log("Encontrada posible tabla de control de versiones (Tabla " + (i+1) + ")");
        tablaVersionesEncontrada = true;

        // Identificar el índice de las columnas relevantes
        const numCeldas = primeraFila.getNumCells();
        let idxFecha = -1;
        let idxDescripcion = -1;

        for (let j = 0; j < numCeldas; j++) {
          try {
            const textoCelda = primeraFila.getCell(j).getText().toLowerCase().trim();
            if (textoCelda.includes("fecha creación") || textoCelda.includes("fecha de creación") ||
                textoCelda.includes("creation date") || textoCelda.includes("date of creation") ||
                textoCelda.includes("fecha") || textoCelda.includes("date")) {
              idxFecha = j;
            } else if (textoCelda.includes("descripción") || textoCelda.includes("descripcion") ||
                       textoCelda.includes("description")) {
              idxDescripcion = j;
            }
          } catch (errorCelda) {
            Logger.log("Error accediendo a celda cabecera " + j + ": " + errorCelda.message);
          }
        }

        // Si no encontramos alguna columna necesaria, pasar a la siguiente tabla
        if (idxFecha === -1 || idxDescripcion === -1) {
          Logger.log("La tabla no tiene las columnas necesarias");
          continue;
        }

        // Procesar las filas de datos
        for (let j = 1; j < numFilas; j++) { // Empezar desde 1 para saltar la cabecera
          try {
            const fila = tabla.getRow(j);
            if (fila.getNumCells() <= Math.max(idxFecha, idxDescripcion)) continue;

            let textoDescripcion;
            try {
              const celdaDescripcion = fila.getCell(idxDescripcion);
              textoDescripcion = celdaDescripcion.getText().toLowerCase().trim();
            } catch (errorDesc) {
              Logger.log("Error obteniendo descripción: " + errorDesc.message);
              continue;
            }

            // Verificar si la descripción contiene "Versión Inicial" o "Initial Version" o "Renovación de vigencia" o "Renewal"
            if (textoDescripcion.includes("versión inicial") ||
                textoDescripcion.includes("initial version") ||
                textoDescripcion.includes("renovación de vigencia") ||
                textoDescripcion.includes("renewal")) {

              let textoFecha;
              try {
                const celdaFecha = fila.getCell(idxFecha);
                textoFecha = celdaFecha.getText().trim();

                // Verificar si la fecha está en formato dd/MM/yyyy
                const esFormatoValido = /\d{1,2}\/\d{1,2}\/\d{4}/.test(textoFecha);
                if (!esFormatoValido) {
                  Logger.log("Ignorando fecha no válida (no está en formato dd/MM/yyyy): " + textoFecha);
                  continue;
                }
              } catch (errorFecha) {
                Logger.log("Error obteniendo fecha: " + errorFecha.message);
                continue;
              }

              // Agregar al resultado con la descripción original (no en minúsculas)
              fechasEncontradas.push({
                fecha: textoFecha,
                descripcion: fila.getCell(idxDescripcion).getText().trim()
              });

              Logger.log("Encontrada fecha válida (" + textoFecha + ") con descripción: " + fila.getCell(idxDescripcion).getText().trim());
            }
          } catch (errorFila) {
            Logger.log("Error procesando fila " + j + ": " + errorFila.message);
          }
        }
      }
    }

    // Si no encontramos ninguna tabla válida
    if (!tablaVersionesEncontrada) {
      return {
        error: "No se encontró ninguna tabla de control de versiones (ni por sección ni por encabezados)",
        fechas: []
      };
    }

    if (fechasEncontradas.length === 0) {
      return {
        error: "No se encontraron fechas con las descripciones buscadas en formato dd/MM/yyyy",
        fechas: []
      };
    }

    return {
      error: null,
      fechas: fechasEncontradas
    };

  } catch (error) {
    Logger.log("Error procesando Google Doc: " + error.message);
    return {
      error: error.message,
      fechas: []
    };
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
 * Función para ejecutar el procesamiento con una lista de URLs
 */
function probarProcesamiento() {
  const urls = [
    "https://docs.google.com/document/d/1zxyI6Xnnwylsk0U8M12rVDJwTXh-tCyWHg0hYEgItcc/edit?tab=t.0"
    ,"https://docs.google.com/document/d/1z9JVYVW6T0ChNNQHYBSpyDuGZu0TWYzY6I2iqEIH5Mk/edit?tab=t.0"
    ,"https://docs.google.com/document/d/1ESoRaJDkZq3syp_HDx5l6qlNtOZkZexic34lse6z-_o/edit?usp=sharing"
    ,"https://docs.google.com/document/d/1x7ZX4i-ORzzgqkb439AGSgxTBlvwalK1qUyGpfhym0E/edit?usp=sharing"
     //Agrega todas tus URLs aquí
  ];

  return procesarDocumentos(urls);
}
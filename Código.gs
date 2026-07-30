/**
 * Punto de entrada para la aplicación web
 */
function doGet() {
 return HtmlService.createTemplateFromFile('index')
   .evaluate()
   .setTitle('Procesador de Documentos MSD')
   .addMetaTag('viewport', 'width=device-width, initial-scale=1')
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
 return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Muestra la interfaz principal del procesador
 */
function mostrarInterface() {
  var html = HtmlService.createHtmlOutputFromFile('Interface')
      .setWidth(800)
      .setHeight(600)
      .setTitle('Procesador de Documentos MSD');

  DocumentApp.getUi().showModalDialog(html, 'Procesador de Documentos MSD');
}

/**
 * Obtiene el correo electrónico del usuario actual
 */
function obtenerEmailUsuario() {
  var user = Session.getEffectiveUser().getEmail();
  Logger.log(user);
  return user;
}

/**
 * Genera un ID único basado en timestamp y número aleatorio
 */
function generarID() {
  var timestamp = Math.floor(Date.now() / 1000);
  var randomNum = Math.floor(Math.random() * 900) + 100; // Número aleatorio entre 100-999
  return timestamp.toString() + randomNum.toString();
}

/**
 * Extrae el ID del documento o dibujo desde su URL
 */
function extraerDocumentID(url) {
  Logger.log("Procesando URL: " + url);

  var docMatch = url.match(/document\/d\/([a-zA-Z0-9_-]+)/);
  var drawingMatch = url.match(/drawings\/d\/([a-zA-Z0-9_-]+)/);

  if (docMatch && docMatch[1]) {
    var id = docMatch[1];
    //Logger.log("ID de documento encontrado: " + id);
    return id;
  }

  if (drawingMatch && drawingMatch[1]) {
    var id = drawingMatch[1];
    //Logger.log("ID de dibujo encontrado: " + id);
    return id;
  }

  var genericMatch = url.match(/\/(d|e)\/([a-zA-Z0-9_-]+)/);
  if (genericMatch && genericMatch[2]) {
    var id = genericMatch[2];
    Logger.log("ID encontrado (método genérico): " + id);
    return id;
  }

  throw new Error("No se pudo encontrar un ID válido en la URL: " + url);
}

/**
 * Busca diagramas en el documento de Google
 */
function buscarDiagramas(docId) {
  try {
    Logger.log("Iniciando búsqueda de diagramas en documento: " + docId);

    var doc = DocumentApp.openById(docId);
    if (!doc) {
      throw new Error("No se pudo abrir el documento");
    }

    var diagramas = [];
    var body = doc.getBody();

    // Función para procesar elementos de texto y buscar enlaces a diagramas
    function procesarTexto(elemento) {
      if (!elemento) return;

      try {
        var text = elemento.asText();
        var fullText = text.getText();

        // Recorrer todo el texto buscando enlaces a drawings
        var lastUrl = null;
        for (var i = 0; i < fullText.length; i++) {
          var url = text.getLinkUrl(i);

          if (url && url !== lastUrl) {
            lastUrl = url;

            // Filtro: solo drawings + "Diagrama"
            if (url.indexOf("docs.google.com/drawings") >= 0 && fullText.indexOf("Diagrama") >= 0) {

              // Extraer nombre: buscar la ocurrencia de "Diagrama" MÁS CERCANA hacia atrás desde i
              var closestDiagramaPos = -1;
              var searchFrom = 0;
              while (true) {
                var pos = fullText.indexOf("Diagrama", searchFrom);
                if (pos < 0 || pos > i) break;
                closestDiagramaPos = pos;
                searchFrom = pos + 1;
              }
              if (closestDiagramaPos < 0) {
                closestDiagramaPos = fullText.indexOf("Diagrama", i);
              }

              var start = Math.max(0, closestDiagramaPos);
              var end = Math.min(fullText.length, start + 100);
              var context = fullText.substring(start, end);

              var nombre = "";
              var match = context.match(/Diagrama[^.,;\r\n"<>]+/);
              if (match) {
                nombre = match[0].trim();
              } else {
                nombre = "Diagrama " + (diagramas.length + 1);
              }

              var exists = false;
              for (var j = 0; j < diagramas.length; j++) {
                if (diagramas[j].url === url) {
                  exists = true;
                  break;
                }
              }

              if (!exists) {
                diagramas.push({
                  nombre: nombre,
                  url: url
                });
                Logger.log("Agregado diagrama: " + nombre + " | URL: " + url);
              }

              // Saltar al siguiente enlace potencial
              i = text.getText().indexOf("Diagrama", i + 1);
              if (i < 0) break;
            }
          } else if (!url) {
            lastUrl = null;
          }
        }
      } catch (e) {
        Logger.log("Error al procesar texto: " + e.toString());
      }
    }

    // Función recursiva para procesar todos los elementos del documento
    function procesarElemento(elemento) {
      if (!elemento) return;

      if (elemento.getType() === DocumentApp.ElementType.TEXT) {
        procesarTexto(elemento);
      }
      else if (elemento.getType() === DocumentApp.ElementType.PARAGRAPH) {
        procesarTexto(elemento.asText());
      }
      else if (elemento.getType() === DocumentApp.ElementType.TABLE) {
        var table = elemento.asTable();
        for (var i = 0; i < table.getNumRows(); i++) {
          var row = table.getRow(i);
          for (var j = 0; j < row.getNumCells(); j++) {
            var cell = row.getCell(j);
            // Procesar el contenido de la celda
            for (var k = 0; k < cell.getNumChildren(); k++) {
              procesarElemento(cell.getChild(k));
            }
          }
        }
      }
      else if (elemento.getType() === DocumentApp.ElementType.LIST_ITEM) {
        procesarTexto(elemento.asListItem().asText());
      }

      // Procesar hijos si los tiene
      if (elemento.getNumChildren) {
        var numChildren = 0;
        try {
          numChildren = elemento.getNumChildren();
        } catch (e) {
          // Algunos elementos no tienen getNumChildren()
        }

        for (var i = 0; i < numChildren; i++) {
          procesarElemento(elemento.getChild(i));
        }
      }
    }
    procesarElemento(body);

    Logger.log("Búsqueda completada. Diagramas encontrados: " + diagramas.length);
    return diagramas;
  } catch (e) {
    Logger.log("Error en buscarDiagramas: " + e.toString());
    throw e;
  }
}

/**
 * Extrae el nombre más preciso para un diagrama
 */
function extraerNombreDiagrama(texto) {
  // Intentar extraer un nombre completo del diagrama
  var regexNombre = /Diagrama[^.,:;"\n\r]+/i;
  var match = texto.match(regexNombre);
  if (match) {
    return match[0].trim();
  }
  return "Diagrama sin nombre";
}

/**
 * Determina el nombre estandarizado del diagrama según su contenido
 */
function obtenerNombreEstandar(nombreOriginal, idFiles) {
  var nombreLower = nombreOriginal.toLowerCase();

  if (nombreLower.includes("acl")) {
    return "diagrama_acls" + idFiles;
  } else if (nombreLower.includes("informacional")) {
    return "diagrama_informacional" + idFiles;
  } else if (nombreLower.includes("alto")) {
    return "diagrama_alto_nivel" + idFiles;
  }

  return "diagrama_generico" + idFiles;
}

/**
 * Descarga un diagrama de Google Drawing como PDF y lo guarda en Drive
 */
function descargarDiagrama(url, nombreEstandar) {
  try {
    var diagramId = extraerDocumentID(url);
    //Logger.log("ID extraido"+ diagramId)

    var carpeta = obtenerCarpetaTemporal();
    //Logger.log("Carpeta temporal asignada")

    // Exportar el diagrama como PDF y guardarlo en Drive
    var diagramaFile = DriveApp.getFileById(diagramId);
    var pdfBlob = diagramaFile.getBlob().getAs('application/pdf');
    pdfBlob.setName(nombreEstandar + ".pdf");

    var archivo = carpeta.createFile(pdfBlob);
    return archivo.getId();
  } catch (e) {
    Logger.log("Error al descargar diagrama: " + e.toString());
    throw e;
  }
}

/**
 * Descarga el documento MSD como DOCX usando exportación directa
 */
function descargarDocumentoMSD(docId, idFiles) {
  try {
    var docFile = DriveApp.getFileById(docId);
    var mimeType = docFile.getMimeType();
    Logger.log("Tipo MIME del archivo: " + mimeType);

    var carpeta = obtenerCarpetaTemporal();
    var nombreArchivo = "archivoMSD_" + idFiles + ".docx";

    // Verificar si es un documento de Google Docs
    if (mimeType === "application/vnd.google-apps.document") {
      Logger.log("Exportando documento de Google Docs a DOCX mediante URL directa");

      // Crear una URL de exportación directa
      var exportUrl = "https://docs.google.com/document/d/" + docId + "/export?format=docx";

      // Utilizar UrlFetchApp para descargar el archivo
      var response = UrlFetchApp.fetch(exportUrl, {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        var docxBlob = response.getBlob().setName(nombreArchivo);
        var archivo = carpeta.createFile(docxBlob);
        Logger.log("Documento exportado con éxito. ID: " + archivo.getId());
        return archivo.getId();
      } else {
        throw new Error("Error al descargar documento: Código de respuesta " + response.getResponseCode());
      }
    } else {
      throw new Error("El archivo seleccionado no es un documento de Google Docs. Por favor, asegúrate de usar una URL que apunte a un documento de Google Docs.");
    }
  } catch (e) {
    Logger.log("Error al descargar documento MSD: " + e.toString());
    throw e;
  }
}

/**
 * Obtiene o crea una carpeta temporal para los archivos
 */
function obtenerCarpetaTemporal() {
  var carpetas = DriveApp.getFoldersByName("MSD_Temp");

  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return DriveApp.createFolder("MSD_Temp");
  }
}

/**
 * Envía los archivos al API para procesamiento en formato Base64
 */
function enviarArchivosAPI(msdFileId, diagramasIds, apiUrl, originalFilename, msaData, sfiData) {
  try {
    var userEmail = Session.getEffectiveUser().getEmail();

    // Obtener los archivos de Drive
    var msdFile = DriveApp.getFileById(msdFileId);
    var diagramasFiles = {};

    for (var tipo in diagramasIds) {
      diagramasFiles[tipo] = DriveApp.getFileById(diagramasIds[tipo]);
    }

    // Validar que el diagrama de ACLs exista antes de intentar codificarlo
    if (!diagramasFiles.acls) {
      throw new Error("No se encontró el diagrama de ACLs en el documento. " +
        "Verifique que el nombre del diagrama contenga la palabra 'ACL'. " +
        "Keys disponibles: [" + Object.keys(diagramasFiles).join(", ") + "]");
    }

    var codeInfo = "";
    try {
      if (diagramasFiles.informacional) {
        codeInfo = Utilities.base64Encode(diagramasFiles.informacional.getBlob().getBytes());
      } else {
        Logger.log("Advertencia: No se encontró diagrama informacional. Se enviará vacío.");
      }
    } catch (e) {
      Logger.log("Error al procesar archivo informacional: " + e.toString());
    }

    var payload = JSON.stringify({
      accion: 'validate',
      usuario: userEmail,
      originalFilename: originalFilename,
      msaData: msaData,
      sfiData: sfiData,
      file: Utilities.base64Encode(msdFile.getBlob().getBytes()),
      file2: Utilities.base64Encode(diagramasFiles.acls.getBlob().getBytes()),
      file3: codeInfo,
      flag: 'msd'
    });

    var options = {
      method: "post",
      contentType: "application/json",
      payload: payload,
      muteHttpExceptions: true
    };

    // Realizar la petición
    Logger.log("Enviando solicitud a: " + apiUrl);
    var response = UrlFetchApp.fetch(apiUrl, options);

    if (response.getResponseCode() == 200) {
      var responseData = JSON.parse(response.getContentText());
      Logger.log("Respuesta completada desde el API");
      return responseData;
    } else {
      Logger.log("Error en respuesta del API: " + response.getResponseCode());
      throw new Error("Error en la respuesta del API: " + response.getResponseCode() + " - " + response.getContentText());
    }
  } catch (e) {
    Logger.log("Error al enviar archivos al API: " + e.toString());
    throw e;
  }
}

/**
 * Función principal modificada con mejor manejo de errores SFI
 */
function procesarDocumentoMSD(docUrl, apiUrl = 'https://j8lvcj7gs4.execute-api.us-east-1.amazonaws.com/prod/talia') {
  try {
    var apiUrl2= obtenerEndpoint();
    Logger.log("Operación realizada por: " + Session.getEffectiveUser().getEmail());

    // Generar ID único para esta ejecución
    var idFiles = generarID();
    var resultados = {};

    // Extraer ID del documento
    var docId = extraerDocumentID(docUrl);
    var doc = DocumentApp.openById(docId);
    var originalFilename = doc.getName();

    // ======= INFORMACIÓN BÁSICA DEL DOCUMENTO =======
    resultados.documentId = docId;
    resultados.nombreDocumento = originalFilename;
    resultados.urlCompleta = docUrl;

    // Buscar diagramas
    var diagramas = buscarDiagramas(docId);
    resultados.diagramasEncontrados = diagramas.length;
    // NO incluir la lista de diagramas en la respuesta

    if (diagramas.length === 0) {
      resultados.estado = "warning";
      resultados.mensaje = "No se encontraron diagramas en el documento.";
      return resultados;
    }

    // ======= PROCESAMIENTO DE DOCUMENTOS MSA =======
    try {
      var urls = extraerMSA(docUrl);
      Logger.log("URLs MSA extraídas: " + urls.length);

      if (urls && urls.length > 0) {
        var msaDataString = procesarDocumentos(urls);
        var msaData = JSON.parse(msaDataString);

        // Agregar información de MSA al resultado (sin fechas)
        resultados.documentosMSA = msaData.map(function(doc) {
          return {
            url: doc.url,
            titulo: doc.titulo,
            fecha: doc.fecha,
            descripcion: doc.descripcion
          };
        });

        Logger.log("Documentos MSA procesados: " + msaData.length);
      } else {
        resultados.documentosMSA = [];
        Logger.log("No se encontraron URLs de MSA");
      }
    } catch (errorMSA) {
      Logger.log("Error procesando MSA: " + errorMSA.message);
      resultados.documentosMSA = [];
    }

    Logger.log(resultados.documentosMSA);

    // ======= PROCESAMIENTO DE DOCUMENTOS SFI CON MEJOR MANEJO DE ERRORES =======
    try {
      var sfiDataString = procesarDocumentoYHoja(docUrl);
      var sfiData = JSON.parse(sfiDataString);

      Logger.log("Resultado SFI: " + JSON.stringify(sfiData));

      if (sfiData.success && sfiData.data.enlaceEncontrado) {
        // SFI encontrado y procesado correctamente
        try {
          var sfiDocId = extraerDocumentIDFromUrl(sfiData.data.enlaceEncontrado);
          if (sfiDocId) {
            var sfiFile = DriveApp.getFileById(sfiDocId);
            sfiData.nombreDocumento = sfiFile.getName();
          }
        } catch (errorNombreSFI) {
          Logger.log("Error obteniendo nombre del SFI: " + errorNombreSFI.message);
        }

        resultados.documentoSFI = sfiData.data;
        if (sfiData.nombreDocumento) {
          resultados.documentoSFI.nombreDocumento = sfiData.nombreDocumento;
        }

        Logger.log("SFI procesado correctamente");
      } else {
        // Error específico en SFI
        resultados.documentoSFI = null;

        if (sfiData.message) {
          if (sfiData.message.includes("no es accesible")) {
            resultados.errorSFI = "Se encontró enlace SFI pero el documento no es accesible (verificar permisos)";
          } else if (sfiData.message.includes("No se encontraron enlaces")) {
            resultados.errorSFI = "No se encontraron enlaces a 'Solicitud Fuentes de Información' en el MSD";
          } else if (sfiData.message.includes("formato válido")) {
            resultados.errorSFI = "Se encontró enlace SFI pero no tiene formato válido de Google Sheets";
          } else {
            resultados.errorSFI = sfiData.message;
          }
        } else {
          resultados.errorSFI = "Error desconocido procesando SFI";
        }

        Logger.log("Error SFI: " + resultados.errorSFI);
      }

    } catch (errorSFI) {
      Logger.log("Error crítico procesando SFI: " + errorSFI.message);
      resultados.documentoSFI = null;
      resultados.errorSFI = "Error crítico al procesar SFI: " + errorSFI.message;
    }

    // ======= PROCESAMIENTO DE DIAGRAMAS (resto del código original) =======
    var diagramasIds = {};

    for (var i = 0; i < diagramas.length; i++) {
      var diagrama = diagramas[i];
      var nombreEstandar = obtenerNombreEstandar(diagrama.nombre, idFiles);

      // Descargar el diagrama
      var diagramaId = descargarDiagrama(diagrama.url, nombreEstandar);

      // Clasificar por tipo
      if (nombreEstandar.includes("acls")) {
        diagramasIds.acls = diagramaId;
      } else if (nombreEstandar.includes("informacional")) {
        diagramasIds.informacional = diagramaId;
      } else if (nombreEstandar.includes("alto_nivel")) {
        diagramasIds.alto_nivel = diagramaId;
      } else {
        Logger.log("Advertencia: diagrama no clasificado: " + diagrama.nombre);
      }
    }

    try {
      // Descargar documento MSD
      var archivo_msd = descargarDocumentoMSD(docId, idFiles);
    } catch (e) {
      if (e.message.includes("no es un documento de Google Docs")) {
        // Error específico para este caso
        resultados.estado = "error";
        resultados.mensaje = "El archivo debe ser un documento de Google Docs. Por favor, asegúrate de usar una URL de un documento de Google Docs (no PDF, Word externo u otro formato).";
        return resultados;
      } else {
        // Otros errores de descarga
        throw e;
      }
    }

    // Preparar datos para API (MSA sin fechas, SFI como está)
    var msaDataForAPI = resultados.documentosMSA || [];
    var sfiDataForAPI = resultados.documentoSFI || {};

    // Enviar al API
    var apiRespuesta = enviarArchivosAPI(archivo_msd, diagramasIds, apiUrl2, originalFilename, msaDataForAPI, sfiDataForAPI);

    // Procesar resultados con la estructura de respuesta
    if (apiRespuesta) {
      // Verificar si tenemos una estructura anidada con 'body'
      if (apiRespuesta.body && typeof apiRespuesta.body === 'string') {
        try {
          // Intentar parsear el contenido del body
          var bodyData = JSON.parse(apiRespuesta.body);

          if (bodyData.datos_extraidos) {
            resultados.datosExtraidos = bodyData.datos_extraidos;
            resultados.estado = "success";
            resultados.mensaje = "Procesamiento completado exitosamente.";
            return resultados;
          }
        } catch (parseError) {
          Logger.log("Error al parsear el body de la respuesta: " + parseError.toString());
        }
      }

      // Verificar la estructura antigua
      if (apiRespuesta.datos_extraidos) {
        resultados.datosExtraidos = apiRespuesta.datos_extraidos;
        resultados.estado = "success";
        resultados.mensaje = "Procesamiento completado exitosamente.";
        return resultados;
      }

      // Si llegamos aquí, no pudimos encontrar los datos extraídos en ningún formato
      resultados.estado = "error";
      resultados.mensaje = "Error en el procesamiento con API: respuesta en formato inesperado.";
      resultados.respuestaAPI = apiRespuesta; // Incluir la respuesta para depuración
    } else {
      resultados.estado = "error";
      resultados.mensaje = "Error en el procesamiento con API: no se recibió respuesta.";
    }

    return resultados;
  } catch (e) {
    Logger.log("Error en procesarDocumentoMSD: " + e.toString());
    return {
      estado: "error",
      mensaje: "Error durante el procesamiento: " + e.message
    };
  }
}

/**
 * Función auxiliar para extraer ID de documento de cualquier URL de Google
 */
function extraerDocumentIDFromUrl(url) {
  try {
    var patterns = [
      /\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+)/,
      /\/([a-zA-Z0-9_-]+)\/edit/
    ];

    for (var i = 0; i < patterns.length; i++) {
      var match = url.match(patterns[i]);
      if (match && match[1]) {
        return match[1];
      }
    }

    return null;
  } catch (error) {
    Logger.log("Error extrayendo ID de URL: " + error.message);
    return null;
  }
}

/**
 * Función para limpiar los archivos temporales
 */
function limpiarArchivosTemporales() {
  try {
    var carpetas = DriveApp.getFoldersByName("MSD_Temp");

    if (carpetas.hasNext()) {
      var carpeta = carpetas.next();
      var archivos = carpeta.getFiles();

      while (archivos.hasNext()) {
        var archivo = archivos.next();
        archivo.setTrashed(true);
      }

      return { estado: "success", mensaje: "Archivos temporales eliminados correctamente." };
    } else {
      return { estado: "warning", mensaje: "No se encontró la carpeta temporal." };
    }
  } catch (e) {
    Logger.log("Error al limpiar archivos: " + e.toString());
    return { estado: "error", mensaje: "Error al limpiar archivos: " + e.message };
  }
}

function obtenerPropiedad() {
  // Obtiene el servicio de propiedades del script
  var scriptProperties = PropertiesService.getScriptProperties();

  // Lee el valor de una clave específica (reemplaza 'MI_CLAVE' por la tuya)
  return scriptProperties.getProperty('AMBIENTE');
}

function obtenerEndpoint() {
  // Obtiene el servicio de propiedades del script
  var scriptProperties = PropertiesService.getScriptProperties();

  // Lee el valor de una clave específica (reemplaza 'MI_CLAVE' por la tuya)
  return scriptProperties.getProperty('ENDPOINT');
}
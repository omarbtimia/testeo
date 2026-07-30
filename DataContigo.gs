/**
 * Script para extraer datos específicos de Google Sheets y enviarlos a un API
 *
 * Este script:
 * 1. Abre el archivo específico de Google Sheets
 * 2. Extrae solo las columnas "UUAA RAW", "UUAA MASTER" y "PROYECTO HABILITADOR"
 * 3. Convierte los datos a formato CSV
 * 4. Envía el CSV a un endpoint/API externo mediante una petición POST
 */

function exportDataContigo() {
  try {
    // ID del archivo de Google Sheets
    const spreadsheetId = '1Ikgi2MdbDUiiYWuGaAuCjT5Cy7eUQFpLHF9bW2gLUCM';

    // ID de la hoja (gid en la URL)
    const sheetId = '1337268916';

    // Abrir el archivo
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // Obtener la hoja específica por su ID
    // Obtenemos todas las hojas y buscamos la que coincida con el ID
    const sheets = spreadsheet.getSheets();
    let targetSheet = null;

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId().toString() === sheetId) {
        targetSheet = sheets[i];
        break;
      }
    }

    if (!targetSheet) {
      throw new Error('Hoja no encontrada con el ID especificado');
    }

    // Obtener todos los datos de la hoja
    const allData = targetSheet.getDataRange().getValues();

    // Verificar que hay datos
    if (allData.length === 0) {
      throw new Error('No se encontraron datos en la hoja');
    }

    // Obtener los encabezados (primera fila)
    const headers = allData[0];

    // Encontrar los índices de las columnas que necesitamos
    const uuaaRawIndex = headers.indexOf('UUAA RAW');
    const uuaaMasterIndex = headers.indexOf('UUAA MASTER');
    const proyectoHabilitadorIndex = headers.indexOf('PROYECTO HABILITADOR');

    // Verificar que todas las columnas existen
    if (uuaaRawIndex === -1 || uuaaMasterIndex === -1 || proyectoHabilitadorIndex === -1) {
      throw new Error('Una o más columnas requeridas no se encontraron');
    }

    // Crear un array para los datos filtrados
    const filteredData = [];

    // Añadir los encabezados al array filtrado
    filteredData.push(['UUAA RAW', 'UUAA MASTER', 'PROYECTO HABILITADOR']);

    // Filtrar los datos para incluir solo las columnas requeridas
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      filteredData.push([
        row[uuaaRawIndex],
        row[uuaaMasterIndex],
        row[proyectoHabilitadorIndex]
      ]);
    }

    // Convertir los datos filtrados a formato CSV
    let csvContent = '';

    // Procesar cada fila y convertirla a formato CSV
    for (let i = 0; i < filteredData.length; i++) {
      // Escapar comillas dobles y encerrar valores con comillas si contienen comas o saltos de línea
      const processedRow = filteredData[i].map(value => {
        // Convertir valores nulos o indefinidos a cadenas vacías
        if (value === null || value === undefined) {
          return '""';
        }

        // Convertir el valor a string
        const stringValue = String(value);

        // Si el valor contiene comillas, comas o saltos de línea, escapar las comillas y encerrar en comillas
        if (stringValue.includes('"') || stringValue.includes(',') || stringValue.includes('\n')) {
          return '"' + stringValue.replace(/"/g, '""') + '"';
        }

        // Si es una cadena vacía o solo tiene espacios, encerrar en comillas
        if (stringValue.trim() === '') {
          return '""';
        }

        return stringValue;
      });

      // Unir los valores de la fila con comas y añadir un salto de línea
      csvContent += processedRow.join(',') + '\n';
    }

    // URL del endpoint al que enviaremos los datos
    //const apiUrl = 'https://ynh6es1s00.execute-api.us-east-1.amazonaws.com/dev/TAL-IA-DEV';
    const apiUrl = obtenerEndpoint();
    var userEmail = Session.getEffectiveUser().getEmail();

    var payload = JSON.stringify({
      usuario: userEmail,
      originalFilename: '',
      msaData: '',
      sfiData: '',
      file: csvContent,
      file2: '',
      file3: '',
      flag: 'datacontigo',
      accion :  'datacontigo'
    });

    // Configuración de la petición
    var options = {
      method: "post",
      contentType: "application/json",
      payload: payload,
      muteHttpExceptions: true
    };

    // Enviar la petición al API
    const response = UrlFetchApp.fetch(apiUrl, options);

    // Verificar la respuesta
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      Logger.log('Datos enviados correctamente al API');
      Logger.log(response.getContentText());
      return { success: true, message: 'Datos enviados correctamente' };
    } else {
      Logger.log('Error al enviar datos al API: ' + response.getResponseCode());
      Logger.log(response.getContentText());
      return { success: false, message: 'Error en la respuesta del API: ' + response.getResponseCode() };
    }

  } catch (error) {
    Logger.log('Error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Función para configurar un disparador que ejecute el script de forma periódica
 * (opcional, para automatizar el proceso)
 */
function createTimeTrigger() {
  // Eliminar triggers existentes para evitar duplicados
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // Crear un nuevo trigger para que se ejecute diariamente
  ScriptApp.newTrigger('exportDataContigo')
    .timeBased()
    .everyDays(1)  // Ejecutar diariamente, ajusta según necesites
    .atHour(1)     // A la 1 AM, ajusta según necesites
    .create();

  Logger.log('Trigger creado correctamente para ejecutar diariamente');
}
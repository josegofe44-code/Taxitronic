/**
 * TAXITRONIC · Apps Script de sincronización
 * -------------------------------------------
 * Recibe por POST los datos del día (de un usuario/licencia concreto) desde
 * la app y los escribe (o actualiza) en la hoja "Datos diarios". También
 * genera un PDF de esa hoja diaria y lo guarda en una carpeta PROPIA de
 * cada licencia, dentro de "PDFs", dentro de la carpeta de Drive donde
 * está esta hoja de cálculo:
 *   taxitronic/PDFs/<licencia>/recaudacion_<licencia>_<fecha>.pdf
 *
 * INSTALACIÓN (resumen, ver README.md del repo para el detalle):
 * 1. Abre tu hoja "taxitronic" en Google Sheets.
 * 2. Extensiones → Apps Script.
 * 3. Pega este archivo como Code.gs y el contenido de appsscript.json en el
 *    manifiesto (icono de engranaje → "Mostrar archivo de manifiesto").
 * 4. Implementar → Nueva implementación → tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que te da y pégala en la constante SYNC_URL de index.html.
 *
 * IMPORTANTE: si ya tenías una implementación anterior, tras pegar este
 * código nuevo tienes que crear una IMPLEMENTACIÓN NUEVA (Implementar →
 * Gestionar implementaciones → lápiz → Versión "Nueva versión" →
 * Implementar) para que los cambios se apliquen, y volver a autorizar los
 * permisos (pide acceso a Drive y a Docs, además de a la hoja, para poder
 * crear los PDF).
 */

var SHEET_NAME = 'Datos diarios';
var LAST_PREPARED_ROW = 400; // filas con fórmulas ya preparadas en el Excel original
var DATE_COL = 1;            // A: Fecha
var FIRST_INCOME_COL = 2;    // B: Efectivo ... G: Emisora (6 columnas)
var FIRST_EXPENSE_COL = 9;   // I: Combustible, J: Otros
var CARRERAS_COL = 13;       // M
var ENVIADA_COL = 14;        // N
var LICENCIA_COL = 15;       // O  ← columna nueva, añádela tú a mano (ver README)
var REFUERZO_COL = 16;       // P  ← columna nueva, añádela tú a mano (ver README)

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.postData) {
      return jsonResponse_({ ok: false, error: 'Sin cuerpo en la petición.' });
    }
    var data = JSON.parse(e.postData.contents);
    if (!data.date) {
      return jsonResponse_({ ok: false, error: 'Falta la fecha (date).' });
    }
    var licencia = data.licencia ? String(data.licencia) : '';

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse_({ ok: false, error: 'No existe la hoja "' + SHEET_NAME + '".' });
    }

    var tz = Session.getScriptTimeZone();
    var targetRow = -1;
    var firstEmptyRow = -1;
    var lastRow = Math.max(sheet.getLastRow(), LAST_PREPARED_ROW);

    for (var r = 2; r <= lastRow; r++) {
      var cellVal = sheet.getRange(r, DATE_COL).getValue();
      if (cellVal === '' || cellVal === null) {
        if (firstEmptyRow === -1) firstEmptyRow = r;
        continue;
      }
      var cellStr = (cellVal instanceof Date)
        ? Utilities.formatDate(cellVal, tz, 'yyyy-MM-dd')
        : String(cellVal);
      var cellLicencia = String(sheet.getRange(r, LICENCIA_COL).getValue() || '');
      // Misma fecha Y misma licencia → es una actualización de esa hoja.
      if (cellStr === data.date && cellLicencia === licencia) { targetRow = r; break; }
    }

    if (targetRow === -1) targetRow = firstEmptyRow;
    if (targetRow === -1) {
      return jsonResponse_({ ok: false, error: 'No quedan filas libres; amplía el rango de fórmulas en el Excel.' });
    }

    sheet.getRange(targetRow, DATE_COL).setValue(new Date(data.date + 'T00:00:00'));
    sheet.getRange(targetRow, FIRST_INCOME_COL, 1, 6).setValues([[
      num_(data.efectivo), num_(data.tarjeta), num_(data.uber),
      num_(data.freenow), num_(data.bolt), num_(data.emisora)
    ]]);
    sheet.getRange(targetRow, FIRST_EXPENSE_COL, 1, 2).setValues([[
      num_(data.combustible), num_(data.otros)
    ]]);
    sheet.getRange(targetRow, CARRERAS_COL).setValue(num_(data.carreras));
    sheet.getRange(targetRow, ENVIADA_COL).setValue(data.enviada ? 'Sí' : 'No');
    sheet.getRange(targetRow, LICENCIA_COL).setValue(licencia);
    sheet.getRange(targetRow, REFUERZO_COL).setValue(data.refuerzo ? 'Sí' : 'No');
    // Las columnas H (Facturado), K (Gastos) y L (Neto) ya tienen fórmulas
    // preparadas en la plantilla y se recalculan solas — no se tocan aquí.

    var pdfUrl = null;
    var pdfError = null;
    if (data.enviada) {
      try {
        pdfUrl = crearPdfDiario_(data, licencia);
      } catch (pdfErr) {
        // Si falla el PDF no queremos que falle todo el guardado de datos.
        pdfError = pdfErr.message;
        console.error('Error generando PDF: ' + pdfErr.message);
      }
    }

    return jsonResponse_({ ok: true, row: targetRow, pdfUrl: pdfUrl, pdfError: pdfError });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Crea un PDF de una página con el desglose del día y lo guarda en la
 * carpeta propia de esa licencia (taxitronic/PDFs/<licencia>/).
 * Devuelve la URL del PDF.
 */
function crearPdfDiario_(data, licencia) {
  var folder;
  try {
    folder = getOrCreateUserPdfFolder_(licencia);
  } catch (folderErr) {
    throw new Error('No se pudo crear/acceder a la carpeta de PDFs: ' + folderErr.message);
  }

  var nombre = 'recaudacion_' + (licencia || 'sinlicencia') + '_' + data.date;
  var doc, docFile, pdfBlob, pdfFile;

  try {
    doc = DocumentApp.create(nombre);
    var body = doc.getBody();
    body.appendParagraph('TAXITRONIC · Hoja diaria').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Licencia: ' + (licencia || '-'));
    body.appendParagraph('Fecha: ' + data.date);
    body.appendParagraph('Refuerzo: ' + (data.refuerzo ? 'Sí' : 'No'));
    body.appendParagraph('');

    var filas = [
      ['Concepto', 'Importe (€)'],
      ['Efectivo', eur_(data.efectivo)],
      ['Tarjeta', eur_(data.tarjeta)],
      ['Uber', eur_(data.uber)],
      ['Freenow', eur_(data.freenow)],
      ['Bolt', eur_(data.bolt)],
      ['Emisora', eur_(data.emisora)],
      ['TOTAL FACTURADO', eur_(num_(data.efectivo) + num_(data.tarjeta) + num_(data.uber) + num_(data.freenow) + num_(data.bolt) + num_(data.emisora))],
      ['Combustible', eur_(data.combustible)],
      ['Otros gastos', eur_(data.otros)],
      ['TOTAL GASTOS', eur_(num_(data.combustible) + num_(data.otros))],
      ['Carreras', String(num_(data.carreras))]
    ];
    body.appendTable(filas);
    doc.saveAndClose();
  } catch (docErr) {
    throw new Error('No se pudo crear el documento temporal (permiso de Documentos): ' + docErr.message);
  }

  try {
    docFile = DriveApp.getFileById(doc.getId());
    pdfBlob = docFile.getAs('application/pdf');
    pdfFile = folder.createFile(pdfBlob).setName(nombre + '.pdf');
    docFile.setTrashed(true); // borra el Google Doc temporal, solo queremos el PDF
  } catch (pdfConvErr) {
    throw new Error('No se pudo convertir/guardar el PDF (permiso de Drive): ' + pdfConvErr.message);
  }

  return pdfFile.getUrl();
}

/**
 * Devuelve (creándola si hace falta) la carpeta PDFs/<licencia> dentro de
 * la carpeta que contiene esta hoja de cálculo.
 */
function getOrCreateUserPdfFolder_(licencia) {
  var ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = ssFile.getParents();
  var baseFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var pdfFolders = baseFolder.getFoldersByName('PDFs');
  var pdfFolder = pdfFolders.hasNext() ? pdfFolders.next() : baseFolder.createFolder('PDFs');

  var nombreUsuario = licencia || 'sinlicencia';
  var userFolders = pdfFolder.getFoldersByName(nombreUsuario);
  return userFolders.hasNext() ? userFolders.next() : pdfFolder.createFolder(nombreUsuario);
}

function doGet(e) {
  return jsonResponse_({ ok: true, message: 'Endpoint de sincronización TAXITRONIC activo.' });
}

/**
 * FUNCIÓN DE PRUEBA — ejecútala tú a mano desde este editor (▶ Ejecutar,
 * con "testCrearPdf" seleccionado arriba) para forzar la pantalla de
 * autorización de Drive/Documentos. Los llamantes anónimos de la app web
 * (o sea, tu propia app) NUNCA pueden conceder ese permiso por ti — solo
 * tú, ejecutando algo manualmente en este editor, una vez.
 * Después mira en Drive → taxitronic → PDFs → TEST si apareció el PDF.
 */
function testCrearPdf() {
  var datosPrueba = {
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    efectivo: 10, tarjeta: 0, uber: 0, freenow: 0, bolt: 0, emisora: 0,
    combustible: 0, otros: 0, carreras: 1, refuerzo: false
  };
  var url = crearPdfDiario_(datosPrueba, 'TEST');
  Logger.log('PDF de prueba creado: ' + url);
}

function num_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function eur_(v) {
  return num_(v).toFixed(2) + ' €';
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

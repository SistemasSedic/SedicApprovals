// Llamar las columnas 

const ACTION = 'Action';   // Se debe agregar manual una vez creado el formulario 
const STATUS = 'Status';   // Se debe agregar manual una vez
const NOMBRE_APROBADOR = 'Nombre aprobador';
const EMAIL = 'Email';
const NUMERO_DEL_CONTRATO= 'Numero del contrato'; 
const ESTADO = 'Estado';
const LINK_CONTRATO = 'Link contrato';
const Link_admon_contratos = 'Link admon contratos';
const ESTADO_DEL_PROCESO = 'Estado del proceso';
const ANEXOS = 'Anexos';

const BASE_SHEET = 'Responses';
const TEMPLATES_SHEET = 'Templates';


/**
 * Crea el menu en la hoja de calculo para enviar las solicitudes
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Send email')
      .addItem('EN REVISIÓN JURIDICA', 'processApproved')
      .addItem('APROBADO GERENCIA', 'processNotApproved')
      .addItem('APROBADO JURIDICA', 'processResearchNeeded') 
      .addToUi();
}

/**
 * Procesa los contratos 'EN REVISIÓN JURIDICA' para enviar los respectivos correos
 */
function processApproved() {
  processRows('EN REVISIÓN JURIDICA');
}

/**
 * Procesa los contratos 'APROBADO GERENCIA' para enviar los respectivos correos
 */
function processNotApproved() {

  processRows('APROBADO GERENCIA');
}

/**
 * Procesa los contratos 'APROBADO JURIDICA' para enviar los respectivos correos
 */
function processResearchNeeded() {
  processRows('APROBADO JURIDICA');
}

/**
 * Genera un status del contrato y envia un correo de acuerdo al estado
 * @param {string} action estado
 * @param {string} emailTemplate Plantilla de correo
 */
function processRows(action, emailTemplate=null) {
  var spr = SpreadsheetApp.openById('1N0gpzFJjHWegKMDDcrXJB9c9c5BPP1zEddLPuBPajIA');
  var ss = spr.getSheetByName('Responses');

  // Toma la ID del cocumento y la transforma en un Template para Email
  let templateRows = spr.getSheetByName(TEMPLATES_SHEET).getDataRange().getValues();
  let templates = templateRows
      .reduce((result, row) => result.set(row[0], row[1]), new Map());

  // carga la informacion tomada de las columnas
  let dataRange = ss.getDataRange();
  let rows = dataRange.getValues();
  let headers = rows.shift();

  let statusRange = dataRange.offset(1, headers.indexOf(STATUS), rows.length, 1);
  let statusValues = statusRange.getValues();

  // Process each row, send an email if necessary and update the `statusValues`.
  rows
      // Convert the row arrays into objects.
      // Start with an empty object, then create a new field
      // for each header name using the corresponding row value.
      .map(rowArray => headers.reduce((rowObject, fieldName, i) => {
        rowObject[fieldName] = rowArray[i];
        return rowObject;
      }, {}))

      // Add the row index (0-based) to the row object, this is used to update
      // the status of the rows that were modified.
      // We do this because the indices won't match after the next `filter` operation.
      // We use the spread operator to unpack the `row` object.
      // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Spread_syntax
      .map((row, i) => ({...row, rowIndex: i}))

      // From all the rows, filter out and only keep the ones that match the
      // action and the status is empty.&& !row[STATUS] 
      .filter(row => row[ACTION] == action && !row[STATUS]  )

      // Send an email and update the status in `statusValues`.
      // We don't need a return value so we use `forEach` instead of `map`.
      .forEach(row => {
        // We start with the doc template HTML body, and then we replace
        // each '{{fieldName}}' with the row's respective value.
        let emailBody = headers.reduce(
          (result, fieldName) => result.replace(`{{${fieldName.toUpperCase()}}}`, row[fieldName]),
          docToHtml(templates.get(emailTemplate || action))
        );

        // Try to send an email, or get the error if it fails.
        let status;
        try {
          MailApp.sendEmail({
            to: row[EMAIL],
            subject: `ADMINISTRADOR DE CONTRATOS ESTADO: ${row[ACTION]}`,
            htmlBody: emailBody,
          });
          status = `${row[ACTION]}: ${new Date}`;
        } catch (e) {
          status = `Error: ${e}`;
        }

        // Update the `statusValues` with the new status or error.
        // We use the `rowIndex` from before to update the correct
        // row in `statusValues`.
        statusValues[row.rowIndex][0] = status;
        Logger.log(`Row ${row.rowIndex+2}: ${status}`);
      });

  // Write statusValues back into the sheet "status" column.
  statusRange.setValues(statusValues);
}

/**
 * Fetches a Google Doc as an HTML string.
 * @param {string} docUrl - The URL of a Google Doc to fetch content from.
 * @return {string} The Google Doc rendered as an HTML string.
 */
function docToHtml(docUrl) {
  let docId = DocumentApp.openByUrl(docUrl).getId();
  return UrlFetchApp.fetch(
    `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=html`,
    {
      method: 'GET',
      headers: {'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`},
      muteHttpExceptions: true,
    },
  ).getContentText();
}

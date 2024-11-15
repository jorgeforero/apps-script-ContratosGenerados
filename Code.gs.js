/**
 * Populate Doc
 * Permite llenar y generar el documento de acuerdo ( Contrato ) a partir
 * de los datos recolectados en la forma de nuevo cliente
 */

// Documento template del acuerdo - TPL_EAD_Convenio Uso Plataforma 8
const DOCTPLID = '__ID_DEL_DOCUMENTO_TEMPLATE__';
// Ubicación documento generado
const LOCATION = '__TEXTO_DE_LA_RUTA_DONDE_SE_GUARDAN_LOS_CONTRATOS_GENERADOS';

/**
 * populate
 * Llena el documento de acuerdo de uso de la plataforma a partir del último registro encontrado en
 * en la hoja de calculo asociada la forma de recolección de datos
 * 
 * @param {void} - void
 * @return {string} - Url del Documento generado || error
 */
function populateDoc() {
  // Obtiene los datos de la hoja asociada al Form
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'Form Responses 1' );
  let clients = sheet.getDataRange().getValues();
  let header = clients[ 0 ];
  // Obtiene el último registro
  let lastRecord = clients[ clients.length - 1 ];
  let client = getRowAsObject( lastRecord, header );
  // Genera un nuevo documento con los datos reemplazados a partir del template DOCTPLID
  let newDoc = getDoc( DOCTPLID, client );
  let url = newDoc.getUrl();
  let name = newDoc.getName();
  return `<a href="${url}" target="_blank">${name}</a><br /><br />Ubicación: ${LOCATION}`;  
};

/**
 * getDoc
 * Genera un nuevo documento a partir del template dato y de un conjuto de datos
 * 
 * @param {string} TemplateId - id del documento que corresponde al template
 * @param {object} Record - Objeto que contiene los datos. La llaves del objeto son usadas para identificar los reemplazables en el template
 * @return {object} - documento generado 
 */
function getDoc( TemplateId, Record ) {
  // Obtiene la plantilla del documento
  let template = DriveApp.getFileById( TemplateId );
  // Crea una copia del documento plantilla y lo nombra con el valor que viene en docname del objeto
  let file = template.makeCopy( Record.doc );
  let newDoc = DocumentApp.openById( file.getId() );
  let body = newDoc.getBody();
  // Reemplaza los marcadores de posición con los valores del objeto
  for ( let key in Record ) {
    let value = Record[ key ];
    // A partir del nombre de la llave del objeto, genera el nombre del reemplazable
    let marker = `XX_${key}_XX`;
    // busca el marcador ( todas las ocurrencias ) en documento y si lo encuentra lo reemplaza en el documento
    body.replaceText( marker, value );
  };
  // Guarda el nuevo documento y retorna su ID
  newDoc.saveAndClose();
  return newDoc;
};

/**
 * getRowAsObject
 * Obtiene un objeto con los valores de la fila dada: RowData. Toma los nombres de las llaves del parámtero Header. Las llaves
 * son dadas en minusculas y los espacios reemplazados por _
 * 
 * @param {array} RowData - Arreglo con los datos de la fila de la hoja
 * @param {array} Header - Arreglo con los nombres del encabezado de la hoja
 * @return {object} obj - Objeto con los datos de la fila y las propiedades nombradas de acuerdo a Header
 */
 function getRowAsObject( RowData, Header ) {
  let obj = {};
  for ( let indx=0; indx<RowData.length; indx++ ) {
    obj[ Header[ indx ].toLowerCase().replace( /\s/g, '_' ) ] = RowData[ indx ];
  };//for
  return obj;
};

/* Interfase */

/**
* onOpen
**/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu( 'Acciones' )
      .addItem( '🧾 Generar Contrato', 'getGetDocSB' )
      .addToUi();
};

/**
* getGetDocSB
**/
function getGetDocSB() {  
  var tpl_form = HtmlService.createHtmlOutputFromFile( 'tpl_sb_createDoc.html' ).getContent();
  // Despliegue del Panel Lateral
  var tpl_form_d = HtmlService.createHtmlOutput( tpl_form ).setTitle( 'Generar Documento' );
  SpreadsheetApp.getUi().showSidebar( tpl_form_d );
};

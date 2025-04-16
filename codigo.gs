// Archivo principal de Google Apps Script

// ID de la hoja de cálculo
const SPREADSHEET_ID = '1nn7E27XSxW0qQ6WfRSoNiAmEauNb2quzYv9eaztjtUM';

// Variables globales
const SHEET_NAME = 'Registros';
const USER_COLUMN = 'Usuario Creador'; // Nueva columna para asociar registros con usuarios

// Constantes para autenticación
const USUARIOS = [
  { username: 'admin', password: 'fecor2025', role: 'admin' },
  { username: 'fecore1', password: '123', role: 'user' },
  { username: 'fecore2', password: '123', role: 'user' },
  { username: 'fecore3', password: '123', role: 'user' },
  { username: 'fecore4', password: '123', role: 'user' }
  // Añadir más usuarios según sea necesario
];

// Propiedades para manejo de sesión
const SESION_DURATION_SECONDS = 3600; // 1 hora

/**
 * Función doGet - Punto de entrada para la aplicación web
 */
function doGet(e) {
  // Obtener el parámetro de acción o usar 'login' por defecto
  var action = e.parameter.action || 'login';
  
  // Verificar si hay una sesión activa
  var userSession = getUserSession();
  
  // Si no hay sesión activa y no es la página de login, redirigir al login
  if (!userSession && action !== 'login') {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sistema de Registro - Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Manejar las diferentes acciones
  switch(action) {
    case 'login':
      return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Sistema de Registro - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    case 'formulario':
      return HtmlService.createTemplateFromFile('formulario')
        .evaluate()
        .setTitle('Formulario de Registro')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    case 'historial':
      return HtmlService.createTemplateFromFile('historial')
        .evaluate()
        .setTitle('Historial de Registros')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    default:
      return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Sistema de Registro - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Incluir archivos HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtener la hoja de cálculo por ID
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Crear la hoja de cálculo y configurar encabezados si no existe
 */
function setupSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    
    // Definir encabezados incluyendo el nuevo campo de Usuario Creador
    var headers = [
      'ID',
      'Fecha de Registro',
      USER_COLUMN, // Nueva columna para el usuario que creó el registro
      // Información General
      'Fiscalía',
      'Fiscal a Cargo',
      // Detalles del Caso
      'Unidad de Inteligencia',
      'Instructor a Cargo',
      'Forma de Inicio de Investigación',
      'Carpeta Fiscal',
      'Fecha del Hecho',
      'Fecha Ingreso Carpeta Fiscal',
      'Etapas del Caso', 
      // Información del Agraviado
      'Tipo de Agraviado',
      'Agraviados',
      'Función que Ejerce',
      'Tipo de Empresa',
      // Información del Denunciado
      'Delitos',
      'Lugar de los Hechos',
      'Denunciados',
      'Datos de Interés del Denunciado',
      'Nombre/Apodo Banda Criminal',
      'Cantidad de Miembros de Banda',
      'Modalidad de Violencia',
      'Modalidad de Amenaza',
      'Atentados Cometidos',
      // Instrumentos y Métodos de Extorsión
      'Instrumentos de Extorsión',
      'Forma de Pago',
      'Números Telefónicos',
      'IMEI de Teléfonos',
      'Cuenta de Pago',
      'Titulares de Pago',
      // Datos de Interés de Pagos
      'Tipo de Pago',
      'Monto Solicitado',
      'Monto Pagado',
      'Número de Pagos',
      'Otros Tipos de Pago',
      // Sumilla y Observaciones
      'Sumilla de Hechos',
      'Observaciones'
    ];
    
    // Establecer los encabezados en la primera fila
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Dar formato a la fila de encabezados
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    // Ajustar ancho de columnas
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Obtiene todos los encabezados de la hoja
 */
function getHeaders() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers;
}

/**
 * Genera un ID único para nuevos registros
 */
function generateUniqueId() {
  return Utilities.getUuid();
}

/**
 * Verifica y limpia la hoja de cálculo si es necesario
 */
function checkAndSetupSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    setupSheet();
  }
}

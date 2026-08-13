/**
 * FIX VALLEY — Backend de Inventario
 * -----------------------------------
 * Este script convierte una Google Sheet en una mini base de datos con
 * su propia dirección web (API), para que la app web (en GitHub Pages)
 * pueda leer y guardar artículos sin necesidad de un servidor propio.
 *
 * INSTRUCCIONES DE INSTALACIÓN (ver README.md para el paso a paso completo):
 * 1. Crea una Google Sheet nueva.
 * 2. En la hoja 1, ponle de nombre "Articulos" y en la fila 1 escribe estos
 *    encabezados, exactamente en este orden:
 *    ID | Categoria | Marca | Tipo | Especificacion | Color | SKU | Costo | PrecioVenta | Stock | StockMinimo | Ubicacion | Actualizado
 * 3. Extensiones > Apps Script. Borra el contenido y pega este archivo completo.
 * 4. Guarda. Implementar > Nueva implementación > tipo "Aplicación web".
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que te da ("URL de la aplicación web") y pégala en la app
 *    web, en el campo de configuración.
 */

var SHEET_NAME = 'Articulos';
var HEADERS = ['ID', 'Categoria', 'Marca', 'Tipo', 'Especificacion', 'Color', 'SKU', 'Costo', 'PrecioVenta', 'Stock', 'StockMinimo', 'Ubicacion', 'Actualizado'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rowToObj_(row) {
  var obj = {};
  HEADERS.forEach(function (h, i) { obj[h] = row[i]; });
  return obj;
}

function doGet(e) {
  var action = (e.parameter.action || 'list');
  try {
    if (action === 'list') return jsonResponse_({ ok: true, data: listArticulos_() });
    if (action === 'add') return jsonResponse_({ ok: true, data: addArticulo_(e.parameter) });
    if (action === 'update') return jsonResponse_({ ok: true, data: updateArticulo_(e.parameter) });
    if (action === 'delete') return jsonResponse_({ ok: true, data: deleteArticulo_(e.parameter.id) });
    return jsonResponse_({ ok: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// Nota: se dejó de usar doPost a propósito. Los navegadores bloquean, por
// seguridad (CORS), las respuestas de POST cuando la app web y la API viven
// en dominios distintos (GitHub Pages vs script.google.com). Usar GET para
// todo evita ese bloqueo y funciona igual de bien para este caso de uso.

function listArticulos_() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = values.slice(1); // quitar encabezados
  return rows
    .filter(function (r) { return r[0] !== ''; })
    .map(rowToObj_);
}

function addArticulo_(data) {
  var sheet = getSheet_();
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  var row = [
    id,
    data.Categoria || '',
    data.Marca || '',
    data.Tipo || '',
    data.Especificacion || '',
    data.Color || '',
    data.SKU || '',
    Number(data.Costo) || 0,
    Number(data.PrecioVenta) || 0,
    Number(data.Stock) || 0,
    Number(data.StockMinimo) || 0,
    data.Ubicacion || '',
    now
  ];
  sheet.appendRow(row);
  return rowToObj_(row);
}

function findRowIndexById_(sheet, id) {
  var ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // +2: encabezado + índice base 1
  }
  return -1;
}

function updateArticulo_(data) {
  var sheet = getSheet_();
  var rowIndex = findRowIndexById_(sheet, data.ID);
  if (rowIndex === -1) throw new Error('No se encontró el artículo con ID ' + data.ID);
  var now = new Date().toISOString();
  var row = [
    data.ID,
    data.Categoria || '',
    data.Marca || '',
    data.Tipo || '',
    data.Especificacion || '',
    data.Color || '',
    data.SKU || '',
    Number(data.Costo) || 0,
    Number(data.PrecioVenta) || 0,
    Number(data.Stock) || 0,
    Number(data.StockMinimo) || 0,
    data.Ubicacion || '',
    now
  ];
  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
  return rowToObj_(row);
}

function deleteArticulo_(id) {
  var sheet = getSheet_();
  var rowIndex = findRowIndexById_(sheet, id);
  if (rowIndex === -1) throw new Error('No se encontró el artículo con ID ' + id);
  sheet.deleteRow(rowIndex);
  return { deleted: id };
}

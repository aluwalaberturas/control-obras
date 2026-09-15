/**
 * ALUWAL - Control de Obras
 * Backend Google Apps Script. Pegar en el editor de Apps Script
 * de un Google Sheet nuevo (Extensiones > Apps Script), correr
 * setup() una vez, y despues Implementar > Nueva implementacion
 * > Aplicacion web (Ejecutar como: yo, Acceso: cualquiera).
 */

const SHEET_OBRAS = 'Obras';
const SHEET_MOV = 'Movimientos';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const obras = sheetToObjects(ss.getSheetByName(SHEET_OBRAS));
    const movimientos = sheetToObjects(ss.getSheetByName(SHEET_MOV));
    return jsonOut({ status: 'ok', obras: obras, movimientos: movimientos });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const accion = body.accion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (accion === 'nueva_obra') {
      const sh = ss.getSheetByName(SHEET_OBRAS);
      sh.appendRow([
        body.id, body.nombre, body.cliente || '', body.direccion || '',
        body.fechaInicio || '', body.estado || 'activa', body.notas || '',
        new Date()
      ]);
      return jsonOut({ status: 'ok' });
    }

    if (accion === 'editar_obra') {
      updateRowById(ss.getSheetByName(SHEET_OBRAS), body.id, {
        nombre: body.nombre, cliente: body.cliente, direccion: body.direccion,
        fechaInicio: body.fechaInicio, estado: body.estado, notas: body.notas
      });
      return jsonOut({ status: 'ok' });
    }

    if (accion === 'nuevo_movimiento') {
      const sh = ss.getSheetByName(SHEET_MOV);
      sh.appendRow([
        body.id, body.obraId, body.fecha, body.tipo, body.descripcion || '',
        body.rubro || '', body.proveedor || '', body.monto || 0,
        body.persona || '', body.dias || '', body.cargadoPor || '',
        body.conIva || '', body.iva || 0, new Date()
      ]);
      return jsonOut({ status: 'ok' });
    }

    if (accion === 'editar_movimiento') {
      updateRowById(ss.getSheetByName(SHEET_MOV), body.id, {
        obraId: body.obraId, fecha: body.fecha, tipo: body.tipo,
        descripcion: body.descripcion || '', rubro: body.rubro || '',
        proveedor: body.proveedor || '', monto: body.monto || 0,
        persona: body.persona || '', dias: body.dias || '',
        cargadoPor: body.cargadoPor || '', conIva: body.conIva || '',
        iva: body.iva || 0
      });
      return jsonOut({ status: 'ok' });
    }

    if (accion === 'eliminar_movimiento') {
      deleteRowById(ss.getSheetByName(SHEET_MOV), body.id);
      return jsonOut({ status: 'ok' });
    }

    return jsonOut({ status: 'error', message: 'accion desconocida: ' + accion });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function findRowIndexById(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function updateRowById(sheet, id, fields) {
  const row = findRowIndexById(sheet, id);
  if (row === -1) return;
  const headers = sheet.getDataRange().getValues()[0];
  headers.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(fields, h) && fields[h] !== undefined) {
      sheet.getRange(row, i + 1).setValue(fields[h]);
    }
  });
}

function deleteRowById(sheet, id) {
  const row = findRowIndexById(sheet, id);
  if (row !== -1) sheet.deleteRow(row);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let obras = ss.getSheetByName(SHEET_OBRAS);
  if (!obras) obras = ss.insertSheet(SHEET_OBRAS);
  obras.getRange(1, 1, 1, 8).setValues([[
    'id', 'nombre', 'cliente', 'direccion', 'fechaInicio', 'estado', 'notas', 'fechaCreacion'
  ]]);

  let mov = ss.getSheetByName(SHEET_MOV);
  if (!mov) mov = ss.insertSheet(SHEET_MOV);
  mov.getRange(1, 1, 1, 14).setValues([[
    'id', 'obraId', 'fecha', 'tipo', 'descripcion', 'rubro', 'proveedor',
    'monto', 'persona', 'dias', 'cargadoPor', 'conIva', 'iva', 'fechaCreacion'
  ]]);

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);
}

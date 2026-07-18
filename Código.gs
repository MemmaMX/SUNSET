const CONFIG = {
  FUENTE_HOJA: 'JOIA Paraíso Resort',
  COL_HAB: 2,
  COL_CHECKOUT: 8,
  HOJA_DESTINO: 'COCKTAIL SUNSET SPAM',
  FILA_LINK: 3,
  COL_LINK: 3,
  COL_LINK_ANCHO: 8,
  FILA_FECHA: 5,
  FILA_INICIO: 6,
  COL_INICIO: 2,
  ROWS_PER_BLOCK: 29,
  COL_GAP: 1,
  URL_WEB_APP: 'https://script.google.com/a/macros/iberostar.com/s/AKfycbxyGOheNRg1h1DLUKEy6uY6S_OCmg0OdILhf6Jdc53gpP48MDfsi50K4NxbjeAAk_oB/exec'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('COCKTAIL SUNSET')
    .addItem('Actualizar tabla', 'actualizarTablaActual')
    .addItem('Abrir flyer', 'abrirFlyer')
    .addToUi();
}

function actualizarTablaActual() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No hay una hoja de cálculo activa. Abre el archivo correcto.');

    const sheet = ss.getSheetByName(CONFIG.HOJA_DESTINO) || ss.getActiveSheet();
    const checks = leerChecks(sheet);

    const linkValue = sheet.getRange(CONFIG.FILA_LINK, CONFIG.COL_LINK).getValue();
    const id = extraerId(String(linkValue || ''));
    if (!id) throw new Error('Pega un link o ID válido de Google Sheets en C3.');

    const habs = obtenerHabitaciones(id);
    prepararHoja(sheet, linkValue);
    limpiarAreaTabla(sheet);
    generarTabla(sheet, habs, checks);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

function abrirFlyer() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head><base target="_top"><meta charset="utf-8"></head>
    <body>
      <script>
        window.open(
          '${CONFIG.URL_WEB_APP}',
          'FlyerWindow',
          'width=1100,height=800,left=150,top=80,resizable=yes,scrollbars=yes'
        );
        google.script.host.close();
      <\/script>
    </body>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html)
      .setWidth(1)
      .setHeight(1),
    ' '
  );
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Flyer')
    .setTitle('COCKTAIL SUNSET FLYER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerHabitaciones(id) {
  if (!id) throw new Error('El ID del archivo externo está vacío.');

  let ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error('No se pudo abrir el archivo externo. Verifica el link en C3.');
  }

  const hoja = ss.getSheetByName(CONFIG.FUENTE_HOJA);
  if (!hoja) throw new Error('No se encontró la hoja "' + CONFIG.FUENTE_HOJA + '" en el archivo externo.');

  const total = hoja.getLastRow();
  if (total < 2) return [];

  const datos = hoja.getRange(2, 1, total - 1, Math.max(CONFIG.COL_HAB, CONFIG.COL_CHECKOUT)).getValues();
  const viernes = getViernesSemanaActual();
  const vistos = new Set();
  const resultado = [];

  for (const f of datos) {
    const numero = extraerNumero(String(f[CONFIG.COL_HAB - 1] || ''));
    if (!numero || !/^7[234]/.test(numero)) continue;

    const checkout = normalizarFecha(f[CONFIG.COL_CHECKOUT - 1]);
    if (!checkout || checkout <= viernes) continue;
    if (vistos.has(numero)) continue;

    vistos.add(numero);
    resultado.push(numero);
  }

  return resultado.sort((a, b) => Number(a) - Number(b));
}

function generarTabla(sheet, habs, checks) {
  const bloques = Math.ceil(habs.length / CONFIG.ROWS_PER_BLOCK) || 1;
  const groupWidth = 4 + CONFIG.COL_GAP;
  const filas = 1 + CONFIG.ROWS_PER_BLOCK;
  const columnas = bloques * groupWidth - CONFIG.COL_GAP;

  const matriz = Array(filas).fill(null).map(() => Array(columnas).fill(''));

  for (let b = 0; b < bloques; b++) {
    const c = b * groupWidth;
    matriz[0][c] = 'HAB';
    matriz[0][c + 1] = 'ENVIADO';
    matriz[0][c + 2] = 'YA ENVIADO';
    matriz[0][c + 3] = 'SOCIOS';
  }

  for (let i = 0; i < habs.length; i++) {
    const b = Math.floor(i / CONFIG.ROWS_PER_BLOCK);
    const r = 1 + (i % CONFIG.ROWS_PER_BLOCK);
    const c = b * groupWidth;
    const hab = habs[i];
    matriz[r][c] = Number(hab);
    const prev = checks.get(hab) || [false, false, false];
    matriz[r][c + 1] = prev[0];
    matriz[r][c + 2] = prev[1];
    matriz[r][c + 3] = prev[2];
  }

  const rango = sheet.getRange(CONFIG.FILA_INICIO, CONFIG.COL_INICIO, filas, columnas);
  rango.setValues(matriz);
  rango.setBackground('#FFFFFF');

  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (let b = 0; b < bloques; b++) {
    const c = CONFIG.COL_INICIO + b * groupWidth;
    sheet.getRange(CONFIG.FILA_INICIO + 1, c + 1, CONFIG.ROWS_PER_BLOCK, 3).setDataValidation(cb);
  }

  formatar(sheet, bloques, filas);
  escribirFecha(sheet, columnas);
}

function formatar(sheet, bloques, filas) {
  const groupWidth = 4 + CONFIG.COL_GAP;

  for (let b = 0; b < bloques; b++) {
    const c = CONFIG.COL_INICIO + b * groupWidth;
    const tabla = sheet.getRange(CONFIG.FILA_INICIO, c, filas, 4);
    tabla.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    tabla.setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(CONFIG.FILA_INICIO, c, 1, 1).setBackground('#009690').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(CONFIG.FILA_INICIO, c + 1, 1, 1).setBackground('#296374').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(CONFIG.FILA_INICIO, c + 2, 1, 1).setBackground('#629FAD').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange(CONFIG.FILA_INICIO, c + 3, 1, 1).setBackground('#EDEDCE').setFontColor('#000000').setFontWeight('bold');

    if (b < bloques - 1) {
      sheet.getRange(CONFIG.FILA_INICIO, c + 4, filas, 1)
        .setBorder(false, true, false, false, false, false, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
    }
  }
}

function escribirFecha(sheet, columnas) {
  const celda = sheet.getRange(CONFIG.FILA_FECHA, CONFIG.COL_INICIO, 1, columnas);
  celda.merge().clear();
  celda.setValue(formatearFecha(new Date()));
  celda.setBackground('#0C2C55');
  celda.setFontColor('#FFFFFF');
  celda.setFontWeight('bold');
  celda.setHorizontalAlignment('center').setVerticalAlignment('middle');
  celda.setFontSize(14);
}

function leerChecks(sheet) {
  const checks = new Map();
  const datos = sheet.getDataRange().getValues();
  const groupWidth = 4 + CONFIG.COL_GAP;

  for (let r = 0; r < datos.length; r++) {
    for (let c = 0; c + 3 < datos[r].length; c += groupWidth) {
      if (datos[r][c] === 'HAB') {
        for (let rr = r + 1; rr <= r + CONFIG.ROWS_PER_BLOCK && rr < datos.length; rr++) {
          const hab = String(datos[rr][c] || '').trim();
          if (/^7[234]\d+$/.test(hab)) {
            checks.set(hab, [
              datos[rr][c + 1] === true,
              datos[rr][c + 2] === true,
              datos[rr][c + 3] === true
            ]);
          }
        }
      }
    }
  }

  return checks;
}

function prepararHoja(sheet, linkValue) {
  sheet.setHiddenGridlines(true);
  const maxRows = Math.max(sheet.getMaxRows(), 100);
  const maxCols = Math.max(sheet.getMaxColumns(), 30);
  if (maxRows > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());
  if (maxCols > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());

  sheet.getRange(1, 1, maxRows, maxCols).setBackground('#FCFCE3');

  const linkRange = sheet.getRange(CONFIG.FILA_LINK, CONFIG.COL_LINK, 1, CONFIG.COL_LINK_ANCHO);
  linkRange.clear();
  linkRange.setValue(linkValue);
  linkRange.merge();
  linkRange.setBackground('#FFFFFF');
  linkRange.setFontColor('#000000');
  linkRange.setHorizontalAlignment('left');
  linkRange.setVerticalAlignment('middle');
  linkRange.setFontSize(11);
  linkRange.setBorder(true, true, true, true, null, null, '#BBBBBB', SpreadsheetApp.BorderStyle.SOLID);
}

function limpiarAreaTabla(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const filas = maxRows - CONFIG.FILA_FECHA + 1;
  const cols = maxCols - CONFIG.COL_INICIO + 1;
  if (filas > 0 && cols > 0) {
    sheet.getRange(CONFIG.FILA_FECHA, CONFIG.COL_INICIO, filas, cols).clear();
  }
}

function getViernesSemanaActual() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const dia = hoy.getDay();
  const diff = 5 - dia;
  return new Date(hoy.setDate(hoy.getDate() + diff));
}

function formatearFecha(fecha) {
  const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  return `${dias[fecha.getDay()]}, ${('0' + fecha.getDate()).slice(-2)} de ${meses[fecha.getMonth()]} del ${fecha.getFullYear()}`;
}

function extraerId(urlOId) {
  const input = String(urlOId || '').trim();
  const m = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{30,}$/.test(input)) return input;
  return '';
}

function extraerNumero(str) {
  const m = String(str).match(/\d+/);
  return m ? m[0] : null;
}

function normalizarFecha(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    const d = new Date(valor);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const n = Number(valor);
  if (!isNaN(n) && n > 30000 && n < 100000) {
    const d = new Date((n - 25569) * 86400 * 1000);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const partes = String(valor).split(/[\/\-.]/);
  if (partes.length === 3) {
    let d = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
    if (isNaN(d.getTime())) d = new Date(parseInt(partes[2]), parseInt(partes[0]) - 1, parseInt(partes[1]));
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }
  return null;
}
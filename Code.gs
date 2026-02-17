const a_sheetId = '1tT-ximGZnuNItSpszy3Y6kmNRF1E5tXEJDFmJWNqLo8';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema de Gestión Movilnet')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const MASTER_FIELD_MAPPING = {
  'prefijo': 'Prefijo',
  'contacto': 'Nº Contacto',
  'cola_atencion': 'Cola De Atención',
  'gestor': 'Ejecutivo De Atención',
  'fecha_atencion': 'Fecha De Atención',
  'numero_contrato': 'Número De Contrato',
  'prefijo2': 'Prefijo2',
  'numero': 'Número',
  'linea_gestionar': 'Línea A Gestionar',
  'tecnologia': 'Tecnologia',
  'tipo_linea': 'Tipo De Línea',
  'tipo_cliente': 'Tipo De Cliente',
  'segmento': 'Segmento',
  'nombre_apellido': 'Nombre Y Apellido Del Titular / Usuario',
  'tipo_documento': 'Tipo de documento',
  'cedula_rif': 'N° De Cédula O Rif',
  'numero_sim': 'Número De Sim',
  'correo': 'Correo Electrónico',
  'telefono_contacto': 'Teléfono De Contacto',
  'motivo_contacto': 'Motivo Del Contacto',
  'detalle_requerimiento': 'Detalle Del Requerimiento',
  'escalado': 'Escalado',
  'estatus': 'Estatus',
  'observaciones': 'Observaciones',
  'estado': 'Estado',
  'ciudad': 'Ciudad',
  'municipio': 'Municipio',
  'parroquia': 'Parroquia (Falla Operativa/ Reclamo)',
  'direccion': 'Dir. (Av, Cll, Crr, Esq, Ctra, Tranv)',
  'referencia': 'Ubicación O Punto De Referencia (Falla Operativa/Reclamo)',
  'genero': 'Género',
  'atencion_preferencial': '¿Requiere Atención Preferencial?',
  'discapacidad': '¿Posee discapacidad?',
  'adulto_mayor': '¿Adulto Mayor?'
};

// -------------------- REGISTRO (EXISTENTE) --------------------
function buscarLinea(lineaAGestionar) {
  try {
    if (!lineaAGestionar) return null;
    const spreadsheet = SpreadsheetApp.openById(a_sheetId);
    const sheet = spreadsheet.getSheetByName('Datos_Operativos');
    if (!sheet || sheet.getLastRow() < 2) return null;
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const lineaIndex = findHeaderIndex(headers, 'Línea A Gestionar');
    if (lineaIndex === -1) {
      console.error('No se encontró la columna "Línea A Gestionar".');
      return null;
    }
    const normalizeDigits = s => (s === null || s === undefined) ? '' : s.toString().replace(/\D/g, '').trim();
    const normalizeText = s => normalizeHeaderName((s === null || s === undefined) ? '' : s.toString());
    const targetRaw = lineaAGestionar === null || lineaAGestionar === undefined ? '' : lineaAGestionar.toString();
    const targetDigits = normalizeDigits(targetRaw);
    const targetNormalized = normalizeText(targetRaw);
    for (let i = values.length - 1; i > 0; i--) {
      const cell = values[i][lineaIndex];
      const cellRaw = (cell === null || cell === undefined) ? '' : cell.toString();
      const cellDigits = normalizeDigits(cellRaw);
      const cellNormalized = normalizeText(cellRaw);
      if (cellRaw === targetRaw || (cellDigits && targetDigits && cellDigits === targetDigits) || (cellNormalized && targetNormalized && cellNormalized === targetNormalized)) {
        const record = values[i];
        const result = {};
        headers.forEach((header, index) => {
          const formKey = Object.keys(MASTER_FIELD_MAPPING).find(key => normalizeHeaderName(MASTER_FIELD_MAPPING[key]) === normalizeHeaderName(header));
          if (formKey) result[formKey] = record[index];
        });
        return result;
      }
    }
    return null;
  } catch (error) {
    console.error('Error en buscarLinea:', error);
    return null;
  }
}

function normalizeHeaderName(s) {
  if (!s) return '';
  return s.toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function findHeaderIndex(headers, target) {
  if (!headers || !headers.length) return -1;
  var t = normalizeHeaderName(target);
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeaderName(headers[i]) === t) return i;
  }
  // try partial contains
  for (var j = 0; j < headers.length; j++) {
    if (normalizeHeaderName(headers[j]).indexOf(t) !== -1) return j;
  }
  return -1;
}

function guardarRegistro(formData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(a_sheetId);
    const hojaDatos = spreadsheet.getSheetByName('Datos_Operativos') || spreadsheet.insertSheet('Datos_Operativos');
    let headers = [];
    if (hojaDatos.getLastRow() > 0) {
      headers = hojaDatos.getRange(1, 1, 1, hojaDatos.getLastColumn()).getValues()[0];
    }
    if (headers.length === 0) {
      const newHeaders = ['Nº', 'Resultado', 'Tipo', 'Canal De Atención', ...Object.values(MASTER_FIELD_MAPPING)];
      hojaDatos.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      headers = newHeaders;
    }

    const fechaHeaderIndex = headers.indexOf('Fecha De Atención');
    if (fechaHeaderIndex !== -1) {
      try {

        hojaDatos.getRange(2, fechaHeaderIndex + 1, hojaDatos.getMaxRows() - 1 || 1, 1).setNumberFormat('dd/MM/yyyy');
      } catch (e) {

        console.error('No se pudo aplicar formato de fecha a la columna:', e);
      }
    }
    const newRow = [];
    const formType = formData.tipo === 'osac' ? 'OSAC' : 'Canal Virtual';
    headers.forEach(header => {
      let value = '';
      switch (header) {
        case 'Nº': value = hojaDatos.getLastRow(); break;
        case 'Resultado': value = formData.escalado === 'NO' ? 'COMPLETO' : 'INCOMPLETO'; break;
        case 'Tipo': value = formType; break;
        case 'Canal De Atención': value = formType === 'OSAC' ? 'OSAC' : 'Canal Virtual'; break;
        case 'Fecha De Atención':

          if (formData && formData.fecha_atencion) {
            try {

              const d = new Date(formData.fecha_atencion + 'T00:00:00');
              if (!isNaN(d.getTime())) value = d; else value = formData.fecha_atencion;
            } catch (e) {
              value = formData.fecha_atencion;
            }
          }
          break;
        default:
          const formKey = Object.keys(MASTER_FIELD_MAPPING).find(key => MASTER_FIELD_MAPPING[key] === header);
          if (formKey && formData[formKey] !== undefined) value = formData[formKey];
          break;
      }
      newRow.push(value);
    });
    hojaDatos.appendRow(newRow);
    return `Registro de ${formType} guardado correctamente.`;
  } catch (error) {
    console.error(`Error en guardarRegistro:`, error);
    throw new Error('Error al guardar el registro: ' + error.message);
  }
}

function guardarRegistroOSAC(formData) { return guardarRegistro(formData); }
function guardarRegistroCanal(formData) { return guardarRegistro(formData); }

// ------------------------------------------------------------
// 2. DASHBOARD (INICIO)
// ------------------------------------------------------------
function obtenerEstadisticasDashboard() {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet || sheet.getLastRow() < 2) {
      return { total: 0, completos: 0, incompletos: 0, promedioDiario: 0 };
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    const resultadoIndex = findHeaderIndex(headers, 'Resultado');
    const fechaIndex = findHeaderIndex(headers, 'Fecha De Atención');
    let total = rows.length;
    let completos = 0, incompletos = 0;
    const today = new Date();
    const thirtyDaysAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
    let last30DaysCount = 0;
    rows.forEach(row => {
      if (row[resultadoIndex] === 'COMPLETO') completos++;
      else if (row[resultadoIndex] === 'INCOMPLETO') incompletos++;
      if (fechaIndex !== -1 && row[fechaIndex] instanceof Date) {
        if (row[fechaIndex] >= thirtyDaysAgo && row[fechaIndex] <= today) last30DaysCount++;
      }
    });
    const promedioDiario = Math.round((last30DaysCount / 30) * 10) / 10;
    return { total, completos, incompletos, promedioDiario };
  } catch (e) {
    console.error('Error en obtenerEstadisticasDashboard:', e);
    return { total: 0, completos: 0, incompletos: 0, promedioDiario: 0 };
  }
}

function obtenerUltimosRegistros(limite = 10) {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    const numIndex = findHeaderIndex(headers, 'Nº');
    const tipoIndex = findHeaderIndex(headers, 'Tipo');
    const resultadoIndex = findHeaderIndex(headers, 'Resultado');
    const ejecutivoIndex = findHeaderIndex(headers, MASTER_FIELD_MAPPING.gestor);
    const fechaIndex = findHeaderIndex(headers, 'Fecha De Atención');
    const ultimos = [];
    for (let i = rows.length - 1; i >= 0 && ultimos.length < limite; i--) {
      const row = rows[i];
      ultimos.push({
        numero: row[numIndex] || (i + 2),
        tipo: row[tipoIndex] || '',
        resultado: row[resultadoIndex] || '',
        ejecutivo: row[ejecutivoIndex] || '',
        fecha: row[fechaIndex] ? Utilities.formatDate(row[fechaIndex], Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''
      });
    }
    return ultimos;
  } catch (e) {
    console.error('Error en obtenerUltimosRegistros:', e);
    return [];
  }
}

// ------------------------------------------------------------
// 3. CONSULTAS
// ------------------------------------------------------------
function buscarRegistrosPorLinea(linea) {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);


    const colaIndex = findHeaderIndex(headers, 'Cola De Atención');
    const ejecutivoIndex = findHeaderIndex(headers, 'Ejecutivo De Atención');
    const fechaIndex = findHeaderIndex(headers, 'Fecha De Atención');
    const lineaIndex = findHeaderIndex(headers, 'Línea A Gestionar');
    const tipoLineaIndex = findHeaderIndex(headers, 'Tipo De Línea');
    const motivoIndex = findHeaderIndex(headers, 'Motivo Del Contacto');
    const detalleIndex = findHeaderIndex(headers, 'Detalle Del Requerimiento');
    const escaladoIndex = findHeaderIndex(headers, 'Escalado');


    if (lineaIndex === -1) return [];

    const resultados = [];
    const normalizeDigits = s => (s === null || s === undefined) ? '' : s.toString().replace(/\D/g, '').trim();
    const normalizeText = s => normalizeHeaderName((s === null || s === undefined) ? '' : s.toString());

    const targetRaw = linea === null || linea === undefined ? '' : linea.toString();
    const targetDigits = normalizeDigits(targetRaw);
    const targetNormalized = normalizeText(targetRaw);

    rows.forEach((row, idx) => {
      const cell = row[lineaIndex];
      const cellRaw = (cell === null || cell === undefined) ? '' : cell.toString();
      const cellDigits = normalizeDigits(cellRaw);
      const cellNormalized = normalizeText(cellRaw);

      if (cellRaw === targetRaw || (cellDigits && targetDigits && cellDigits === targetDigits) || (cellNormalized && targetNormalized && cellNormalized === targetNormalized)) {

        const registro = {
          fila: idx + 2, 
          cola_atencion: (colaIndex !== -1) ? row[colaIndex] : '',
          ejecutivo: (ejecutivoIndex !== -1) ? row[ejecutivoIndex] : '',
          fecha_atencion: (fechaIndex !== -1 && row[fechaIndex] instanceof Date) ? Utilities.formatDate(row[fechaIndex], Session.getScriptTimeZone(), 'dd/MM/yyyy') : (row[fechaIndex] || ''),
          linea_gestionar: row[lineaIndex] || '',
          tipo_linea: (tipoLineaIndex !== -1) ? row[tipoLineaIndex] : '',
          motivo_contacto: (motivoIndex !== -1) ? row[motivoIndex] : '',
          detalle_requerimiento: (detalleIndex !== -1) ? row[detalleIndex] : '',
          escalado: (escaladoIndex !== -1) ? row[escaladoIndex] : ''
        };
        resultados.push(registro);
      }
    });

    return resultados;
  } catch (e) {
    console.error('Error en buscarRegistrosPorLinea:', e);
    return [];
  }
}

// ------------------------------------------------------------
// 4. REPORTES
// ------------------------------------------------------------
function obtenerReportesPorRango(fechaInicio, fechaFin) {
  try {

    let inicioDate, finDate;
    
    if (typeof fechaInicio === 'string') {
      inicioDate = new Date(fechaInicio + 'T00:00:00');
    } else if (fechaInicio instanceof Date) {
      inicioDate = new Date(fechaInicio);
      inicioDate.setHours(0, 0, 0, 0);
    } else {
      throw new Error('fechaInicio inválida');
    }
    
    if (typeof fechaFin === 'string') {
      finDate = new Date(fechaFin + 'T23:59:59');
    } else if (fechaFin instanceof Date) {
      finDate = new Date(fechaFin);
      finDate.setHours(23, 59, 59, 999);
    } else {
      throw new Error('fechaFin inválida');
    }

    const inicioUTC = inicioDate.getTime();
    const finUTC = finDate.getTime();

    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const fechaIndex = findHeaderIndex(headers, 'Fecha De Atención');
    if (fechaIndex === -1) {
      console.error('No se encontró la columna Fecha De Atención');
      return [];
    }

    const resultados = [];
    const timeZone = Session.getScriptTimeZone();

    rows.forEach((row, idx) => {
      const celda = row[fechaIndex];
      let fechaRegistro = null;


      if (celda instanceof Date) {
        fechaRegistro = celda;
      }

      else if (typeof celda === 'string') {

        let d = new Date(celda);
        if (!isNaN(d.getTime())) {
          fechaRegistro = d;
        } else {

          try {
            fechaRegistro = Utilities.parseDate(celda, timeZone, 'dd/MM/yyyy');
          } catch (e) {

          }
        }
      }

      else if (typeof celda === 'number') {
        fechaRegistro = new Date((celda - 25569) * 86400 * 1000);
      }

      if (fechaRegistro && !isNaN(fechaRegistro.getTime())) {
        const tiempoRegistro = fechaRegistro.getTime();
        if (tiempoRegistro >= inicioUTC && tiempoRegistro <= finUTC) {
          const registro = { fila: idx + 2 };


          headers.forEach((header, col) => {
            const formKey = Object.keys(MASTER_FIELD_MAPPING).find(key =>
              normalizeHeaderName(MASTER_FIELD_MAPPING[key]) === normalizeHeaderName(header)
            );
            if (formKey) registro[formKey] = row[col];
          });


          registro.tipo = row[findHeaderIndex(headers, 'Tipo')] || '';
          registro.resultado = row[findHeaderIndex(headers, 'Resultado')] || '';
          registro.canal_atencion = row[findHeaderIndex(headers, 'Canal De Atención')] || '';
          registro.fecha = Utilities.formatDate(fechaRegistro, timeZone, 'dd/MM/yyyy');

          resultados.push(registro);
        }
      }
    });

    return resultados;
  } catch (e) {
    console.error('Error en obtenerReportesPorRango:', e);
    return [];
  }
}

// ------------------------------------------------------------
// Obtener todos los registros
// ------------------------------------------------------------
function obtenerTodosRegistros() {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet) {
      console.error('obtenerTodosRegistros: hoja Datos_Operativos no encontrada');
      return [];
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      console.log('obtenerTodosRegistros: no hay datos (lastRow=' + lastRow + ', lastCol=' + lastCol + ')');
      return [];
    }
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0] || [];
    const rows = data.slice(1);
    const resultados = rows.map(row => {
      const obj = {};
      for (var idx = 0; idx < headers.length; idx++) {
        obj[headers[idx]] = row[idx];
      }
      return obj;
    });
    console.log('obtenerTodosRegistros: retornando ' + resultados.length + ' registros (lastRow=' + lastRow + ')');
    return resultados;
  } catch (e) {
    console.error('Error en obtenerTodosRegistros:', e);
    return [];
  }
}


function obtenerTotalRegistros() {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet) return 0;
    const last = sheet.getLastRow();
    if (last < 2) return 0;
    return Math.max(0, last - 1);
  } catch (e) {
    console.error('Error en obtenerTotalRegistros:', e);
    return 0;
  }
}


function generarCSVTodosRegistros() {
  try {
    const sheet = SpreadsheetApp.openById(a_sheetId).getSheetByName('Datos_Operativos');
    if (!sheet) {
      console.error('generarCSVTodosRegistros: hoja no encontrada');
      return '';
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return '';
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (!data || data.length < 2) return '';
    const headers = data[0];
    const rows = data.slice(1);
    const timeZone = Session.getScriptTimeZone();

    // Construir CSV
    const escapeCell = v => {
      if (v === null || v === undefined) return '""';
      if (v instanceof Date) {
        try { v = Utilities.formatDate(v, timeZone, 'dd/MM/yyyy'); } catch (e) { v = v.toString(); }
      }
      v = v.toString();
      v = v.replace(/"/g, '""');
      return '"' + v + '"';
    };

    let csv = headers.map(h => escapeCell(h)).join(';') + '\n';
    rows.forEach(r => {
      const line = [];
      for (var i = 0; i < headers.length; i++) line.push(escapeCell(r[i]));
      csv += line.join(';') + '\n';
    });
    return csv;
  } catch (e) {
    console.error('Error en generarCSVTodosRegistros:', e);
    return '';
  }
}



// ------------------------------------------------------------
// 5. USUARIOS (AUTENTICACIÓN Y CRUD)
// ------------------------------------------------------------
function inicializarHojaUsuarios() {
  const ss = SpreadsheetApp.openById(a_sheetId);
  let sheet = ss.getSheetByName('Usuarios');
  if (!sheet) {
    sheet = ss.insertSheet('Usuarios');
    sheet.getRange(1,1,1,7).setValues([['ID','Usuario','Nombre','Password','Rol','Activo','FechaCreacion']]);
    const defaultAdmin = {
      id: 1,
      usuario: 'admin',
      nombre: 'Administrador',
      password: hashPassword('admin'),
      rol: 'admin',
      activo: 'SI',
      fecha: new Date()
    };
    sheet.appendRow([1, 'admin', 'Administrador', defaultAdmin.password, 'admin', 'SI', new Date()]);
  }
  return sheet;
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function autenticarUsuario(usuario, password) {
  try {
    const sheet = inicializarHojaUsuarios();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    const userIdx = headers.indexOf('Usuario');
    const passIdx = headers.indexOf('Password');
    const nombreIdx = headers.indexOf('Nombre');
    const rolIdx = headers.indexOf('Rol');
    const activoIdx = headers.indexOf('Activo');
    for (let row of rows) {
      if (row[userIdx] === usuario && row[passIdx] === hashPassword(password) && row[activoIdx] === 'SI') {
        return {
          usuario: row[userIdx],
          nombre: row[nombreIdx],
          rol: row[rolIdx],
          autenticado: true
        };
      }
    }
    return { autenticado: false };
  } catch (e) {
    console.error('Error en autenticarUsuario:', e);
    return { autenticado: false };
  }
}

function listarUsuarios() {
  try {
    const sheet = inicializarHojaUsuarios();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    return rows.map(row => ({
      id: row[0],
      usuario: row[1],
      nombre: row[2],
      rol: row[4],
      activo: row[5],
      fecha: row[6] ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''
    }));
  } catch (e) {
    console.error('Error en listarUsuarios:', e);
    return [];
  }
}

function agregarUsuario(datos) {
  try {
    const sheet = inicializarHojaUsuarios();
    const lastRow = sheet.getLastRow();
    let newId = 1;
    if (lastRow > 1) {
      const ids = sheet.getRange(2,1, lastRow-1,1).getValues().flat();
      newId = Math.max(...ids) + 1;
    }
    sheet.appendRow([
      newId,
      datos.usuario,
      datos.nombre,
      hashPassword(datos.password),
      datos.rol,
      'SI',
      new Date()
    ]);
    return { success: true, message: 'Usuario agregado correctamente.' };
  } catch (e) {
    console.error('Error en agregarUsuario:', e);
    return { success: false, message: 'Error al agregar usuario: ' + e.message };
  }
}

function editarUsuario(id, datos) {
  try {
    const sheet = inicializarHojaUsuarios();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        const row = i + 1;
        const usuarioIdx = headers.indexOf('Usuario') + 1;
        const nombreIdx = headers.indexOf('Nombre') + 1;
        const rolIdx = headers.indexOf('Rol') + 1;
        const activoIdx = headers.indexOf('Activo') + 1;
        sheet.getRange(row, usuarioIdx).setValue(datos.usuario);
        sheet.getRange(row, nombreIdx).setValue(datos.nombre);
        sheet.getRange(row, rolIdx).setValue(datos.rol);
        sheet.getRange(row, activoIdx).setValue(datos.activo);
        if (datos.password && datos.password.trim() !== '') {
          const passIdx = headers.indexOf('Password') + 1;
          sheet.getRange(row, passIdx).setValue(hashPassword(datos.password));
        }
        return { success: true, message: 'Usuario actualizado correctamente.' };
      }
    }
    return { success: false, message: 'Usuario no encontrado.' };
  } catch (e) {
    console.error('Error en editarUsuario:', e);
    return { success: false, message: 'Error al editar usuario: ' + e.message };
  }
}

function eliminarUsuario(id) {
  try {
    const sheet = inicializarHojaUsuarios();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Usuario eliminado correctamente.' };
      }
    }
    return { success: false, message: 'Usuario no encontrado.' };
  } catch (e) {
    console.error('Error en eliminarUsuario:', e);
    return { success: false, message: 'Error al eliminar usuario: ' + e.message };
  }
}


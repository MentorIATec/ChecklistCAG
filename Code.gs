// --- CONFIGURACIÓN CRÍTICA ---
const SPREADSHEET_ID = "1rz_PBq2_nISHaTLJ6Qkqh9ULp7yYxd5K45I2N7i9zUg";
const SHEET_NAME = "Hoja 1"; // Nombre con el espacio especial
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwJE7Il3oBseJZE9OW_5VLAm3KNUYDsB8MOa9b5ucBKFUmcYEox100DwNqha5Qlvrs/exec";
const RESPONSES_SHEET_NAME = "Respuestas del Checklist"; 

// --- CONFIG EMAIL + FILTROS ---
const CONFIG = {
  recentWindowDays: 60,               // ventana de recencia para fechas
  requireAllRecentToExclude: true,    // ahora excluye SOLO si las 3 son recientes
  minDaysBetweenSends: 5,             // frecuencia mínima entre envíos iniciales
  minDaysBetweenReminders: 5,         // frecuencia mínima entre recordatorios
  subjectTemplate: 'Camino a tu Graduación — {{NOMBRE}} ({{PENDIENTES}} pendientes)',
};

// --- DETALLES DE LA SESIÓN ---
const SESSION_TITLE = "Lanzamiento Profesional: Sesión de Cierre";
const SESSION_DATE = "2025-09-17";
const SESSION_TIME_START = "11:00";
const SESSION_TIME_END = "12:00";
const SESSION_LOCATION = "Área de Mentoría, Centrales Sur 3er Piso";
const SESSION_DETAILS = "Un espacio para resolver tus últimas dudas y afinar tu plan de carrera.";

// --- MENTOR/A (firma) ---
const MENTOR_NAME = "Equipo de Desarrollo Profesional"; // Cambia aquí el nombre mostrado en la firma

// ======================================================================================
// MENÚ PERSONALIZADO
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Herramientas de Campaña')
    .addItem('✅ Verificar Nombres de Pestañas', 'verificarNombresDePestañas')
    .addSeparator()
    .addItem('1. Enviar Correos Iniciales', 'enviarCorreosInicialesConConfirmacion')
    .addItem('2. Enviar Recordatorios a Pendientes', 'enviarRecordatoriosConConfirmacion')
    .addToUi();
}

// ======================================================================================
// HERRAMIENTA DE DIAGNÓSTICO
function verificarNombresDePestañas() {
  const ui = SpreadsheetApp.getUi();
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let message = "Nombres encontrados:\n\n";
    sheets.forEach(sheet => {
      message += `"${sheet.getName()}"\n`;
    });
    ui.alert("Diagnóstico de Pestañas", message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Error", "No se pudo acceder a la hoja de cálculo activa.", ui.ButtonSet.OK);
  }
}

// ======================================================================================
// FUNCIONES DEL MENÚ
function enviarCorreosInicialesConConfirmacion() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirmación de Envío",
    "¿Estás seguro de que quieres enviar el correo inicial a TODOS los estudiantes?",
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    enviarCorreosConEnlaces("inicial");
  }
}

function enviarRecordatoriosConConfirmacion() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirmación de Envío",
    "¿Enviar un recordatorio a los estudiantes que no han respondido?",
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    enviarCorreosConEnlaces("recordatorio");
  }
}

// ======================================================================================
// LÓGICA PRINCIPAL DE ENVÍO
function enviarCorreosConEnlaces(tipo) {
  const ui = SpreadsheetApp.getUi();
  if (WEB_APP_URL === "PEGA_AQUÍ_LA_URL_DE_TU_NUEVA_IMPLEMENTACIÓN") {
    ui.alert("Error de Configuración", "Crea una nueva implementación, copia la URL y pégala en WEB_APP_URL.", ui.ButtonSet.OK);
    return;
  }

  let data;
  try {
    data = loadRowsByHeader_();
  } catch (e) {
    ui.alert("Error Crítico de Acceso", e.message, ui.ButtonSet.OK);
    return;
  }
  const { sheet, header, rows } = data;
  ensureTrackingCols_(sheet, header);

  let candidatos = rows.filter(r => !!(getBySynonyms_(r, ['email','correo','correo tec','correo institucional'])||'').toString().includes('@'));

  if (tipo === 'recordatorio') {
    // excluir quienes ya respondieron según RESPONSES_SHEET_NAME
    const responsesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);
    const responded = new Set();
    if (responsesSheet && responsesSheet.getLastRow() > 1) {
      const vals = responsesSheet.getRange(2, 2, responsesSheet.getLastRow() - 1, 1).getValues();
      vals.forEach(v => { const m = String(v[0]||'').trim(); if (m) responded.add(m); });
    }
    candidatos = candidatos.filter(r => !responded.has(String(getBySynonyms_(r, ['matricula','matrícula','id','matricula tec'])||'').trim()));
  }

  // Aplicar filtros de recencia y frecuencia
  candidatos = candidatos.filter(r => passesRecencyFilter_(r));
  candidatos = candidatos.filter(r => passesFrequency_(r, tipo === 'recordatorio' ? 'Último Recordatorio' : 'Último Envío', tipo === 'recordatorio' ? CONFIG.minDaysBetweenReminders : CONFIG.minDaysBetweenSends));

  if (!candidatos.length) {
    ui.alert('Información', 'No hay estudiantes que pasen los filtros para contactar.', ui.ButtonSet.OK);
    return;
  }

  ui.alert(`Se enviarán ${candidatos.length} correos. El proceso puede tardar.`);

  const now = new Date();
  let sent = 0;
  candidatos.forEach((row) => {
    const email = String(getBySynonyms_(row, ['email','correo','correo tec','correo institucional'])||'');
    let payload;
    try {
      payload = renderEmailEmailingCode_(row, tipo === 'recordatorio');
    } catch (err) {
      // Si no está disponible el generador, no enviar nada.
      Logger.log('ERROR renderEmailEmailingCode_: ' + (err && err.message));
      return; // skip
    }
    try {
      MailApp.sendEmail({ to: email, subject: payload.subject, htmlBody: payload.html, name: 'Camino a Tu Graduación' });
      if (tipo === 'recordatorio') {
        row['Último Recordatorio'] = now;
        row['Recordatorios'] = (toNumber_(row['Recordatorios']) || 0) + 1;
      } else {
        row['Último Envío'] = now;
        row['Envios'] = (toNumber_(row['Envios']) || 0) + 1;
      }
      writeBackRow_(sheet, header, row.__rowIndex, row);
      sent++;
    } catch (e) {
      // continúa con el siguiente
    }
  });

  // Totales históricos luego de esta corrida
  const totals = getTrackingTotals_(sheet);
  if (tipo === 'recordatorio') {
    ui.alert('Proceso Completado', `Recordatorios en esta corrida: ${sent}\nTotal recordatorios históricos: ${totals.recordatorios}\nTotal envíos históricos: ${totals.envios}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Proceso Completado', `Envíos en esta corrida: ${sent}\nTotal envíos históricos: ${totals.envios}\nTotal recordatorios históricos: ${totals.recordatorios}`, ui.ButtonSet.OK);
  }
}

// ======================================================================================
// PLANTILLAS DE CORREO
function getContenidoCorreo(tipo, nombre, link) {
  let asunto = "";
  let cuerpo = "";
  const buttonStyle = "background-color: #D4AF37; color: white; padding: 12px 24px; text-align: center; text-decoration: none; display: inline-block; border-radius: 8px; font-size: 16px; font-weight: bold;";
  
  if (tipo === "inicial") {
    asunto = `${nombre}, el último paso antes de tu futuro profesional.`;
    cuerpo = `
      Hola, ${nombre},<br><br>
      Estás en la recta final de tu carrera universitaria. 🎓<br><br>
      Antes de que cruces el escenario, queremos asegurarnos de que tengas todo lo necesario para arrancar con éxito tu vida profesional.<br><br>
      <strong>Te tomará solo 2 minutos</strong> completar este checklist personalizado que hemos preparado para ti.<br><br>
      <div style="text-align: center; margin: 30px 0;">
        <a href="${link}" style="${buttonStyle}">Iniciar mi Checklist</a>
      </div>
      <br>
      ¡Nos vemos en la sesión de cierre!<br><br>
      <em>Equipo de Desarrollo Profesional</em>
    `;
  } else if (tipo === "recordatorio") {
    asunto = `${nombre}, ¿dudas sobre tu graduación? No te quedes atrás.`;
    cuerpo = `
      Hola, ${nombre},<br><br>
      Sabemos que el final del semestre es una locura, pero no queremos que te pierdas la oportunidad de resolver tus últimas dudas antes de graduarte.<br><br>
      <strong>⏰ Último recordatorio:</strong> Completa tu checklist de graduación y agenda tu sesión de cierre.<br><br>
      <div style="text-align: center; margin: 30px 0;">
        <a href="${link}" style="${buttonStyle}">Completar mi Checklist ahora</a>
      </div>
      <br>
      No dejes pasar esta oportunidad.<br><br>
      <em>Equipo de Desarrollo Profesional</em>
    `;
  }
  
  return {asunto, cuerpo};
}

// ======================================================================================
// CÓDIGO DE LA WEB APP
function doGet(e) {
  const template = HtmlService.createTemplateFromFile("Index");
  
  // Pasar datos de la sesión
  template.sessionTitle = SESSION_TITLE;
  template.sessionDate = SESSION_DATE;
  template.sessionTimeStart = SESSION_TIME_START;
  template.sessionTimeEnd = SESSION_TIME_END;
  template.sessionLocation = SESSION_LOCATION;
  template.sessionDetails = SESSION_DETAILS;
  
  // Obtener datos del estudiante
  const studentId = e.parameter.id;
  if (studentId) {
    const studentData = getStudentData(studentId);
    template.studentName = studentData ? studentData.nombre : "futuro/a EXATEC";
    template.studentId = studentId;
  } else {
    template.studentName = "futuro/a EXATEC";
    template.studentId = "desconocido";
  }
  
  return template.evaluate()
    .setTitle("Checklist de Graduación")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

function getStudentData(matricula) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift(); // Eliminar encabezados
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().trim() === matricula.toString().trim()) {
        return {
          matricula: data[i][0],
          nombre: data[i][3]
        };
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ======================================================================================
// FUNCIÓN MEJORADA PARA GUARDAR RESPUESTAS Y ASISTENCIA
function guardarRespuesta(datos) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);

    // Encabezados esperados (incluye datos de sesión)
    const expected = [
      "Timestamp",
      "Matrícula",
      "Situación Profesional",
      "Meta Principal",
      "Balance de Vida",
      "Nivel de Preparación",
      "Asistencia a Sesión",
      "Fecha Sesión",
      "Hora Inicio",
      "Hora Fin",
      "Ubicación"
    ];

    // Crear o asegurar encabezados
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(expected);
    } else {
      const header = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), expected.length)).getValues()[0];
      let changed = false;
      expected.forEach((h, i) => {
        if (!header[i]) { sheet.getRange(1, i+1).setValue(h); changed = true; }
      });
      if (changed) SpreadsheetApp.flush();
    }

    // Construir fila alineada a los encabezados actuales
    const headerNow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const rowMap = {
      'Timestamp': new Date(),
      'Matrícula': datos.matricula,
      'Situación Profesional': datos.pregunta1,
      'Meta Principal': datos.pregunta2,
      'Balance de Vida': datos.pregunta3,
      'Nivel de Preparación': datos.pregunta4,
      'Asistencia a Sesión': 'Pendiente',
      'Fecha Sesión': SESSION_DATE,
      'Hora Inicio': SESSION_TIME_START,
      'Hora Fin': SESSION_TIME_END,
      'Ubicación': SESSION_LOCATION,
    };
    const rowArr = headerNow.map(h => rowMap[h] !== undefined ? rowMap[h] : '');
    sheet.appendRow(rowArr);
    
    return "Respuestas guardadas correctamente.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ======================================================================================
// FUNCIÓN MEJORADA PARA REGISTRAR DECISIÓN DE ASISTENCIA
function registrarAsistencia(matricula, decision) {
  try {
    // Actualizar en la hoja de respuestas
    const responsesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);
    const responsesData = responsesSheet.getDataRange().getValues();
    
    // Buscar la fila con esta matrícula (empezando desde la fila 2 para saltar encabezados)
    for (let i = 1; i < responsesData.length; i++) {
      if (responsesData[i][1] && responsesData[i][1].toString().trim() === matricula.toString().trim()) {
        // Actualizar columna 7 (Asistencia a Sesión)
        responsesSheet.getRange(i + 1, 7).setValue(decision);
        break;
      }
    }
    
    // También actualizar en la hoja principal (opcional, para mantener compatibilidad)
    const mainSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    
    let asistenciaCol = headers.indexOf("Asistencia a Sesión") + 1;
    if (asistenciaCol === 0) {
      mainSheet.getRange(1, mainSheet.getLastColumn() + 1).setValue("Asistencia a Sesión");
      asistenciaCol = mainSheet.getLastColumn();
    }
    
    const matriculas = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndex = matriculas.indexOf(matricula) + 2;
    
    if (rowIndex > 1) {
      mainSheet.getRange(rowIndex, asistenciaCol).setValue(decision);
    }
    
    return "Decisión de asistencia registrada.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ======================================================================================
// NUEVA LÓGICA: INTEGRACIÓN CON Emailing_Code.gs (HTML y formato)

// Columnas de tracking que agregaremos si faltan
const TRACKING_COLS = ['Último Envío','Último Recordatorio','Envios','Recordatorios'];

/* Eliminada: runCampaignEmailingCode (se mantiene solo enviarCorreosConEnlaces) */
function runCampaignEmailingCode() {
  const { sheet, header, rows } = loadRowsByHeader_();
  ensureTrackingCols_(sheet, header);
  const now = new Date();

  let sent = 0;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const email = (row['Email'] || row['email'] || '').toString().trim();
    const nombre = (row['Nombre'] || row['nombre completo'] || '').toString().trim();
    if (!email) continue;

    if (!passesRecencyFilter_(row)) continue;
    if (!passesFrequency_(row, 'Último Envío', CONFIG.minDaysBetweenSends)) continue;

    const { html, subject } = renderEmailEmailingCode_(row, false);
    MailApp.sendEmail({ to: email, subject, htmlBody: html, name: 'Camino a Tu Graduación' });

    row['Último Envío'] = now;
    row['Envios'] = (toNumber_(row['Envios']) || 0) + 1;
    writeBackRow_(sheet, header, row.__rowIndex, row);
    sent++;
  }
  SpreadsheetApp.getUi().alert(`Envíos realizados: ${sent}`);
}

/* Eliminada: sendRemindersManualEmailingCode (se mantiene solo enviarCorreosConEnlaces) */
function sendRemindersManualEmailingCode() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let reminders = ss.getSheetByName('Reminders');
  if (!reminders) {
    reminders = ss.insertSheet('Reminders');
    reminders.getRange(1,1).setValue('Email');
    SpreadsheetApp.getUi().alert('Hoja "Reminders" creada. Agrega emails en la Columna A.');
    return;
  }
  const emails = reminders.getRange(2,1, Math.max(reminders.getLastRow()-1,0), 1).getValues().map(r=>String(r[0]||'').trim()).filter(Boolean);
  if (!emails.length) { SpreadsheetApp.getUi().alert('No hay emails en "Reminders".'); return; }

  const { sheet, header, rows } = loadRowsByHeader_();
  ensureTrackingCols_(sheet, header);
  const index = Object.create(null);
  rows.forEach((r, idx) => { const key = (r['Email'] || r['email'] || '').toString().trim().toLowerCase(); if (key) index[key] = { r, idx }; });

  const now = new Date();
  let sent = 0;
  for (const em of emails) {
    const rec = index[em.toLowerCase()];
    if (!rec) continue;
    const row = rec.r;
    if (!passesRecencyFilter_(row)) continue;
    if (!passesFrequency_(row, 'Último Recordatorio', CONFIG.minDaysBetweenReminders)) continue;

    const { html, subject } = renderEmailEmailingCode_(row, true);
    MailApp.sendEmail({ to: em, subject, htmlBody: html, name: 'Camino a Tu Graduación' });

    row['Último Recordatorio'] = now;
    row['Recordatorios'] = (toNumber_(row['Recordatorios']) || 0) + 1;
    writeBackRow_(sheet, header, rec.idx + 2, row);
    sent++;
  }
  SpreadsheetApp.getUi().alert(`Recordatorios enviados: ${sent}`);
}

// Render via Emailing_Code.gs generator
function renderEmailEmailingCode_(row, isReminder) {
  const nombre = String(getBySynonyms_(row, ['nombre','nombre completo','alumno','estudiante'])||'').trim();
  const email = String(getBySynonyms_(row, ['email','correo','correo tec','correo institucional'])||'').trim();
  const matricula = String(getBySynonyms_(row, ['matricula','matrícula','id','matricula tec'])||'').trim();
  // Flags recientes preferentes (Si/No) tal como vienen en la hoja
  const rawProp = getBySynonyms_(row, ['propósito de vida reciente','proposito de vida reciente','proposito reciente']);
  const rawMeta = getBySynonyms_(row, ['meta ocupacional reciente','meta reciente']);
  const propYN = (rawProp != null && rawProp !== '') ? normalizeYesNo_(rawProp) : (flagRecentOrByDate_(row, ['proposito de vida reciente','proposito reciente'], ['proposito vida fecha','fecha proposito de vida','proposito de vida fecha']) ? 'Sí' : 'No');
  const metaYN = (rawMeta != null && rawMeta !== '') ? normalizeYesNo_(rawMeta) : (flagRecentOrByDate_(row, ['meta ocupacional reciente','meta reciente'], ['meta ocupacional fecha','fecha meta ocupacional']) ? 'Sí' : 'No');
  const estudiante = {
    'Nombre': nombre,
    'Programa': String(getBySynonyms_(row, ['programa','programa academico','carrera'])||''),
    'Email': email,
    'Inglés (Inv2025)': String(getBySynonyms_(row, ['ingles inv2025','ingles','nivel de ingles'])||''),
    'Horas de SS': toNumber_(getBySynonyms_(row, ['horas de ss','horas ss','servicio social horas','horas servicio social'])) || 0,
    'Propósito de vida reciente': propYN,
    'Meta ocupacional reciente': metaYN,
  };
  const link = WEB_APP_URL ? `${WEB_APP_URL}${matricula ? ('?id=' + encodeURIComponent(matricula)) : ''}` : '#';
  const mentor = String((typeof MENTOR_NAME !== 'undefined' && MENTOR_NAME) ? MENTOR_NAME : (getBySynonyms_(row, ['mentor','mentor nombre'])||''));
  var generator = null;
  if (typeof generarCorreoPersonalizadoSinGradientes === 'function') {
    generator = generarCorreoPersonalizadoSinGradientes;
  } else if (typeof generarCorreoPersonalizadoV2 === 'function') {
    generator = generarCorreoPersonalizadoV2;
  }
  if (!generator) {
    throw new Error('Emailing_Code: no se encontró el generador (generarCorreoPersonalizadoSinGradientes o generarCorreoPersonalizadoV2).');
  }
  var res = generator(estudiante, link, mentor);
  if (!res || !res.html) {
    throw new Error('Emailing_Code: el generador no devolvió HTML. Revisa Emailing_Code.gs');
  }
  const subject = (isReminder ? 'Recordatorio: ' : '') + interpolate_(CONFIG.subjectTemplate, {
    NOMBRE: nombre,
    PENDIENTES: String(res && typeof res.pendientes === 'number' ? res.pendientes : '')
  });
  return { html: res.html, subject };
}

// ===================== Helpers de hoja y filtros =====================
function loadRowsByHeader_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`No se encontró la pestaña '${SHEET_NAME}'.`);
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return { sheet, header: [], rows: [] };
  const header = sheet.getRange(1,1,1,lc).getValues()[0];
  const values = sheet.getRange(2,1,lr-1,lc).getValues();
  const rows = values.map((r, idx) => {
    const obj = {};
    header.forEach((h,i)=>{ obj[h] = r[i]; });
    obj.__rowIndex = idx + 2; // índice 1-based en la hoja
    return obj;
  });
  return { sheet, header, rows };
}

// ===================== Normalización y acceso por sinónimos =====================
function normalizeKey_(s) {
  s = String(s||'').toLowerCase().trim();
  // quitar acentos
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  // simplificar
  s = s.replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
  return s;
}

function getBySynonyms_(row, syns) {
  const map = {};
  for (const k in row) {
    map[normalizeKey_(k)] = k;
  }
  for (const s of syns) {
    const nk = normalizeKey_((s));
    if (map[nk] != null) return row[map[nk]];
  }
  // intentos con variantes comunes
  for (const s of syns) {
    const nk = normalizeKey_(s);
    for (const key in map) {
      if (key.indexOf(nk) !== -1) return row[map[key]];
    }
  }
  return undefined;
}

function flagRecentOrByDate_(row, flagSyns, dateSyns) {
  const flagVal = getBySynonyms_(row, flagSyns);
  const dYes = detectYes_(flagVal);
  const dNo = detectNo_(flagVal);
  if (dYes === true) return true;
  if (dNo === true) return false;
  const dateVal = getBySynonyms_(row, dateSyns);
  return isRecent_(dateVal);
}

function detectYes_(val) {
  if (val == null) return undefined;
  let s = String(val).trim().toLowerCase();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!s) return undefined;
  if (/^(si|yes|true|1)$/i.test(s)) return true;
  // tolerante: empieza con "si"
  if (s.startsWith('si')) return true;
  return undefined;
}

function detectNo_(val) {
  if (val == null) return undefined;
  let s = String(val).trim().toLowerCase();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!s) return undefined;
  if (/^(no|false|0)$/i.test(s)) return true;
  if (s.startsWith('no')) return true;
  return undefined;
}

function normalizeYesNo_(val) {
  const y = detectYes_(val);
  const n = detectNo_(val);
  if (y === true) return 'Sí';
  if (n === true) return 'No';
  return 'No'; // por defecto, tratamos como No si es ambiguo
}

function ensureTrackingCols_(sheet, header) {
  let added = false;
  TRACKING_COLS.forEach(col => {
    if (header.indexOf(col) === -1) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(col);
      header.push(col);
      added = true;
    }
  });
  if (added) SpreadsheetApp.flush();
}

function writeBackRow_(sheet, header, rowIndex1, rowObj) {
  const rowArr = header.map(h => rowObj[h]);
  sheet.getRange(rowIndex1, 1, 1, header.length).setValues([rowArr]);
}

function getTrackingTotals_(sheet) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return { envios: 0, recordatorios: 0 };
  const header = sheet.getRange(1,1,1,lc).getValues()[0];
  const idxEnv = header.indexOf('Envios');
  const idxRec = header.indexOf('Recordatorios');
  let envios = 0, recs = 0;
  if (idxEnv !== -1) {
    const col = sheet.getRange(2, idxEnv+1, lr-1, 1).getValues();
    envios = col.reduce((s,v)=> s + (toNumber_(v[0])||0), 0);
  }
  if (idxRec !== -1) {
    const col = sheet.getRange(2, idxRec+1, lr-1, 1).getValues();
    recs = col.reduce((s,v)=> s + (toNumber_(v[0])||0), 0);
  }
  return { envios, recordatorios: recs };
}

function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  let s = String(value).trim();
  // formatos comunes: DD/MM/YYYY, YYYY-MM-DD
  const m1 = s.match(/^([0-3]?\d)[\/\-]([0-1]?\d)[\/\-](\d{4})$/);
  if (m1) {
    const d = Number(m1[1]);
    const m = Number(m1[2]) - 1;
    const y = Number(m1[3]);
    const dt = new Date(y, m, d);
    return isNaN(dt) ? null : dt;
  }
  const m2 = s.match(/^(\d{4})[\-\/]([0-1]?\d)[\-\/]([0-3]?\d)$/);
  if (m2) {
    const y = Number(m2[1]);
    const m = Number(m2[2]) - 1;
    const d = Number(m2[3]);
    const dt = new Date(y, m, d);
    return isNaN(dt) ? null : dt;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function daysSince_(date) {
  const d = parseDate_(date);
  if (!d) return undefined;
  return Math.floor((Date.now() - d.getTime()) / 86400000);
}
function isRecent_(date) {
  const diff = daysSince_(date);
  if (diff === undefined) return false;
  return diff <= CONFIG.recentWindowDays;
}
function toNumber_(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}

function passesRecencyFilter_(row) {
  const metaRecent = flagRecentOrByDate_(row, ['meta ocupacional reciente','meta reciente'], ['meta ocupacional fecha','fecha meta ocupacional']);
  const propRecent = flagRecentOrByDate_(row, ['proposito de vida reciente','proposito reciente'], ['proposito vida fecha','fecha proposito de vida','proposito de vida fecha']);
  const crmRecent = isRecent_(getBySynonyms_(row, ['entrevista crm fecha','crm fecha','entrevista fecha']));
  if (CONFIG.requireAllRecentToExclude) {
    return !(metaRecent && propRecent && crmRecent);
  }
  return !(metaRecent || propRecent || crmRecent);
}

function passesFrequency_(row, fieldName, minDays) {
  const diff = daysSince_(row[fieldName]);
  return diff === undefined || diff >= minDays;
}

function interpolate_(template, data) {
  return String(template || '').replace(/\{\{\s*(\w+)\s*\}\}/g, function(_, k){ return String(data[k] || ''); });
}

/**
 * GDIL - Sistema de Gestión de Incidencias LIS (PRO) - v0.6a
 * Desarrollado por TM Luis Ferrada Toro
 */

const SHEET_ID = '1eFYquYeKcwvLjHPvLdbKKghY5rWzsxRx9bNjj3Za_Tw'; // <-- REEMPLAZA ESTO
const CORREO_GDIL = 'ldferrada91@gmail.com';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('GDIL LIS Tracker v0.6a')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function registrarLog(usuario, accion, detalle) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Logs');
    if (sheet) sheet.appendRow([new Date(), usuario, accion, detalle]);
  } catch (e) { console.error("Error Log: " + e.message); }
}

function validarLogin(run, clave) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetUsuarios = ss.getSheetByName('Usuarios');
    const sheetConfig = ss.getSheetByName('Configuracion');
    
    if (!sheetUsuarios || !sheetConfig) return { success: false, message: 'Faltan pestañas en Google Sheets.' };

    const usuariosData = sheetUsuarios.getDataRange().getValues();
    if (usuariosData.length <= 1) return { success: false, message: 'La pestaña Usuarios está vacía.' };
    usuariosData.shift(); 
    
    let user = null;
    for (let r of usuariosData) {
      if (String(r[0]).trim() === String(run).trim() && String(r[1]).trim() === String(clave).trim()) {
        user = { run: r[0], nombre: r[2], rol: r[3], email: r[4], telefono: r[5], laboratorio: r[6] || '' };
        break;
      }
    }

    if (user) {
      registrarLog(user.nombre, 'LOGIN_EXITOSO', `Ingreso (Rol: ${user.rol})`);
      const configData = sheetConfig.getDataRange().getValues();
      let config = { laboratorios: [], fases: [], impactos: [] };
      if (configData.length > 1) {
        configData.shift();
        configData.forEach(row => {
          if (row[0] === 'Laboratorio') config.laboratorios.push(row[1]);
          if (row[0] === 'Fase') config.fases.push(row[1]);
          if (row[0] === 'Impacto') config.impactos.push(row[1]);
        });
      }
      let listaUsuarios = usuariosData.map(r => ({ nombre: r[2], email: r[4] }));
      return { success: true, user: user, config: config, usuariosDisponibles: listaUsuarios };
    } else {
      return { success: false, message: 'RUN o clave incorrectos.' };
    }
  } catch (error) { return { success: false, message: 'Fallo: ' + error.message }; }
}

function cambiarClave(run, claveActual, nuevaClave) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Usuarios');
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]).trim() === String(run).trim() && String(data[i][1]).trim() === String(claveActual).trim()) {
        sheet.getRange(i + 1, 2).setValue(nuevaClave);
        registrarLog(data[i][2], 'CAMBIO_CLAVE', 'Actualizó contraseña');
        return {success: true, message: 'Clave actualizada con éxito.'};
      }
    }
    return {success: false, message: 'La clave actual ingresada es incorrecta.'};
  } catch(e) { return {success: false, message: 'Error: ' + e.message}; }
}

function generarSiguienteID() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Casos');
  const hoyStr = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd");
  const data = sheet.getDataRange().getValues();
  let contador = 1;
  for (let i = data.length - 1; i > 0; i--) {
    let idFull = String(data[i][0]);
    if (idFull.includes(hoyStr)) {
      let partes = idFull.split('-');
      if(partes.length > 2) { contador = parseInt(partes[partes.length - 1]) + 1; break; }
    }
  }
  return `GDIL-${hoyStr}-${String(contador).padStart(4, '0')}`;
}

// --- GENERADOR DE PDF MEJORADO ---
function crearBlobPDF(data, id) {
  let colorPri = data.prioridad === 'Alta' ? '#ef4444' : (data.prioridad === 'Media' ? '#f59e0b' : '#22c55e');
  let est = data.estado || 'Recepcionado';

  // Banner de Estado
  let estColor = '#0284c7'; // Azul por defecto (Recepcionado)
  if (est === 'En revisión') estColor = '#8b5cf6'; // Morado
  if (est === 'Postergado') estColor = '#f59e0b'; // Ámbar
  if (est === 'Finalizado') estColor = '#22c55e'; // Verde

  let html = `
    <!DOCTYPE html><html><head><style>
      body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #1e293b; padding: 40px; }
      .header { text-align: center; margin-bottom: 20px; }
      .title { color: #0f172a; font-size: 26px; margin: 0; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }
      .subtitle { color: #64748b; font-size: 13px; margin-top: 5px; text-transform: uppercase; letter-spacing: 2px; }
      .status-banner { background-color: ${estColor}; color: #ffffff; padding: 15px; text-align: center; font-size: 18px; font-weight: bold; border-radius: 8px; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 30px; }
      .section-title { background-color: #f8fafc; color: #0284c7; padding: 10px 15px; font-size: 14px; font-weight: bold; border-left: 4px solid #0284c7; margin: 20px 0 10px 0; }
      .info-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px; }
      .info-table th { text-align: left; padding: 10px; border-bottom: 1px solid #e2e8f0; color: #64748b; width: 30%; font-weight: bold; }
      .info-table td { padding: 10px; border-bottom: 1px solid #e2e8f0; color: #334155; font-weight: bold; }
      .text-box { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 15px; font-size: 13px; line-height: 1.6; white-space: pre-wrap; color: #334155; }
      .footer { margin-top: 50px; text-align: center; font-size: 10px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 20px; }
    </style></head><body>
      <div class="header">
        <h1 class="title">Informe de Reporte de Caso</h1>
        <div class="subtitle">Comité de Gestión de la Información (GDIL) - SSASUR</div>
      </div>

      <div class="status-banner">ESTADO DEL CASO: ${est}</div>

      <div class="section-title">Datos Generales del Ticket</div>
      <table class="info-table">
        <tr><th>N° ID Caso</th><td style="color: #0284c7; font-size: 16px;">${id}</td></tr>
        <tr><th>Fecha y Hora del Evento</th><td>${data.fecha} a las ${data.hora} hrs</td></tr>
        <tr><th>Laboratorio Origen</th><td>${data.laboratorio}</td></tr>
        <tr><th>Fase / Sistema LIS</th><td>${data.fase}</td></tr>
        <tr><th>Profesional Notificador</th><td>${data.usuario_nombre}</td></tr>
        <tr><th>Nivel de Prioridad</th><td style="color: ${colorPri};">${data.prioridad}</td></tr>
        <tr><th>Impacto Operativo</th><td>${data.impacto}</td></tr>
      </table>

      <div class="section-title">Descripción de la No Conformidad</div>
      <div class="text-box">${data.descripcion}</div>
  `;

  if(est === 'Finalizado' && data.comentario) {
    html += `
      <div class="section-title">Resolución del Caso</div>
      <table class="info-table"><tr><th>Cerrado por</th><td>${data.admin || 'Admin GDIL'}</td></tr></table>
      <div class="text-box" style="border-left: 4px solid #22c55e;">${data.comentario}</div>
    `;
  }

  html += `<div class="footer">Este informe es generado automáticamente por el Sistema de Gestión de Incidencias LIS de la red SSASUR.<br>Generado el: ${Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm")}</div></body></html>`;
  
  return Utilities.newBlob(html, MimeType.HTML).setName(`Reporte_${id}.pdf`).getAs(MimeType.PDF);
}

function enviarCorreoNotificacion(caso, id, motivo) {
  const blob = crearBlobPDF(caso, id);
  const subject = `[${caso.prioridad}] [${id}] - ${motivo}`;
  const body = `Estimado equipo,\n\nSe adjunta el Informe de Reporte de Caso actualizado correspondiente al ticket ${id}.\nMotivo del correo: ${motivo}\n\nAtentamente,\nSistema GDIL SSASUR.`;

  MailApp.sendEmail({ to: CORREO_GDIL, subject: subject, body: body, attachments: [blob] });
  if (caso.Email_Notificador && caso.Email_Notificador.includes('@')) {
    MailApp.sendEmail({ to: caso.Email_Notificador, subject: subject, body: body, attachments: [blob] });
  }
}

// --- GESTIÓN DE CASOS ---
function registrarCaso(data, currentUser) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Casos');
  const timestamp = new Date();
  const id = generarSiguienteID(); 
  
  sheet.appendRow([
    id, data.fecha, data.hora, data.laboratorio, data.fase, 
    data.usuario_nombre, data.usuario_email, data.impacto, data.descripcion, 
    data.prioridad, "Recepcionado", timestamp, timestamp, "", ""
  ]);
  
  let casoData = { ...data, estado: 'Recepcionado', Email_Notificador: data.usuario_email };
  enviarCorreoNotificacion(casoData, id, "Registro de Nuevo Evento");
  return { success: true, id: id };
}

function actualizarEstado(id, nuevoEstado, currentUser) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Casos');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 11).setValue(nuevoEstado);
      sheet.getRange(i + 1, 13).setValue(new Date());
      registrarLog(currentUser.nombre, 'CAMBIO_ESTADO', `Ticket ${id} -> ${nuevoEstado}`);
      
      let casoData = {
        fecha: Utilities.formatDate(new Date(data[i][1]), "GMT-3", "dd/MM/yyyy"), hora: data[i][2],
        laboratorio: data[i][3], fase: data[i][4], usuario_nombre: data[i][5], 
        Email_Notificador: data[i][6], impacto: data[i][7], descripcion: data[i][8], 
        prioridad: data[i][9], estado: nuevoEstado, comentario: String(data[i][13] || ''), admin: String(data[i][14] || '')
      };
      enviarCorreoNotificacion(casoData, id, `Actualización de Estado: ${nuevoEstado}`);
      return true;
    }
  }
  return false;
}

function finalizarCaso(id, comentario, currentUser) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Casos');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 11).setValue('Finalizado');
      sheet.getRange(i + 1, 13).setValue(new Date());
      sheet.getRange(i + 1, 14).setValue(comentario); 
      sheet.getRange(i + 1, 15).setValue(currentUser.nombre); 
      registrarLog(currentUser.nombre, 'FINALIZO_CASO', `Ticket ${id}`);
      
      let casoData = {
        fecha: Utilities.formatDate(new Date(data[i][1]), "GMT-3", "dd/MM/yyyy"), hora: data[i][2],
        laboratorio: data[i][3], fase: data[i][4], usuario_nombre: data[i][5], 
        Email_Notificador: data[i][6], impacto: data[i][7], descripcion: data[i][8], 
        prioridad: data[i][9], estado: 'Finalizado', comentario: comentario, admin: currentUser.nombre
      };
      enviarCorreoNotificacion(casoData, id, "Caso Finalizado y Resuelto");
      return true;
    }
  }
  return false;
}

// --- HISTORIAL ---
function getCasos(user) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Casos');
    if(!sheet) return { casos: [], metrics: {} };
    
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { casos: [], metrics: {hoy: 0, atrasados: 0, finalizados: 0, labMax: "N/A"} };
    
    data.shift(); 
    const hoy = new Date();
    const hoyStr = Utilities.formatDate(hoy, "GMT-3", "yyyy-MM-dd");
    let casos = [];
    let counts = { hoy: 0, atrasados: 0, finalizados: 0, labs: {} };
    
    const formatearSeguro = (valor, esHora) => {
      if (valor instanceof Date) return Utilities.formatDate(valor, "GMT-3", esHora ? "HH:mm" : "dd/MM/yyyy");
      return String(valor || '').trim();
    };
    
    data.forEach(r => {
      if(!r[0]) return;
      
      let rLab = String(r[3] || '').trim();
      let rUserLab = String(user.laboratorio || '').trim();
      
      if (user.rol !== 'Admin' && rLab !== rUserLab) return;
      
      let fCreacion = r[11]; let fMod = r[12];
      let dateTarget = hoy;
      if (fMod && fMod instanceof Date) dateTarget = fMod;
      else if (fCreacion && fCreacion instanceof Date) dateTarget = fCreacion;
      
      let dias = Math.floor((hoy - dateTarget) / (1000 * 60 * 60 * 24));
      let estado = String(r[10] || '').trim();
      let fechaCreacionReal = (fCreacion && fCreacion instanceof Date) ? Utilities.formatDate(fCreacion, "GMT-3", "yyyy-MM-dd") : "";

      if (fechaCreacionReal === hoyStr) counts.hoy++;
      if (estado === 'Finalizado') counts.finalizados++;
      else {
        if ((estado === 'Recepcionado' && dias >= 5) || (estado === 'En revisión' && dias >= 5) || (estado === 'Postergado' && dias >= 10)) counts.atrasados++;
      }
      counts.labs[rLab] = (counts.labs[rLab] || 0) + 1;

      casos.push({
        ID: String(r[0] || ''), Fecha: formatearSeguro(r[1], false), Hora: formatearSeguro(r[2], true), 
        Laboratorio: rLab, Fase: String(r[4] || '').trim(), Notificador: String(r[5] || '').trim(), 
        Impacto: String(r[7] || ''), Descripcion: String(r[8] || ''), Prioridad: String(r[9] || ''), 
        Estado: estado, Dias_Estado: dias, Comentario: String(r[13] || ''), AdminRes: String(r[14] || '')
      });
    });
    
    let labMax = "N/A"; let maxNum = 0;
    for (let l in counts.labs) { if (counts.labs[l] > maxNum) { maxNum = counts.labs[l]; labMax = l; } }
    
    return { casos: casos.reverse(), metrics: { hoy: counts.hoy, atrasados: counts.atrasados, finalizados: counts.finalizados, labMax: labMax } };
  } catch (error) { throw new Error("Error en getCasos: " + error.message); }
}

function generarPDFIndividual(id, userStr) {
  const userObj = JSON.parse(userStr);
  const dataAll = getCasos(userObj); 
  const caso = dataAll.casos.find(c => c.ID === id);
  if (!caso) return null;
  
  let dataParaPDF = {
    prioridad: caso.Prioridad, fecha: caso.Fecha, hora: caso.Hora, laboratorio: caso.Laboratorio, 
    fase: caso.Fase, usuario_nombre: caso.Notificador, impacto: caso.Impacto, 
    descripcion: caso.Descripcion, estado: caso.Estado, comentario: caso.Comentario, admin: caso.AdminRes
  };
  
  const blob = crearBlobPDF(dataParaPDF, id);
  return Utilities.base64Encode(blob.getBytes());
}

// --- MANTENEDOR CONFIG (Edit & Reorder) ---
function obtenerDatosConfiguracion() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const confSheet = ss.getSheetByName('Configuracion').getDataRange().getValues();
  const usrSheet = ss.getSheetByName('Usuarios').getDataRange().getValues();
  confSheet.shift(); usrSheet.shift();
  let config = { Laboratorio: [], Fase: [], Impacto: [], Usuario: [] };
  confSheet.forEach(r => { if (r[0] && r[1]) { if(!config[r[0]]) config[r[0]] = []; config[r[0]].push(r[1]); } });
  usrSheet.forEach(r => { if(r[0]) config.Usuario.push({ run: r[0], clave: r[1], nombre: r[2], rol: r[3], email: r[4], telefono: r[5], laboratorio: r[6] }); });
  return config;
}

function gestionarConfig(accion, tipo, datos, currentUser) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(tipo === 'Usuario' ? 'Usuarios' : 'Configuracion');
  const sheetData = sheet.getDataRange().getValues();
  
  if (accion === 'AGREGAR') {
    if (tipo === 'Usuario') sheet.appendRow([datos.run, datos.clave, datos.nombre, datos.rol, datos.email, datos.telefono, datos.laboratorio]);
    else sheet.appendRow([tipo, datos.valor]);
    return { success: true };
  }
  if (accion === 'ELIMINAR') {
    for (let i = 1; i < sheetData.length; i++) {
      if (tipo === 'Usuario' && String(sheetData[i][0]) === String(datos.id)) { sheet.deleteRow(i + 1); return { success: true }; }
      else if (tipo !== 'Usuario' && sheetData[i][0] === tipo && String(sheetData[i][1]) === String(datos.id)) { sheet.deleteRow(i + 1); return { success: true }; }
    }
  }
  if (accion === 'EDITAR') {
    for (let i = 1; i < sheetData.length; i++) {
      if (tipo === 'Usuario' && String(sheetData[i][0]) === String(datos.id)) { 
        sheet.getRange(i + 1, 3).setValue(datos.nombre);
        sheet.getRange(i + 1, 4).setValue(datos.rol);
        sheet.getRange(i + 1, 7).setValue(datos.laboratorio);
        return { success: true }; 
      }
      else if (tipo !== 'Usuario' && sheetData[i][0] === tipo && String(sheetData[i][1]) === String(datos.id)) { 
        sheet.getRange(i + 1, 2).setValue(datos.nuevo); return { success: true }; 
      }
    }
  }
  if (accion === 'MOVER' && tipo !== 'Usuario') {
    let rowsOfType = [];
    for(let i=1; i<sheetData.length; i++){ if(sheetData[i][0] === tipo) rowsOfType.push(i); }
    let currIdx = -1;
    for(let j=0; j<rowsOfType.length; j++){ if(String(sheetData[rowsOfType[j]][1]) === String(datos.id)){ currIdx = j; break; } }
    
    if(currIdx !== -1) {
      let swapIdx = currIdx + datos.dir;
      if(swapIdx >= 0 && swapIdx < rowsOfType.length) {
         let r1 = rowsOfType[currIdx] + 1; let r2 = rowsOfType[swapIdx] + 1;
         let val1 = sheet.getRange(r1, 2).getValue(); let val2 = sheet.getRange(r2, 2).getValue();
         sheet.getRange(r1, 2).setValue(val2); sheet.getRange(r2, 2).setValue(val1);
         return {success: true};
      }
    }
    return {success: false, message: 'Límite alcanzado'};
  }
  return { success: false, message: 'No encontrado o acción no válida' };
}

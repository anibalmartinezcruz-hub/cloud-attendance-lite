function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Gestión de Asistencia Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function incluir(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- UTILIDAD DE HASHING (SEGURIDAD) ---
function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let hash = "";
  for (i = 0; i < digest.length; i++) {
    let byte = digest[i];
    if (byte < 0) byte += 256;
    let byteStr = byte.toString(16);
    if (byteStr.length == 1) byteStr = "0" + byteStr;
    hash += byteStr;
  }
  return hash;
}

// LOGIN
function login(email, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Usuarios");
    const datos = sheet.getDataRange().getValues();
    
    // Buscamos el usuario en la tabla
    for (let i = 1; i < datos.length; i++) {
      // Asumiendo: Col B (índice 1) es Email y Col C (índice 2) es Password
      // COMPATIBILIDAD: Probamos primero con Hash, si no coincide, probamos texto plano (para migración suave)
      const passStored = String(datos[i][2]);
      const passInputHash = hashPassword(password);
      
      if (datos[i][1] === email && (passStored === passInputHash || passStored === String(password))) {
        const nombre = datos[i][3]; // Col D: Nombre
        const rol = datos[i][4];    // Col E: Rol

        // --- REGISTRO DE AUDITORÍA: ÉXITO ---
        // Acción: "Acceso", Módulo: "Seguridad", Detalles: Nombre y Rol
        registrarLog(datos[i][1], "Acceso", "Seguridad", `Inicio de sesión exitoso: ${nombre} (${rol})`);
        
        return { success: true, email: datos[i][1], nombre: nombre, rol: rol };
      }
    }
    
    // --- REGISTRO DE AUDITORÍA: FALLO ---
    registrarLog(email, "Intento Fallido", "Seguridad", `Credenciales incorrectas para el correo: ${email}`);
    
    return { success: false, message: "Usuario o contraseña incorrectos" };

  } catch (e) {
    // Registro de error técnico en auditoría
    registrarLog("Sistema", "Error Crítico", "Seguridad", `Falla en función login: ${e.toString()}`);
    return { success: false, message: "Error de conexión con la base de datos." };
  }
}

// OBTENER EVENTOS (Con conteo de inscritos y horarios)
function getEventos(userEmail, rol) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE = ss.getSheetByName("Eventos");
    const sheetP = ss.getSheetByName("Participantes");
    
    if (!sheetE) throw new Error("No se encontró la hoja 'Eventos'");
    const eventosData = sheetE.getDataRange().getValues();
    const participantesData = sheetP ? sheetP.getDataRange().getValues() : [];
    eventosData.shift(); // Quitar encabezados

    // Filtramos y validamos cada fila
    const filtrados = eventosData.filter(fila => {
      if (!fila[0] || !fila[4]) return false; 
      
      const emailResponsable = fila[4].toString().toLowerCase().trim();
      const emailUsuario = userEmail ? userEmail.toLowerCase().trim() : "";

      // Visibilidad por Rol
      if (rol === "root" || rol === "admin") return true;
      return emailResponsable === emailUsuario;
    });
    
    return filtrados.map(f => {
      const idEvento = f[0].toString();
      const cupoMax = parseInt(f[5]) || 0;
      
      const inscritos = participantesData.length > 0 ? 
                        participantesData.filter(p => p[3] && p[3].toString() === idEvento).length : 0;
      
      const fechaInicio = (f[2] instanceof Date) ? f[2] : new Date();
      const fechaFin = (f[6] instanceof Date) ? f[6] : new Date();

      return { 
        id: idEvento, 
        nombre: f[1] || "Sin título", 
        fecha: fechaInicio.toLocaleDateString('es-ES'),
        horaInicio: fechaInicio.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}),
        horaFin: fechaFin.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}),
        isoInicio: Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"),
        isoFin: Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"),
        tipo: f[3] || "General",
        cupoMax: cupoMax,
        inscritos: inscritos,
        responsable: f[4] || ""
      };
    });
  } catch (e) {
    Logger.log("Error crítico en getEventos: " + e.message);
    return [];
  }
}

// CRUD EVENTOS (Solo Admin/Root)
function crearEventoServidor(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetE = ss.getSheetByName("Eventos");
  const nuevoId = "E-" + (sheetE.getLastRow() + 1);
  // --- VALIDACIONES SERVIDOR ---
  if (datos.cupo < 1) return { success: false, message: "Error: El cupo debe ser mayor a 0." };
  if (new Date(datos.inicio) >= new Date(datos.fin)) return { success: false, message: "Error: La fecha de inicio debe ser anterior a la de fin." };

  sheetE.appendRow([nuevoId, datos.nombre, new Date(datos.inicio), datos.tipo, datos.responsable.toLowerCase().trim(), datos.cupo, new Date(datos.fin)]);
  return { success: true, message: "Evento creado exitosamente" };
}

function editarEventoServidor(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetE = ss.getSheetByName("Eventos");
  const data = sheetE.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === datos.id.toString()) {
      // --- VALIDACIONES SERVIDOR ---
      if (datos.cupo < 1) return { success: false, message: "Error: El cupo debe ser mayor a 0." };
      if (new Date(datos.inicio) >= new Date(datos.fin)) return { success: false, message: "Error: La fecha de inicio debe ser anterior a la de fin." };
      
      const fila = i + 1;
      const rowData = data[i];
      
      // Detectar cambios auditables
      let cambios = [];
      
      // Helper para formatear valor
      const fmt = (v) => (v instanceof Date) ? v.toISOString().split('T')[0] : v;

      // Comparaciones
      if(fmt(rowData[1]) != fmt(datos.nombre)) cambios.push(`Nombre: ${rowData[1]} -> ${datos.nombre}`);
      // Fechas (Checkeamos solo fecha YYYY-MM-DD o tiempo si es crítico, aqui simplifico)
      // Nota: rowData[2] viene de sheet como Date. new Date(datos.inicio) es Date.
      if(Math.abs(new Date(rowData[2]) - new Date(datos.inicio)) > 60000) cambios.push("Fecha Inicio");
      if(rowData[3] != datos.tipo) cambios.push(`Tipo: ${rowData[3]} -> ${datos.tipo}`);
      if(rowData[4].toString().toLowerCase() != datos.responsable.toString().toLowerCase()) cambios.push(`Responsable: ${rowData[4]} -> ${datos.responsable}`);
      if(rowData[5] != datos.cupo) cambios.push(`Cupo: ${rowData[5]} -> ${datos.cupo}`);

      // Actualizar
      sheetE.getRange(fila, 2, 1, 6).setValues([[datos.nombre, new Date(datos.inicio), datos.tipo, datos.responsable.toLowerCase().trim(), datos.cupo, new Date(datos.fin)]]);
      
      // Log auditoría detallado
      if (cambios.length > 0) {
         const actor = datos.usuarioActor || "Sistema";
         registrarLog(actor, "Edición Detallada", "Eventos", `ID ${datos.id}. Cambios: ${cambios.join(', ')}`);
      }

      return { success: true, message: "Evento actualizado y auditado." };
    }
  }
  return { success: false, message: "No encontrado" };
}

function eliminarEventoServidor(idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetE = ss.getSheetByName("Eventos");
  const data = sheetE.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idEvento.toString()) {
      sheetE.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// --- INSCRIPCIÓN Y ASISTENCIA (CORREGIDO: VALIDACIÓN DE CUPO) ---
function agregarParticipante(nuevo, userRol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName("Participantes");
  const sheetE = ss.getSheetByName("Eventos");
  
  if (!sheetP || !sheetE) return { success: false, message: "Error: No se encontraron las hojas." };
  
  const eventosData = sheetE.getDataRange().getValues();
  const participantesData = sheetP.getDataRange().getValues();
  
  const correoLimpio = nuevo.correo.toLowerCase().trim();
  const idEventoNuevo = nuevo.idEvento.toString();

  // 1. VALIDACIÓN DE DUPLICADOS
  const yaInscritoEnMismo = participantesData.some(p => 
    p[2].toString().toLowerCase().trim() === correoLimpio && 
    p[3].toString() === idEventoNuevo
  );
  if (yaInscritoEnMismo) return { success: false, message: "El alumno ya está inscrito en este taller." };

  // 2. BUSCAR EVENTO
  const evNuevo = eventosData.find(f => f[0].toString() === idEventoNuevo);
  if (!evNuevo) return { success: false, message: "Evento no encontrado." };

  // --- 2.5 VALIDACIÓN DE CUPO (NUEVO) ---
  const cupoMax = parseInt(evNuevo[5]); // Columna F es índice 5 (0,1,2,3,4,5)
  // Contamos cuántos hay inscritos actualmente en ese ID de evento
  const totalInscritos = participantesData.filter(p => p[3].toString() === idEventoNuevo).length;

  if (totalInscritos >= cupoMax) {
     return { success: false, message: "Error: El taller ha alcanzado su cupo máximo (" + cupoMax + ")." };
  }
  // ----------------------------------------

  // 3. VALIDACIÓN DE CHOQUE DE HORARIOS
  const inicioNuevo = new Date(evNuevo[2]).getTime();
  const finNuevo = new Date(evNuevo[6]).getTime();

  const misOtrasInscripciones = participantesData.filter(p => p[2].toString().toLowerCase().trim() === correoLimpio);
  for (let registro of misOtrasInscripciones) {
    const infoEvInscrito = eventosData.find(e => e[0].toString() === registro[3].toString());
    if (infoEvInscrito) {
      const inicioExistente = new Date(infoEvInscrito[2]).getTime();
      const finExistente = new Date(infoEvInscrito[6]).getTime();
      // Si se solapan
      if (inicioNuevo < finExistente && finNuevo > inicioExistente) {
        return { success: false, message: `Choque de horario con: ${infoEvInscrito[1]}` };
      }
    }
  }

  // 4. REGISTRO Y ENVÍO
 try {
    const nuevoIdP = "P-" + (sheetP.getLastRow() + 1);
    const nombreAlumno = nuevo.nombre.trim();
    const nombreTaller = evNuevo[1];
    
    sheetP.appendRow([nuevoIdP, nombreAlumno, correoLimpio, idEventoNuevo]);

    const fechaTaller = new Date(evNuevo[2]).toLocaleDateString('es-ES', { weekday: 'long', day: 'numeric', month: 'long' });
    const horaTaller = new Date(evNuevo[2]).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    // --- GENERACIÓN QR ---
    // Formato: ID_PARTICIPANTE | ID_EVENTO
    const qrData = `${nuevoIdP}|${idEventoNuevo}`;
    const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(qrData)}&size=300&dark=0f172a&ecLevel=H`;

    enviarCorreoConfirmacion(correoLimpio, nombreAlumno, nombreTaller, fechaTaller, horaTaller, qrUrl);
    
    // Log corregido: Usuario actor vs Alumno inscrito
    const actor = nuevo.usuarioActor || correoLimpio; // Si no hay actor (registro público?), usamos el alumno
    registrarLog(actor, "Inscripción", "Participantes", `Inscribió al alumno: ${correoLimpio} en: ${nombreTaller}`);

    return { success: true, message: "Inscripción exitosa. QR Enviado." };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}


// --- FUNCIÓN MODIFICADA: DETECTA SI LA ASISTENCIA DEBE BLOQUEARSE ---
// --- COPIA ESTO EN gs.txt (Reemplaza la función anterior) ---

function getParticipantesConEstado(idEvento, fechaConsulta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pData = ss.getSheetByName("Participantes").getDataRange().getValues();
  const aData = ss.getSheetByName("Asistencia").getDataRange().getValues();
  pData.shift(); aData.shift();

  const hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fechaFiltro = fechaConsulta ? fechaConsulta : hoyStr;

  // Verificar si la fecha está bloqueada
  const hayRegistrosEseDia = aData.some(a => {
    if (a[1].toString() !== idEvento.toString()) return false;
    let fechaRegistro = (a[3] instanceof Date) ? Utilities.formatDate(a[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : a[3].toString().substring(0, 10);
    return fechaRegistro === fechaFiltro;
  });
  
  const asistenciaCerrada = (fechaFiltro !== hoyStr && hayRegistrosEseDia);

  const lista = pData.filter(p => p[3].toString() === idEvento.toString()).map(p => {
    const idParticipante = p[0].toString();
    
    // Buscar registro específico de ese día
    const registroHoy = aData.find(a => {
      if (a[1].toString() !== idEvento.toString()) return false;
      if (a[2].toString() !== idParticipante) return false;      
      let fechaRegistro = (a[3] instanceof Date) ? Utilities.formatDate(a[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : a[3].toString().substring(0, 10);
      return fechaRegistro === fechaFiltro;
    });

    // AQUÍ ESTÁ LA CLAVE: Si no hay registro, enviamos "Ausente" explícitamente
    const estadoHoy = registroHoy ? registroHoy[4] : "Ausente";

    const totalAsistencias = aData.filter(a => 
      a[1].toString() === idEvento.toString() && 
      a[2].toString() === idParticipante &&
      (a[4] === "Presente" || a[4] === "Justificado")
    ).length;

    return { 
      id: idParticipante, 
      nombre: p[1], 
      estado: estadoHoy, // Esto es lo que faltaba o estaba fallando
      conteo: totalAsistencias 
    };
  });

  return { participantes: lista, bloqueado: asistenciaCerrada };
}

function registrarAsistencia(idEvento, asistencias, userEmail, fechaSesion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Asistencia");
  const aData = sheet.getDataRange().getValues();
  
  // Ajuste de fecha para asegurar que se guarde correctamente
  const fechaGuardar = fechaSesion ? new Date(fechaSesion + "T12:00:00") : new Date();
  const fechaString = Utilities.formatDate(fechaGuardar, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 1. BORRAR REGISTROS ANTERIORES DE ESE DÍA (Limpieza)
  // Recorremos de abajo hacia arriba para no afectar los índices al borrar
  for (let i = aData.length - 1; i >= 1; i--) {
    const row = aData[i];
    let rowDate = "";
    if (row[3] instanceof Date) {
      rowDate = Utilities.formatDate(row[3], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rowDate = row[3].toString().substring(0, 10);
    }

    if (row[1].toString() === idEvento.toString() && rowDate === fechaString) {
      sheet.deleteRow(i + 1);
    }
  }

  // 2. INSERTAR NUEVOS REGISTROS
  asistencias.forEach(reg => {
    // ⚠️ AQUÍ ESTÁ EL CAMBIO IMPORTANTE ⚠️
    // Debemos permitir guardar si es "Presente" O si es "Justificado"
    if(reg.estado === "Presente" || reg.estado === "Justificado") {
       
       sheet.appendRow([
         Utilities.getUuid(), 
         idEvento, 
         reg.idParticipante, 
         fechaGuardar, 
         reg.estado, // Esto guardará la palabra correcta ("Presente" o "Justificado")
         userEmail
       ]);
    }
  });

  return { success: true, message: "Asistencia actualizada correctamente." };
}

function getMiHorario(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE = ss.getSheetByName("Eventos");
    const sheetP = ss.getSheetByName("Participantes");
    
    const eData = sheetE.getDataRange().getValues();
    const pData = sheetP ? sheetP.getDataRange().getValues() : [];
    eData.shift(); 
    if(pData.length > 0) pData.shift();

    const emailUser = email.toLowerCase().trim();
    let misTalleres = [];
    
    // DOCENTE
    const talleresComoDocente = eData.filter(f => f[4] && f[4].toString().toLowerCase().trim() === emailUser);
    talleresComoDocente.forEach(f => {
      misTalleres.push({
        id: f[0],
        nombre: f[1],
        inicio: f[2] instanceof Date ? f[2].toLocaleString('es-ES', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' }) : "Sin fecha",
        fin: f[6] instanceof Date ? f[6].toLocaleString('es-ES', { hour:'2-digit', minute:'2-digit' }) : "Sin fecha",
        inicioISO: f[2] instanceof Date ? f[2].toISOString() : null,
        finISO: f[6] instanceof Date ? f[6].toISOString() : null,
        tipo: f[3] + " (Instructor)"
      });
    });

    // ALUMNO
    const inscripciones = pData.filter(p => p[2] && p[2].toString().toLowerCase().trim() === emailUser);
    inscripciones.forEach(ins => {
      const info = eData.find(e => e[0].toString() === ins[3].toString());
      if (info) {
        misTalleres.push({
          id: info[0],
          nombre: info[1],
          inicio: info[2] instanceof Date ? info[2].toLocaleString('es-ES', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' }) : "Sin fecha",
          fin: info[6] instanceof Date ? info[6].toLocaleString('es-ES', { hour:'2-digit', minute:'2-digit' }) : "Sin fecha",
          inicioISO: info[2] instanceof Date ? info[2].toISOString() : null,
          finISO: info[6] instanceof Date ? info[6].toISOString() : null,
          tipo: info[3] + " (Alumno)"
        });
      }
    });
    return misTalleres;
  } catch (e) {
    console.log("Error en getMiHorario: " + e.message);
    return [];
  }
}

// OBTENER TODOS LOS USUARIOS
function getUsuariosServidor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  return data.map((f, index) => ({
    id: index + 2,
    email: f[1],
    pass: f[2],
    nombre: f[3],
    rol: f[4]
  }));
}

// CREAR O EDITAR USUARIO
function guardarUsuarioServidor(u) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  const data = sheet.getDataRange().getValues();
  
  // Hash de la contraseña
  const passHash = hashPassword(u.pass);
  
  let filaEncontrada = -1;
  // Buscamos si el usuario ya existe
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().toLowerCase() === u.email.toLowerCase()) {
      filaEncontrada = i + 1;
      break;
    }
  }

  if (filaEncontrada !== -1) {
    // ACTUALIZAR USUARIO
    const rowData = data[filaEncontrada - 1];
    let cambios = [];
    
    if(rowData[3] != u.nombre) cambios.push(`Nombre: ${rowData[3]} -> ${u.nombre}`);
    if(rowData[4] != u.rol) cambios.push(`Rol: ${rowData[4]} -> ${u.rol}`);
    if(rowData[2] != passHash) cambios.push("Contraseña: (Cambiada)");

    sheet.getRange(filaEncontrada, 3, 1, 3).setValues([[passHash, u.nombre, u.rol]]);
    
    // Log
    if(cambios.length > 0) {
        const actor = u.usuarioActor || "Sistema";
        registrarLog(actor, "Edición Detallada", "Usuarios", `Usuario ${u.email}. Cambios: ${cambios.join(', ')}`);
    }

    return { success: true, message: "Usuario actualizado (Seguro)" };
  } else {
    // CREAR NUEVO USUARIO
    const nuevoId = "U-" + (sheet.getLastRow() + 1);
    sheet.appendRow([nuevoId, u.email, passHash, u.nombre, u.rol]);
    return { success: true, message: "Usuario creado (Seguro)" };
  }
}

// ELIMINAR USUARIO
function eliminarUsuarioServidor(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// AUDITORÍA
function registrarLog(usuario, accion, modulo, detalles) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Auditoria");
  if (!sheet) {
    sheet = ss.insertSheet("Auditoria");
    sheet.appendRow(["Fecha y Hora", "Usuario", "Acción", "Módulo", "Detalles"]);
  }
  const userMail = usuario || "Sistema/Desconocido";
  sheet.appendRow([new Date(), userMail, accion, modulo, detalles]);
}

function getLogsServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Auditoria");
    if (!sheet) sheet = ss.getSheetByName("Auditoría"); // Fallback con tilde
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    data.shift();

    const logsFormateados = data.reverse().map(fila => {
      let fechaTexto = "";
      if (fila[0] instanceof Date) {
        fechaTexto = Utilities.formatDate(fila[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else {
        fechaTexto = fila[0].toString(); 
      }
      return [fechaTexto, fila[1] || "", fila[2] || "", fila[3] || "", fila[4] || ""];
    });
    return logsFormateados.slice(0, 100); 
  } catch (e) {
    return [];
  }
}

// EMAIL
function enviarCorreoConfirmacion(email, nombre, taller, fecha, hora, qrUrl) {
  const asunto = `🎟️ Tu Acceso QR: ${taller}`;
  const cuerpoHtml = `
    <div style="font-family: sans-serif; border: 1px solid #e2e8f0; border-radius: 10px; padding: 20px; max-width: 600px; color: #1e293b; text-align: center;">
      <h2 style="color: #2563eb;">¡Hola, ${nombre}!</h2>
      <p>Tu inscripción está confirmada. Presenta este código al llegar.</p>
      
      <div style="margin: 20px auto;">
        <img src="${qrUrl}" alt="Código QR de Acceso" style="width: 200px; height: 200px; border: 4px solid #334155; border-radius: 10px;">
      </div>

      <div style="background: #f8fafc; padding: 15px; border-radius: 8px; margin: 20px 0; text-align: left;">
        <p style="margin: 5px 0;"><strong>📍 Taller:</strong> ${taller}</p>
        <p style="margin: 5px 0;"><strong>📅 Fecha:</strong> ${fecha}</p>
        <p style="margin: 5px 0;"><strong>⏰ Hora:</strong> ${hora}</p>
      </div>
      <p style="font-size: 0.8rem; color: #64748b;">No respondas a este correo automático.</p>
    </div>
  `;
  MailApp.sendEmail({
    to: email,
    subject: asunto,
    htmlBody: cuerpoHtml
  });
}

function getEstadisticasAvanzadas(filtroDocente, filtroInicio, filtroFin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener datos crudos
  const sheetE = ss.getSheetByName("Eventos");
  const sheetP = ss.getSheetByName("Participantes");
  const sheetA = ss.getSheetByName("Asistencia");

  const eData = sheetE ? sheetE.getDataRange().getValues() : [];
  const pData = sheetP ? sheetP.getDataRange().getValues() : [];
  const aData = sheetA ? sheetA.getDataRange().getValues() : [];

  // Quitar encabezados
  if(eData.length > 0) eData.shift();
  if(pData.length > 0) pData.shift();
  if(aData.length > 0) aData.shift();

  // 2. Preparar Mapas de Conteo
  // Mapa: ID_Evento -> Cantidad Inscritos
  const inscritosMap = {};
  pData.forEach(p => {
    const idEv = p[3]; // Col D es ID Evento
    inscritosMap[idEv] = (inscritosMap[idEv] || 0) + 1;
  });

  // Mapa: ID_Evento -> Cantidad Asistentes (Solo "Presente")
  const asistenciaMap = {};
  aData.forEach(a => {
    // Col B es ID Evento, Col E es Estado
    if (a[4] === "Presente") {
      const idEv = a[1];
      asistenciaMap[idEv] = (asistenciaMap[idEv] || 0) + 1;
    }
  });

  // 3. Filtrar Eventos
  let eventosFiltrados = eData.filter(e => {
    const id = e[0];
    const fecha = new Date(e[2]);
    const docente = e[4].toString().toLowerCase();

    // Filtro Docente
    if (filtroDocente && filtroDocente !== "" && docente !== filtroDocente.toLowerCase()) {
      return false;
    }

    // Filtro Fechas
    if (filtroInicio) {
      const fInicio = new Date(filtroInicio);
      if (fecha < fInicio) return false; // Ajustar lógica según necesidad exacta de horas
    }
    if (filtroFin) {
      const fFin = new Date(filtroFin);
      // Ajustamos fFin al final del día para incluir eventos de ese día
      fFin.setHours(23,59,59);
      if (fecha > fFin) return false;
    }

    return true;
  });

  // 4. Construir Datos para Gráficas
  const labels = [];
  const dataAsistieron = []; // Verde
  const dataAusentes = [];   // Rojo/Gris (Inscritos - Asistieron)
  const dataCupoLibre = [];  // (Cupo - Inscritos)

  let totalCupos = 0;
  let totalInscritos = 0;
  let totalAsistencias = 0;

  eventosFiltrados.forEach(e => {
    const id = e[0];
    const nombre = e[1];
    const cupoMax = parseInt(e[5]) || 0;
    
    const inscritos = inscritosMap[id] || 0;
    const asistieron = asistenciaMap[id] || 0;
    const ausentes = Math.max(0, inscritos - asistieron); // Los que se inscribieron pero no fueron

    labels.push(nombre);
    dataAsistieron.push(asistieron);
    dataAusentes.push(ausentes);
    
    totalCupos += cupoMax;
    totalInscritos += inscritos;
    totalAsistencias += asistieron;
  });

  // Lista única de docentes para el filtro (dropdown)
  const docentesUnicos = [...new Set(eData.map(e => e[4]))];

  return {
    labels: labels,
    asistieron: dataAsistieron,
    ausentes: dataAusentes,
    kpis: {
      totalTalleres: eventosFiltrados.length,
      totalInscritos: totalInscritos,
      tasaAsistencia: totalInscritos > 0 ? Math.round((totalAsistencias / totalInscritos) * 100) : 0,
      ocupacionGlobal: totalCupos > 0 ? Math.round((totalInscritos / totalCupos) * 100) : 0
    },
    listaDocentes: docentesUnicos
  };
}
// --- REPORTE DE MATRIZ DE ASISTENCIA ---
function getReporteMatriz(idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName("Participantes");
  const sheetA = ss.getSheetByName("Asistencia");

  const pData = sheetP.getDataRange().getValues();
  const aData = sheetA.getDataRange().getValues();
  
  // Quitar headers
  pData.shift(); 
  aData.shift();

  // 1. Obtener alumnos del evento
  const alumnos = pData
    .filter(p => p[3].toString() === idEvento.toString())
    .map(p => ({ id: p[0].toString(), nombre: p[1] }));

  if (alumnos.length === 0) return { fechas: [], data: [] };

  // 2. Obtener todas las fechas únicas registradas para este evento
  const registrosEvento = aData.filter(r => r[1].toString() === idEvento.toString());
  const fechasSet = new Set();
  
  registrosEvento.forEach(r => {
    let fechaStr = "";
    if (r[3] instanceof Date) {
      fechaStr = Utilities.formatDate(r[3], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      fechaStr = r[3].toString().substring(0, 10);
    }
    fechasSet.add(fechaStr);
  });

  // Ordenar fechas cronológicamente
  const fechasOrdenadas = Array.from(fechasSet).sort();

  // 3. Construir la matriz
  const reporte = alumnos.map(alumno => {
    const asistenciaPorFecha = {};
    
    fechasOrdenadas.forEach(fecha => {
      // Buscar si existe registro "Presente" para este alumno en esta fecha
      const asistio = registrosEvento.some(r => {
        let fRegistro = (r[3] instanceof Date) ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[3].toString().substring(0, 10);
        return r[2].toString() === alumno.id && fRegistro === fecha && r[4] === "Presente";
      });
      asistenciaPorFecha[fecha] = asistio;
    });

    return {
      nombre: alumno.nombre,
      asistencias: asistenciaPorFecha
    };
  });

  return {
    fechas: fechasOrdenadas, // Array ["2024-01-01", "2024-01-02"]
    data: reporte            // Array de alumnos con sus marcas
  };
}
// --- PROCESAMIENTO DE QR DESDE EL ESCÁNER ---
function procesarQRServidor(codigoQR, idEventoActual, emailDocente) {
  try {
    // El formato esperado es "ID_PARTICIPANTE|ID_EVENTO"
    const partes = codigoQR.split("|");
    if (partes.length < 2) return { success: false, message: "❌ Formato de QR inválido" };

    const idParticipante = partes[0].trim();
    const idEventoQR = partes[1].trim();

    // 1. Validar que el QR corresponda al evento que se está gestionando
    if (idEventoQR !== idEventoActual.toString()) {
      return { success: false, message: "⚠️ Este QR pertenece a otro evento." };
    }

    // 2. Registrar Asistencia
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetA = ss.getSheetByName("Asistencia");
    
    // Verificar si ya asistió HOY
    const data = sheetA.getDataRange().getValues();
    const hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const yaRegistro = data.some(row => {
      let rowDate = (row[3] instanceof Date) ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[3].toString().substring(0,10);
      return row[1].toString() === idEventoQR && row[2].toString() === idParticipante && rowDate === hoyStr;
    });

    if (yaRegistro) {
      return { success: true, message: "✅ El alumno ya fue registrado hoy." };
    }

    // Insertar
    sheetA.appendRow([
      Utilities.getUuid(),
      idEventoQR,
      idParticipante,
      new Date(), // Fecha y hora actual del escaneo
      "Presente",
      emailDocente + " (QR SCAN)"
    ]);

    return { success: true, message: "✅ Asistencia Registrada Correctamente" };

  } catch (e) {
    return { success: false, message: "❌ Error: " + e.message };
  }
}
// --- CARGA MASIVA DE PARTICIPANTES ---
function procesarCargaMasiva(idEvento, listaAlumnos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName("Participantes");
  const sheetE = ss.getSheetByName("Eventos");
  
  // 1. Obtener datos del Evento para validaciones
  const eventosData = sheetE.getDataRange().getValues();
  const evInfo = eventosData.find(e => e[0].toString() === idEvento.toString());
  
  if (!evInfo) return { success: false, message: "Evento no encontrado." };
  
  const nombreTaller = evInfo[1];
  const fechaTaller = new Date(evInfo[2]).toLocaleDateString();
  const cupoMax = parseInt(evInfo[5]);

  // 2. Obtener inscritos actuales para validar duplicados y cupo
  const partData = sheetP.getDataRange().getValues();
  const inscritosActuales = partData.filter(p => p[3].toString() === idEvento.toString());
  
  if (inscritosActuales.length >= cupoMax) {
    return { success: false, message: "El taller ya está lleno." };
  }

  let agregados = 0;
  let duplicados = 0;
  const filasNuevas = [];
  const correosParaEnviar = []; // Para enviar emails después

  // 3. Procesar lista
  // Calculamos cupo disponible
  let cupoRestante = cupoMax - inscritosActuales.length;
  const ultimaFila = sheetP.getLastRow();

  listaAlumnos.forEach((alumno, index) => {
    if (cupoRestante <= 0) return; // Si se llena a mitad de carga, ignora el resto

    const nombre = alumno.nombre.trim();
    const correo = alumno.correo.toLowerCase().trim();

    // Validar si ya existe en este evento
    const yaExiste = partData.some(p => 
      p[2].toString().toLowerCase() === correo && 
      p[3].toString() === idEvento.toString()
    );

    // Validar si ya está en la lista de carga actual (duplicado en el mismo excel)
    const duplicadoEnLote = filasNuevas.some(f => f[2] === correo);

    if (!yaExiste && !duplicadoEnLote) {
      const nuevoId = "P-" + (ultimaFila + index + 1) + Math.floor(Math.random() * 1000);
      
      // Preparar fila para guardar de golpe
      filasNuevas.push([nuevoId, nombre, correo, idEvento]);
      
      // Preparar datos para correo QR
      const qrData = `${nuevoId}|${idEvento}`;
      const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(qrData)}&size=300&dark=0f172a&ecLevel=H`;
      
      correosParaEnviar.push({ email: correo, nombre: nombre, qr: qrUrl });
      
      agregados++;
      cupoRestante--;
    } else {
      duplicados++;
    }
  });

  // 4. Guardar en Bloque (Mucho más rápido que appendRow uno por uno)
  if (filasNuevas.length > 0) {
    sheetP.getRange(ultimaFila + 1, 1, filasNuevas.length, 4).setValues(filasNuevas);
  }

  // 5. Enviar Correos (Opcional: Si son muchos, esto puede tardar. Se hace un try/catch silencioso)
  // Nota: Google limita los correos diarios. Para cargas de >50 personas, ten cuidado.
  try {
    correosParaEnviar.forEach(c => {
       enviarCorreoConfirmacion(c.email, c.nombre, nombreTaller, fechaTaller, "Ver Detalles", c.qr);
    });
  } catch(e) {
    console.log("Error enviando correos masivos: " + e.message);
  }

  return { 
    success: true, 
    message: `Proceso finalizado.<br>✅ Agregados: ${agregados}<br>⚠️ Duplicados/Omitidos: ${duplicados}<br>📦 Cupo restante: ${cupoRestante}` 
  };
}
// --- GENERADOR DE DIPLOMAS AUTOMÁTICO ---



// --- FUNCIÓN PARA PROBARLO MANUALMENTE ---
// Ejecuta esta función desde el editor para ver si te llega el correo
function testDiploma() {
  generarDiplomaPDF("Juan Pérez Test", "Curso de Prueba", "9.5", "djgar2005@gmail.com");
}
// --- MÓDULO DE CALIFICACIONES Y DIPLOMAS (LMS) ---

// 1. CREAR UNA NUEVA TAREA
function crearTarea(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Actividades");
  if(!sheet) return { success: false, message: "Falta hoja Actividades" };
  
  // --- VALIDACIONES SERVIDOR ---
  if (datos.peso < 0) return { success: false, message: "Error: El peso de la actividad no puede ser negativo." };

  const id = "ACT-" + (sheet.getLastRow() + 1) + Math.floor(Math.random()*100);
  sheet.appendRow([id, datos.idEvento, datos.titulo, datos.desc, datos.peso]);
  
  return { success: true, message: "Actividad creada correctamente" };
}

// 2. GUARDAR UNA NOTA (En tiempo real)
function guardarNota(idActividad, idParticipante, nota) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Calificaciones");
  if(!sheet) sheet = ss.insertSheet("Calificaciones"); // Prevenir error

  // --- VALIDACIONES SERVIDOR ---
  if (nota < 0) return { success: false, message: "Error: La nota no puede ser negativa." };
  
  const data = sheet.getDataRange().getValues();
  // Buscar si ya existe la nota para actualizarla
  for(let i=0; i<data.length; i++) {
    if(data[i][1] == idActividad && data[i][2] == idParticipante) {
      sheet.getRange(i+1, 4).setValue(nota); // Actualizar Columna D (Puntaje)
      return { success: true };
    }
  }
  // Si no existe, crear nueva fila
  const idNota = "NOT-" + Date.now();
  sheet.appendRow([idNota, idActividad, idParticipante, nota, new Date()]);
  return { success: true };
}

// 1. FUNCIÓN PARA EL LIBRO DE CALIFICACIONES (CORREGIDA: ENVÍA 'nombre' y 'curso')
function getLibroCalificacionesV2(idEvento) {
  var respuesta = { actividades: [], alumnos: [], curso: "Curso" }; 
  try {
    if (!idEvento) return respuesta;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shPart = ss.getSheetByName("Participantes");
    const shAsist = ss.getSheetByName("Asistencia");
    const shAct = ss.getSheetByName("Actividades");
    const shCalif = ss.getSheetByName("Calificaciones");
    const shEventos = ss.getSheetByName("Eventos");

    // A. OBTENER NOMBRE DEL CURSO REAL
    if (shEventos) {
       const datosEventos = shEventos.getDataRange().getValues();
       const eventoEncontrado = datosEventos.find(e => e[0].toString() === idEvento.toString());
       if (eventoEncontrado) respuesta.curso = eventoEncontrado[1]; 
    }

    if (!shPart || !shAsist) return respuesta;

    // B. OBTENER ALUMNOS
    const rawP = shPart.getDataRange().getValues(); rawP.shift();
    const alumnos = rawP.filter(p => p[3].toString() === idEvento.toString()).map(p => {
      let fechaDiploma = "";
      if (p.length > 4 && p[4]) fechaDiploma = p[4].toString(); 
      return { id: p[0], nombre: p[1], email: p[2], diplomaEnviado: fechaDiploma };
    });

    // C. OBTENER ACTIVIDADES (CORRECCIÓN 'UNDEFINED': Usamos 'nombre')
    let actividades = [];
    if (shAct) {
      const rawA = shAct.getDataRange().getValues();
      if (rawA.length > 0 && rawA[0][0] === "ID_Actividad") rawA.shift();
      actividades = rawA.filter(a => a[1].toString() === idEvento.toString()).map(a => ({
        id: a[0], 
        nombre: a[2], // <--- AQUÍ ESTABA EL ERROR, AHORA DICE 'nombre'
        peso: a[4] 
      }));
    }

    // D. CALCULAR NOTAS (Igual que antes)
    const rawN = shCalif ? shCalif.getDataRange().getValues() : [];
    const rawAsist = shAsist.getDataRange().getValues(); rawAsist.shift();
    const asistenciasEvento = rawAsist.filter(r => r[1].toString() === idEvento.toString());
    
    // ... Cálculo de asistencia ...
    const fechasUnicas = new Set();
    asistenciasEvento.forEach(r => {
      let fStr = (r[3] instanceof Date) ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[3].toString().substring(0, 10);
      fechasUnicas.add(fStr);
    });
    const totalSesiones = fechasUnicas.size || 1;
    const asistMap = {};
    alumnos.forEach(alum => {
      const conteo = asistenciasEvento.filter(r => r[2].toString() === alum.id.toString() && (r[4] === "Presente" || r[4] === "Justificado")).length;
      asistMap[alum.id] = Math.round((conteo / totalSesiones) * 100);
    });

    const gradebook = alumnos.map(alum => {
      let sumaPonderada = 0;
      let notaRecuperacion = null;
      const notasAlumno = {};

      actividades.forEach(act => {
        const notaRow = rawN.find(n => n[1].toString() === act.id.toString() && n[2].toString() === alum.id.toString());
        const val = notaRow ? parseFloat(notaRow[3]) : 0;
        notasAlumno[act.id] = val;

        if (act.peso == 0) {
           if (notaRow && val > 0) notaRecuperacion = val;
        } else {
           if (notaRow) sumaPonderada += val * (act.peso / 100);
        }
      });

      let notaFinal = actividades.length > 0 ? parseFloat(sumaPonderada.toFixed(1)) : 0;
      if (notaRecuperacion !== null) { notaFinal = notaRecuperacion; if(notaFinal > 10) notaFinal = 10; }

      const asistencia = asistMap[alum.id] || 0;
      let estado = asistencia < 70 ? "SIN_DERECHO" : (notaFinal >= 7.0 ? "APROBADO" : "RECUPERACION");

      return {
        id: alum.id, nombre: alum.nombre, email: alum.email,
        diplomaEnviado: alum.diplomaEnviado, asistencia: asistencia,
        notas: notasAlumno, promedio: notaFinal, estado: estado
      };
    });

    respuesta.actividades = actividades;
    respuesta.alumnos = gradebook;
    return respuesta;

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    return respuesta;
  }
}

// 4. GENERAR Y ENVIAR DIPLOMA (CORREGIDA PARA USAR V2)
function procesarDiplomaBackend(idEvento, idParticipante, forzarReenvio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Validaciones
  // ⚠️ AQUÍ ESTABA EL ERROR: Cambiamos a ...V2
  const libro = getLibroCalificacionesV2(idEvento); 
  
  const alumno = libro.alumnos.find(a => a.id.toString() === idParticipante.toString());
  
  if(!alumno) return { success: false, message: "Alumno no encontrado" };
  if(alumno.estado !== "APROBADO") return { success: false, message: "El alumno no cumple los requisitos." };

  // 2. CHEQUEO: ¿Ya se envió antes?
  if (alumno.diplomaEnviado && !forzarReenvio) {
    return { 
      success: false, 
      yaEnviado: true, 
      fecha: alumno.diplomaEnviado,
      message: "Este diploma ya fue enviado anteriormente." 
    };
  }

  // 3. Generar PDF
  const ev = ss.getSheetByName("Eventos").getDataRange().getValues().find(e => e[0].toString() === idEvento.toString());
  const nombreCurso = ev ? ev[1] : "Taller";
  
  const resultado = generarDiplomaPDF(alumno.nombre, nombreCurso, alumno.promedio, alumno.email);

  if (resultado.success) {
    // 4. MARCAR EN LA BASE DE DATOS
    const shPart = ss.getSheetByName("Participantes");
    const data = shPart.getDataRange().getValues();
    
    for(let i=0; i<data.length; i++) {
      if(data[i][0].toString() === idParticipante.toString()) {
        shPart.getRange(i+1, 5).setValue(new Date()); // Escribir fecha en Columna E
        break;
      }
    }
  }

  return resultado;
}

// --- GENERADOR DE PDF (Tu función anterior, integrada) ---
function generarDiplomaPDF(nombreAlumno, nombreCurso, notaFinal, emailAlumno) {
  // 🔽🔽 PEGA AQUÍ TU ID DE GOOGLE SLIDES 🔽🔽
  const ID_PLANTILLA = "1UwxbrCQJe992sMOI3CFCJJGzrTU2FA-LcH9BIfQH7-M"; 
  // 🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼🔼

  try {
    const archivoPlantilla = DriveApp.getFileById(ID_PLANTILLA);
    const copiaArchivo = archivoPlantilla.makeCopy("Diploma Temp");
    const idCopia = copiaArchivo.getId();
    const presentacion = SlidesApp.openById(idCopia);
    
    // Reemplazar textos
    presentacion.replaceAllText("{{Nombre}}", nombreAlumno); // Ajusta si en tu slide pusiste {{NombreAlumno}}
    presentacion.replaceAllText("{{Curso}}", nombreCurso);
    presentacion.replaceAllText("{{Nota}}", notaFinal.toString());
    presentacion.replaceAllText("{{Fecha}}", new Date().toLocaleDateString());
    
    presentacion.saveAndClose();
    
    const pdfBlob = copiaArchivo.getAs(MimeType.PDF);
    pdfBlob.setName("Diploma - " + nombreAlumno + ".pdf");
    
    // Enviar Email
    MailApp.sendEmail({
      to: emailAlumno,
      subject: "🎓 Tu Diploma: " + nombreCurso,
      body: `Hola ${nombreAlumno},\n\nFelicidades por aprobar con nota ${notaFinal}.\nAdjunto tu diploma.\n\nAtte, La Dirección.`,
      attachments: [pdfBlob]
    });
    
    copiaArchivo.setTrashed(true); // Borrar temporal
    return { success: true, message: "Diploma enviado a " + emailAlumno };
    
  } catch (e) {
    return { success: false, message: "Error PDF: " + e.message };
  }
}
// --- KARDEX DEL ALUMNO (CON DETALLE DE NOTAS Y ASISTENCIA) ---
function obtenerKardexAlumno(idParticipanteActual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const shE = ss.getSheetByName("Eventos");
  const shA = ss.getSheetByName("Asistencia");
  const shAct = ss.getSheetByName("Actividades");
  const shCalif = ss.getSheetByName("Calificaciones");

  const dataP = shP.getDataRange().getValues();
  const alumnoActualRow = dataP.find(r => r[0].toString() === idParticipanteActual.toString());
  
  if (!alumnoActualRow) return { success: false, message: "Alumno no encontrado." };
  
  const emailAlumno = alumnoActualRow[2].toString().toLowerCase().trim();
  const nombreAlumno = alumnoActualRow[1];
  const inscripciones = dataP.filter(r => r[2].toString().toLowerCase().trim() === emailAlumno);

  const dataE = shE.getDataRange().getValues();
  const dataA = shA ? shA.getDataRange().getValues() : [];
  const dataAct = shAct ? shAct.getDataRange().getValues() : [];
  const dataCalif = shCalif ? shCalif.getDataRange().getValues() : [];

  const historial = inscripciones.map(insc => {
    const idInscripcion = insc[0];
    const idEvento = insc[3];
    
    // A) Datos del Evento
    const evento = dataE.find(e => e[0].toString() === idEvento.toString());
    const nombreCurso = evento ? evento[1] : "Evento Eliminado";
    const fecha = evento ? new Date(evento[2]).toLocaleDateString() : "-";

    // B) Asistencia (Cálculo + Detalle)
    const asistenciasEvento = dataA.filter(r => r[1].toString() === idEvento.toString());
    const fechasUnicas = new Set(asistenciasEvento.map(r => r[3].toString().substring(0,10)));
    const totalSesiones = fechasUnicas.size || 1;
    
    // Detalle de asistencias para este alumno
    const detalleAsistencia = asistenciasEvento
      .filter(r => r[2].toString() === idInscripcion.toString())
      .map(r => ({
         fecha: (r[3] instanceof Date) ? r[3].toLocaleDateString() : r[3].toString().substring(0,10),
         estado: r[4]
      }));

    // Contar asistencias válidas
    const misAsistencias = detalleAsistencia.filter(r => r.estado === "Presente" || r.estado === "Justificado").length;
    const porcAsistencia = Math.round((misAsistencias / totalSesiones) * 100);
    
    // C) Notas y Desglose
    const actividadesEvento = dataAct.filter(a => a[1].toString() === idEvento.toString());
    let sumaNotas = 0;
    const desgloseNotas = []; 
    
    actividadesEvento.forEach(act => {
       const idAct = act[0];
       const calif = dataCalif.find(c => c[1].toString() === idAct.toString() && c[2].toString() === idInscripcion.toString());
       const valor = calif ? parseFloat(calif[3]) : 0;
       
       if(act[4] > 0) sumaNotas += valor * (act[4] / 100);

       desgloseNotas.push({ actividad: act[2], nota: valor, peso: act[4] });
    });
    
    const notaFinal = parseFloat(sumaNotas.toFixed(1));

    // D) Estado
    let estado = "REPROBADO";
    if (porcAsistencia >= 70 && notaFinal >= 7.0) estado = "APROBADO";
    else if (porcAsistencia >= 70 && notaFinal < 7.0) estado = "RECUPERACION";
    else if (porcAsistencia < 70) estado = "SIN DERECHO";

    return {
      idEvento: idEvento,
      curso: nombreCurso,
      fecha: fecha,
      nota: notaFinal,
      asistencia: porcAsistencia,
      estado: estado,
      desglose: desgloseNotas,       // Detalle Notas
      historialAsistencia: detalleAsistencia // Detalle Asistencia (NUEVO)
    };
  });
  
  // Estadísticas globales
  const cursosAprobados = historial.filter(h => h.estado === "APROBADO").length;
  const promedioGeneral = historial.length > 0 ? (historial.reduce((sum, h) => sum + h.nota, 0) / historial.length).toFixed(1) : 0;

  return { 
    success: true, 
    perfil: {
      nombre: nombreAlumno,
      email: emailAlumno,
      totalCursos: historial.length,
      aprobados: cursosAprobados,
      promedio: promedioGeneral
    },
    historial: historial 
  };
}
// --- FUNCIÓN PARA MATRIZ DE ASISTENCIA (HISTORIAL) ---
function obtenerMatrizAsistencia(idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAsist = ss.getSheetByName("Asistencia");
  const shPart = ss.getSheetByName("Participantes");
  
  // 1. Obtener Participantes del Evento
  const rawP = shPart.getDataRange().getValues(); rawP.shift();
  const participantes = rawP.filter(p => p[3].toString() === idEvento.toString()).map(p => ({
    id: p[0], nombre: p[1]
  }));
  
  // 2. Obtener Asistencias del Evento
  const rawA = shAsist.getDataRange().getValues(); rawA.shift();
  const registros = rawA.filter(r => r[1].toString() === idEvento.toString());
  
  // 3. Obtener Fechas Únicas (Columnas de la tabla)
  const fechasSet = new Set();
  registros.forEach(r => {
    let fechaStr = "";
    if (r[3] instanceof Date) fechaStr = Utilities.formatDate(r[3], Session.getScriptTimeZone(), "dd/MM");
    else fechaStr = r[3].toString().substring(0, 10); // Ajusta si usas otro formato
    fechasSet.add(fechaStr);
  });
  // Ordenar fechas (opcional, pero recomendado)
  const fechasOrdenadas = Array.from(fechasSet).sort();

  // 4. Cruzar Datos (Matriz)
  const reporte = participantes.map(p => {
    let asistenciasCount = 0;
    const historial = {};
    
    fechasOrdenadas.forEach(fechaCol => {
      // Buscar registro de este alumno en esta fecha
      const encontrado = registros.find(r => {
        let fReg = (r[3] instanceof Date) ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "dd/MM") : r[3].toString().substring(0, 10);
        return r[2].toString() === p.id.toString() && fReg === fechaCol;
      });
      
      const estado = encontrado ? encontrado[4] : "Ausente"; // Col 4 es el estado
      historial[fechaCol] = estado;
      
      // 🧠 LÓGICA CORREGIDA: Sumar si es Presente O Justificado
      if (estado === "Presente" || estado === "Justificado") {
        asistenciasCount++;
      }
    });
    
    // Calcular Porcentaje
    const totalSesiones = fechasOrdenadas.length;
    const porcentaje = totalSesiones > 0 ? Math.round((asistenciasCount / totalSesiones) * 100) : 0;
    
    return {
      nombre: p.nombre,
      historial: historial,
      porcentaje: porcentaje
    };
  });
  
  return { fechas: fechasOrdenadas, datos: reporte };
}
// --- CRUD DE PARTICIPANTES (NUEVO MÓDULO) ---

function getTodosLosParticipantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const shE = ss.getSheetByName("Eventos");
  
  if(!shP) return [];

  const dataP = shP.getDataRange().getValues();
  const dataE = shE.getDataRange().getValues();
  
  // Quitar encabezados
  dataP.shift(); 
  
  // Mapear para devolver objetos limpios
  return dataP.map(p => {
    // Buscar nombre del evento para mostrarlo en la tabla en lugar del ID
    const evento = dataE.find(e => e[0].toString() === p[3].toString());
    const nombreEvento = evento ? evento[1] : "Evento Eliminado (" + p[3] + ")";
    
    return {
      id: p[0],           // ID Inscripción
      nombre: p[1],       // Nombre Alumno
      email: p[2],        // Correo
      idEvento: p[3],     // ID Evento
      nombreEvento: nombreEvento
    };
  }).reverse(); // Mostramos los más nuevos primero
}

function editarParticipanteDirecto(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const data = shP.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === datos.id.toString()) {
      // Actualizamos Nombre y Correo (Columnas 2 y 3 -> índices 1 y 2)
      shP.getRange(i+1, 2, 1, 2).setValues([[datos.nombre, datos.email]]);
      return { success: true, message: "Datos actualizados correctamente." };
    }
  }
  return { success: false, message: "Participante no encontrado." };
}

function eliminarParticipanteDirecto(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const data = shP.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === id.toString()) {
      shP.deleteRow(i+1);
      return { success: true, message: "Inscripción eliminada." };
    }
  }
  return { success: false, message: "No se pudo eliminar." };
}
// --- OBTENER DATOS PARA EL LIBRO DE CALIFICACIONES (GRADEBOOK) ---
function getGradebookData(idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAct = ss.getSheetByName("Actividades");
  const shP = ss.getSheetByName("Participantes");
  const shC = ss.getSheetByName("Calificaciones");

  // Validar que existan las hojas
  if (!shAct || !shP || !shC) {
    return { actividades: [], alumnos: [] };
  }

  const dataAct = shAct.getDataRange().getValues();
  const dataP = shP.getDataRange().getValues();
  const dataC = shC.getDataRange().getValues();

  // 1. Obtener Actividades del Evento
  // Asumimos: Col 0: ID_Act, Col 1: ID_Evento, Col 2: Nombre, Col 4: Peso
  const actividades = dataAct
    .filter(r => r[1].toString() === idEvento.toString())
    .map(r => ({
      id: r[0],
      nombre: r[2],
      peso: r[4]
    }));

  // Si no hay actividades, devolvemos vacío
  if (actividades.length === 0) {
    return { actividades: [], alumnos: [] };
  }

  // 2. Obtener Alumnos inscritos en este evento
  // Asumimos: Col 0: ID_Inscripcion, Col 1: Nombre, Col 3: ID_Evento
  const alumnosInscritos = dataP.filter(r => r[3].toString() === idEvento.toString());

  // 3. Cruzar Alumnos con Calificaciones
  const reporteAlumnos = alumnosInscritos.map(p => {
    const idInscripcion = p[0];
    const nombre = p[1];
    let sumaNotas = 0;
    
    // Para cada actividad, buscamos si el alumno tiene nota
    const notas = actividades.map(act => {
      // Calificaciones: Col 1: ID_Act, Col 2: ID_Inscripcion, Col 3: Nota
      const registro = dataC.find(c => 
        c[1].toString() === act.id.toString() && 
        c[2].toString() === idInscripcion.toString()
      );
      
      const valor = registro ? parseFloat(registro[3]) : 0;
      
      // Sumar al promedio ponderado
      if (act.peso > 0) {
        sumaNotas += valor * (act.peso / 100);
      }
      
      // Retornar "-" si no tiene nota o es 0, para que se vea limpio en la tabla
      return registro ? valor : "-"; 
    });

    const promedioFinal = parseFloat(sumaNotas.toFixed(1));

    return {
      nombre: nombre,
      notas: notas,
      promedio: promedioFinal
    };
  });

  return {
    actividades: actividades,
    alumnos: reporteAlumnos
  };
}
function DIAGNOSTICO_TOTAL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shE = ss.getSheetByName("Eventos");
  const shP = ss.getSheetByName("Participantes");
  const shAct = ss.getSheetByName("Actividades");
  const shCal = ss.getSheetByName("Calificaciones");

  Logger.log("--- INICIO DIAGNÓSTICO ---");
  
  // 1. Verificar Eventos
  if(!shE) { Logger.log("❌ ERROR CRÍTICO: No existe hoja Eventos"); return; }
  const dataE = shE.getDataRange().getValues();
  // Asumimos que fila 1 es header, fila 2 es el primer evento
  if(dataE.length < 2) { Logger.log("⚠️ Hoja Eventos vacía"); return; }
  
  const idEvento = dataE[1][0]; // Tomamos el primer evento real (Ej: EV-001)
  Logger.log(`🔎 Analizando Evento ID: "${idEvento}" Nombre: "${dataE[1][1]}"`);

  // 2. Verificar Actividades
  if(shAct) {
    const acts = shAct.getDataRange().getValues().filter(r => r[1].toString() === idEvento.toString());
    Logger.log(`📘 Actividades encontradas para este evento: ${acts.length}`);
    if(acts.length === 0) Logger.log("⚠️ AVISO: No hay actividades creadas. La tabla saldrá sin columnas de notas.");
  } else {
    Logger.log("❌ ERROR: No existe hoja Actividades");
  }

  // 3. Verificar Participantes
  if(shP) {
    // Importante: Chequeamos espacios en blanco con .trim()
    const parts = shP.getDataRange().getValues().filter(r => r[3].toString().trim() === idEvento.toString().trim());
    Logger.log(`👥 Participantes encontrados para este evento: ${parts.length}`);
    
    if(parts.length === 0) {
      Logger.log("⚠️ AVISO: 0 Participantes. Posible error: Los IDs no coinciden exactamente (Revisa espacios extra).");
      // Imprimir los primeros 3 IDs de participantes para comparar
      const rawP = shP.getDataRange().getValues();
      Logger.log(`   Ejemplo de ID en hoja Participantes: "${rawP[1][3]}" vs Evento: "${idEvento}"`);
    }
  } else {
    Logger.log("❌ ERROR: No existe hoja Participantes");
  }
  
  Logger.log("--- FIN DIAGNÓSTICO ---");
}
// --- FUNCIÓN PARA DIBUJAR LAS TARJETAS (Actualizada: Botón de Notas para Docentes) ---
function mostrarEventos(lista) {
  const container = document.getElementById('cards-container');
  if(!container) return; // Evita errores si no existe el contenedor
  
  container.innerHTML = ""; 

  if (!lista || lista.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; width:100%; margin-top:50px; opacity:0.7;">
         <div style="font-size:4rem;">📭</div>
         <p style="color:white; font-size:1.2rem;">No tienes cursos asignados.</p>
      </div>`;
    return;
  }

  lista.forEach(ev => {
    // 1. Cálculos de progreso
    const total = ev.cupoTotal || 1;
    const inscritos = ev.inscritos || 0;
    const porcentaje = Math.round((inscritos / total) * 100);
    const libres = total - inscritos;
    
    let colorBarra = "#3b82f6"; // Azul
    if (porcentaje >= 80) colorBarra = "#f59e0b"; // Naranja
    if (porcentaje >= 100) colorBarra = "#ef4444"; // Rojo

    // 2. DEFINIR BOTONES SEGÚN EL ROL
    let botonesAccion = "";
    
    // Estilo común para los botones flotantes (Fondo oscuro semitransparente)
    const estiloBotones = "position:absolute; top:15px; right:15px; display:flex; gap:5px; background:rgba(15, 23, 42, 0.9); padding:5px; border-radius:8px; z-index:10; box-shadow: 0 2px 5px rgba(0,0,0,0.3);";

    if (user.rol === "root" || user.rol === "admin") {
       // ADMIN: Ve los 3 botones (Editar, Borrar, Notas)
       botonesAccion = `
         <div style="${estiloBotones}">
            <button onclick="abrirEditarEvento('${ev.id}')" title="Editar" style="cursor:pointer; border:none; background:none; font-size:1.1rem;">✏️</button>
            <button onclick="borrarEvento('${ev.id}')" title="Eliminar" style="cursor:pointer; border:none; background:none; font-size:1.1rem;">🗑️</button>
            <button onclick="irANotas('${ev.id}', '${ev.nombre}')" title="Calificaciones y Tareas" style="cursor:pointer; border:none; background:none; font-size:1.1rem;">📝</button>
         </div>
       `;
    } else if (user.rol === "docente") {
       // DOCENTE: Ve SOLO el botón de Notas (📝) en la misma posición
       botonesAccion = `
         <div style="${estiloBotones}">
            <button onclick="irANotas('${ev.id}', '${ev.nombre}')" title="Gestionar Notas y Tareas" style="cursor:pointer; border:none; background:none; font-size:1.1rem; color:white;">📝</button>
         </div>
       `;
    }

    // 3. CREAR TARJETA
    const card = document.createElement('div');
    card.className = "card animate-in"; 
    card.style.position = "relative"; 
    
    card.innerHTML = `
      ${botonesAccion} <div style="color:#60a5fa; font-size:0.75rem; font-weight:bold; text-transform:uppercase; letter-spacing:1px; margin-bottom:5px;">
         ${ev.tipo || 'Curso'}
      </div>
      
      <h3 style="margin:0 0 10px 0; color:white; font-size:1.2rem; line-height:1.4;">
         ${ev.nombre}
      </h3>
      
      <div style="font-size:0.9rem; color:#94a3b8; margin-bottom:15px; display:flex; flex-direction:column; gap:3px;">
         <span>📅 ${new Date(ev.fecha).toLocaleDateString()}</span>
         <span>⏰ ${ev.horario || 'Por definir'}</span>
      </div>

      <div style="background:#1e293b; height:6px; border-radius:3px; overflow:hidden; margin-bottom:10px;">
         <div style="width:${porcentaje}%; height:100%; background:${colorBarra}; transition:width 0.5s;"></div>
      </div>
      
      <div style="display:flex; justify-content:space-between; font-size:0.8rem; color:#cbd5e1;">
         <span>👥 Inscritos: <b>${inscritos}</b></span>
         <span style="color:${libres === 0 ? '#ef4444' : '#10b981'}">🎟 Libres: <b>${libres}</b></span>
      </div>
    `;

    container.appendChild(card);
  });
}
// --- ENVIAR ANUNCIO MASIVO (NUEVO) ---
function enviarAnuncioCurso(idEvento, asunto, mensaje, docenteEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const shE = ss.getSheetByName("Eventos");

  // 1. Obtener Nombre del Curso
  const dataE = shE.getDataRange().getValues();
  const evento = dataE.find(r => r[0].toString() === idEvento.toString());
  const nombreCurso = evento ? evento[1] : "Curso";

  // 2. Obtener Destinatarios (Filtrar por ID Evento)
  const dataP = shP.getDataRange().getValues();
  const destinatarios = dataP
    .filter(r => r[3].toString() === idEvento.toString()) // Col 3 es ID Evento
    .map(r => ({ nombre: r[1], email: r[2] }));

  if (destinatarios.length === 0) return { success: false, message: "No hay alumnos inscritos en este curso." };

  // 3. Enviar Correos
  let enviados = 0;
  const asuntoFinal = `📢 Aviso: ${nombreCurso} - ${asunto}`;
  
  destinatarios.forEach(dest => {
    if(dest.email && dest.email.includes("@")) {
      try {
        MailApp.sendEmail({
          to: dest.email,
          subject: asuntoFinal,
          htmlBody: `
            <div style="font-family: sans-serif; padding: 20px; color: #333; border: 1px solid #e2e8f0; border-radius: 8px;">
              <h2 style="color: #2563eb; margin-top:0;">${nombreCurso}</h2>
              <p>Hola ${dest.nombre},</p>
              <p>Tu docente <strong>${docenteEmail}</strong> ha publicado el siguiente anuncio:</p>
              
              <div style="background: #f8fafc; padding: 15px; border-left: 5px solid #2563eb; margin: 20px 0; font-size: 1.1rem;">
                ${mensaje.replace(/\n/g, '<br>')}
              </div>
              
              <p style="font-size: 0.8rem; color: #94a3b8; margin-top: 20px;">
                Este es un mensaje automático enviado desde Eventos Manager.
              </p>
            </div>
          `
        });
        enviados++;
      } catch(e) { 
        console.log("Error enviando a " + dest.email + ": " + e.message); 
      }
    }
  });

  // 4. Registrar en Auditoría
  registrarLog(docenteEmail, "Comunicación", "Anuncio Masivo", `Enviado a ${enviados} alumnos de ${nombreCurso}`);

  return { success: true, message: `✅ Anuncio enviado exitosamente a ${enviados} alumnos.` };
}
// --- REPORTE DE INTELIGENCIA ACADÉMICA ---
function obtenerAnalisisCurso(idEvento) {
  // Reutilizamos la lógica del Gradebook para asegurar consistencia
  const datos = getLibroCalificacionesV2(idEvento);
  const alumnos = datos.alumnos;
  
  if(!alumnos || alumnos.length === 0) return { error: true, message: "Sin alumnos." };

  // 1. Detectar Alumnos en Riesgo
  // Criterio: Reprobado, Sin Derecho, Recuperación o Asistencia < 70%
  const enRiesgo = alumnos.filter(a => 
    a.estado === "REPROBADO" || 
    a.estado === "SIN_DERECHO" || 
    a.estado === "RECUPERACION" ||
    a.asistencia < 70
  ).map(a => ({
    nombre: a.nombre,
    motivo: a.asistencia < 70 ? `Falta Asistencia (${a.asistencia}%)` : `Bajo Promedio (${a.promedio})`
  }));

  // 2. Detectar Top 3 Mejores
  const topAlumnos = alumnos
    .filter(a => a.estado === "APROBADO")
    .sort((a, b) => b.promedio - a.promedio) // Ordenar de mayor a menor
    .slice(0, 3) // Solo los 3 primeros
    .map(a => ({ nombre: a.nombre, promedio: a.promedio }));

  // 3. Datos para la Gráfica
  const conteo = {
    aprobados: alumnos.filter(a => a.estado === "APROBADO").length,
    reprobados: alumnos.filter(a => a.estado !== "APROBADO").length
  };

  return {
    success: true,
    riesgo: enRiesgo,
    top: topAlumnos,
    grafica: conteo
  };
}
// --- MÓDULO DE RECUPERACIÓN DE CONTRASEÑA ---

// 1. SOLICITAR CÓDIGO (Genera token y envía correo)
function solicitarRecuperacionPass(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  const data = sheet.getDataRange().getValues();
  
  // A. Verificar y crear columnas si no existen (Auto-Mantenimiento)
  if (data[0].length < 7) {
     sheet.getRange(1, 6).setValue("Token_Recuperacion");
     sheet.getRange(1, 7).setValue("Expiracion_Token");
  }

  const emailLimpio = email.trim().toLowerCase();
  let filaUsuario = -1;
  let nombreUsuario = "";

  // B. Buscar Usuario
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim().toLowerCase() === emailLimpio) {
      filaUsuario = i + 1; // Fila real en Sheet (base 1)
      nombreUsuario = data[i][3];
      break;
    }
  }

  if (filaUsuario === -1) {
    // Por seguridad, no decimos "no existe", decimos "si existe, se envió".
    return { success: true, message: "Si el correo existe, recibirás instrucciones." };
  }

  // C. Generar Token (6 dígitos)
  const token = Math.floor(100000 + Math.random() * 900000).toString();
  
  // D. Calcular Expiración (15 minutos)
  const ahora = new Date();
  const expiracion = new Date(ahora.getTime() + 15 * 60000); // +15 min

  // E. Guardar en Base de Datos (Cols 6 y 7)
  sheet.getRange(filaUsuario, 6).setValue(token);
  sheet.getRange(filaUsuario, 7).setValue(expiracion);

  // F. Enviar Correo
  try {
    MailApp.sendEmail({
      to: emailLimpio,
      subject: "🔐 Recupera tu contraseña - Eventos Manager",
      htmlBody: `
        <div style="font-family:sans-serif; padding:20px; border:1px solid #ddd; border-radius:10px; text-align:center;">
           <h2 style="color:#2563eb;">Código de Recuperación</h2>
           <p>Hola ${nombreUsuario}, usa este código para restablecer tu contraseña:</p>
           <div style="background:#f1f5f9; padding:15px; font-size:2rem; letter-spacing:5px; font-weight:bold; margin:20px 0;">
              ${token}
           </div>
           <p style="color:#666; font-size:0.9rem;">Este código expira en 15 minutos.</p>
        </div>
      `
    });
  } catch(e) {
    return { success: false, message: "Error enviando correo: " + e.message };
  }

  return { success: true, message: "Correo enviado. Revisa tu bandeja." };
}

// 2. VALIDAR Y CAMBIAR CONTRASEÑA
function cambiarPassConToken(email, tokenInput, nuevaPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  const data = sheet.getDataRange().getValues();
  
  const emailLimpio = email.trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim().toLowerCase() === emailLimpio) {
      
      const tokenGuardado = data[i][5];
      const expiracion = new Date(data[i][6]);
      const ahora = new Date();

      // Validaciones
      if (!tokenGuardado || tokenGuardado.toString() !== tokenInput.toString()) {
         return { success: false, message: "Código incorrecto." };
      }
      
      if (ahora > expiracion) {
         return { success: false, message: "El código ha expirado. Pide uno nuevo." };
      }

      // CAMBIO DE CONTRASEÑA (Columna 3 - Password) — CIFRAR CON SHA-256
      const passHasheada = hashPassword(nuevaPass.toString());
      sheet.getRange(i + 1, 3).setNumberFormat('@').setValue(passHasheada);
      
      // Limpiar Token para que no se reuse
      sheet.getRange(i + 1, 6).setValue(""); 
      sheet.getRange(i + 1, 7).setValue("");

      return { success: true, message: "¡Contraseña actualizada exitosamente!" };
    }
  }
  
  return { success: false, message: "Usuario no encontrado." };
}
// --- GESTIÓN DE ACTIVIDADES (EDITAR Y BORRAR) ---

function editarActividadServidor(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Actividades");
  const data = sh.getDataRange().getValues();
  
  // Buscar por ID (Columna 0)
  for(let i=1; i<data.length; i++) {
    if(data[i][0].toString() === datos.id.toString()) {
      // --- VALIDACIONES SERVIDOR ---
      if (datos.peso < 0) return { success: false, message: "Error: El peso de la actividad no puede ser negativo." };

      // Actualizamos: Título (Col 2), Desc (Col 3), Peso (Col 4)
      sh.getRange(i+1, 3, 1, 3).setValues([[datos.titulo, datos.desc, datos.peso]]);
      return { success: true, message: "Actividad actualizada correctamente." };
    }
  }
  return { success: false, message: "Error: Actividad no encontrada." };
}

function eliminarActividadServidor(idActividad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Actividades");
  const shCalif = ss.getSheetByName("Calificaciones"); // Para limpiar notas huérfanas
  
  const data = sh.getDataRange().getValues();
  
  // 1. Borrar la actividad
  for(let i=1; i<data.length; i++) {
    if(data[i][0].toString() === idActividad.toString()) {
      sh.deleteRow(i+1);
      
      // 2. (Opcional) Limpiar calificaciones asociadas para no dejar basura
      if(shCalif) {
        const dataC = shCalif.getDataRange().getValues();
        // Recorremos inversamente para borrar sin afectar índices
        for(let j=dataC.length-1; j>=0; j--) {
           if(dataC[j][1].toString() === idActividad.toString()) {
             shCalif.deleteRow(j+1);
           }
        }
      }
      
      return { success: true, message: "Actividad eliminada." };
    }
  }
  return { success: false, message: "No se pudo eliminar." };
}
// 2. FUNCIÓN PARA ANALYTICS (CORREGIDA: ARREGLA GRÁFICOS VACÍOS)
function getTendenciasCurso(idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAct = ss.getSheetByName("Actividades");
  const shCalif = ss.getSheetByName("Calificaciones");
  const shAsist = ss.getSheetByName("Asistencia");
  const shPart = ss.getSheetByName("Participantes");

  const alumnos = shPart.getDataRange().getValues().filter(r => r[3].toString() === idEvento.toString());
  const totalAlumnos = alumnos.length;
  if (totalAlumnos === 0) return { error: true, message: "Sin alumnos inscritos." };

  // CORRECCIÓN GRÁFICO NOTAS (UNDEFINED)
  const rawAct = shAct.getDataRange().getValues();
  // Filtramos y mapeamos correctamente usando indices de array
  const actividades = rawAct.filter(r => r[1].toString() === idEvento.toString()).map(r => ({
      id: r[0], 
      nombre: r[2] // Columna C es el nombre
  }));

  const rawCalif = shCalif ? shCalif.getDataRange().getValues() : [];
  
  const evolucionNotas = actividades.map(act => {
    const notas = rawCalif.filter(n => n[1].toString() === act.id.toString());
    let suma = 0; let conteo = 0;
    notas.forEach(n => {
      const val = parseFloat(n[3]);
      if(!isNaN(val)) { suma += val; conteo++; }
    });
    return { label: act.nombre, valor: conteo > 0 ? (suma / conteo).toFixed(1) : 0 };
  });

  // GRÁFICO ASISTENCIA
  const rawAsist = shAsist.getDataRange().getValues(); rawAsist.shift(); 
  const asistEvento = rawAsist.filter(r => r[1].toString() === idEvento.toString());
  const fechasMap = {};
  
  asistEvento.forEach(r => {
    let fStr = (r[3] instanceof Date) ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "dd/MM") : r[3].toString().substring(0, 5);
    if(!fechasMap[fStr]) fechasMap[fStr] = 0;
    if(r[4] === "Presente" || r[4] === "Justificado") fechasMap[fStr]++;
  });

  const tendenciaAsist = Object.keys(fechasMap).map(f => ({
    label: f, valor: Math.round((fechasMap[f] / totalAlumnos) * 100)
  }));

  return { success: true, notas: evolucionNotas, asistencia: tendenciaAsist };
}
// --- DASHBOARD DE INTELIGENCIA DE DATOS PRO ---
function getReporteInteligencia(filtroDocente, fechaInicio, fechaFin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shE = ss.getSheetByName("Eventos");
    const shP = ss.getSheetByName("Participantes");
    const shA = ss.getSheetByName("Asistencia");
    const shAct = ss.getSheetByName("Actividades");
    const shCalif = ss.getSheetByName("Calificaciones");

    if (!shE || !shP || !shA) return { error: true, message: "Faltan hojas base." };

    const rawE = shE.getDataRange().getValues(); rawE.shift();
    const rawP = shP.getDataRange().getValues(); rawP.shift();
    const rawA = shA.getDataRange().getValues(); rawA.shift();
    const rawAct = shAct ? shAct.getDataRange().getValues() : []; rawAct.shift();
    const rawCalif = shCalif ? shCalif.getDataRange().getValues() : []; rawCalif.shift();

    // 1. Filtrar Eventos
    let eventos = rawE.filter(e => {
      // e[0]=ID, e[1]=Nombre, e[2]=Inicio, e[3]=Tipo, e[4]=Responsable, e[5]=Cupo, e[6]=Fin
      if (!e[2]) return false;
      const fecha = new Date(e[2]);
      const doc = e[4] ? e[4].toString().toLowerCase() : "";
      
      let okFecha = true;
      if (fechaInicio && fechaFin) {
         okFecha = (fecha >= new Date(fechaInicio) && fecha <= new Date(fechaFin));
      }
      let okDoc = true;
      if (filtroDocente && filtroDocente !== "Todos los Docentes" && filtroDocente !== "") {
         okDoc = (doc === filtroDocente.toLowerCase());
      }
      return okFecha && okDoc;
    });

    if (eventos.length === 0) return { error: true, message: "No hay datos para este periodo." };

    // VARIABLES GLOBALES
    let totalInscritos = 0;
    let totalCupo = 0;
    let sumaPromedios = 0;
    let conteoNotas = 0;
    let aprobados = 0;
    let reprobados = 0;
    let docentesStats = {}; // { "juan@test.com": {asistencias: 50, clases: 5} }

    // DATA PARA GRÁFICOS
    let labelsCursos = [];
    let dataPromedios = [];
    let excelData = [["ID", "Curso", "Docente", "Inscritos", "Asistencia %", "Promedio", "Estado"]];

    eventos.forEach(e => {
      const idEv = e[0].toString();
      const nombre = e[1];
      const docente = e[4] ? e[4].toString() : "";
      const cupo = parseInt(e[5]) || 0;

      // Alumnos del curso
      const alumnosCurso = rawP.filter(p => p[3].toString() === idEv);
      const numInscritos = alumnosCurso.length;
      totalInscritos += numInscritos;
      totalCupo += cupo;

      // Asistencia del curso
      const asistenciasCurso = rawA.filter(a => a[1].toString() === idEv && (a[4] === "Presente" || a[4] === "Justificado")).length;
      const tasaAsist = numInscritos > 0 ? Math.round((asistenciasCurso / numInscritos) * 100) : 0; // Aproximado

      // Notas del curso
      let sumaNotasCurso = 0;
      let alumnosConNota = 0;
      
      // Buscar actividades del curso
      const actividadesCurso = rawAct.filter(act => act[1].toString() === idEv);
      
      alumnosCurso.forEach(alum => {
         let notaFinal = 0;
         actividadesCurso.forEach(act => {
            const nota = rawCalif.find(c => c[1].toString() === act[0].toString() && c[2].toString() === alum[0].toString());
            if(nota) notaFinal += parseFloat(nota[3]) * (act[4]/100);
         });
         
         // Si hubo actividades, contamos la nota
         if(actividadesCurso.length > 0) {
            sumaNotasCurso += notaFinal;
            alumnosConNota++;
            if(notaFinal >= 7.0) aprobados++; else reprobados++;
         }
      });

      const promedioCurso = alumnosConNota > 0 ? (sumaNotasCurso / alumnosConNota) : 0;
      if(alumnosConNota > 0) {
         sumaPromedios += promedioCurso;
         conteoNotas++;
      }

      // Stats Docente
      if(!docentesStats[docente]) docentesStats[docente] = { cursos: 0, sumaAsist: 0 };
      docentesStats[docente].cursos++;
      docentesStats[docente].sumaAsist += tasaAsist;

      // Data Gráficos
      labelsCursos.push(nombre);
      dataPromedios.push(promedioCurso.toFixed(1));

      // Excel
      excelData.push([idEv, nombre, docente, numInscritos, tasaAsist + "%", promedioCurso.toFixed(1), promedioCurso >= 7 ? "OK" : "Bajo"]);
    });

    // 3. CALCULAR KPI FINALES
    const promedioGlobal = conteoNotas > 0 ? (sumaPromedios / conteoNotas).toFixed(1) : "0.0";
    const tasaOcupacion = totalCupo > 0 ? Math.round((totalInscritos / totalCupo) * 100) : 0;
    
    // Mejor Docente
    let mejorDocente = "N/A";
    let mejorScore = -1;
    for (const [email, stats] of Object.entries(docentesStats)) {
       const score = stats.sumaAsist / stats.cursos;
       if(score > mejorScore) { mejorScore = score; mejorDocente = email; }
    }

    return {
      success: true,
      kpis: {
        totalCursos: eventos.length,
        totalAlumnos: totalInscritos,
        promedioGlobal: promedioGlobal,
        tasaAprobacion: (aprobados + reprobados) > 0 ? Math.round((aprobados / (aprobados + reprobados)) * 100) : 0,
        mejorDocente: mejorDocente
      },
      charts: {
        cursos: labelsCursos,
        promedios: dataPromedios,
        distribucion: [aprobados, reprobados] // Pie chart
      },
      excelData: excelData
    };

  } catch (e) {
    return { error: true, message: "Error backend: " + e.toString() };
  }
}

// --- HELPER: LISTA DE DOCENTES (PARA FILTRO) ---
function getListaDocentes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Eventos");
    if(!sheet) return [];
    
    // Obtener columna de responsables (Col F / índice 4 en array 0-indexed después de quitar header?? No)
    // sheet.getDataRange().getValues() -> Col F es index 4?
    // A=0, B=1, C=2, D=3, E=4 (Tipo), F=5 (??? No, revisemos editarEventoServidor)
    // En editarEventoServidor: datos.responsable es rowData[4]
    // rowData[0]=ID, [1]=Nombre, [2]=Fecha, [3]=Tipo, [4]=Responsable, [5]=Cupo.
    // Correcto. Index 4 es Responsable.
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Quitar headers
    
    const docentes = new Set();
    data.forEach(r => {
      if(r[4]) docentes.add(r[4].toString().trim());
    });
    
    return Array.from(docentes).sort();
  } catch(e) {
    return [];
  }
}

// --- VISTA PREVIA DEL DIPLOMA (SIN ENVIAR) ---
function previewDiplomaBackend(idEvento, idParticipante) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener datos del alumno y curso
  const libro = getLibroCalificacionesV2(idEvento);
  const alumno = libro.alumnos.find(a => a.id.toString() === idParticipante.toString());
  
  if(!alumno) return { success: false, message: "Alumno no encontrado" };
  if(alumno.estado !== "APROBADO") return { success: false, message: "El alumno aún no aprueba el curso." };

  const ev = ss.getSheetByName("Eventos").getDataRange().getValues().find(e => e[0].toString() === idEvento.toString());
  const nombreCurso = ev ? ev[1] : "Curso";

  // 🔽 ID DE TU PLANTILLA DE SLIDES (La misma que usas para enviar) 🔽
  const ID_PLANTILLA = "1UwxbrCQJe992sMOI3CFCJJGzrTU2FA-LcH9BIfQH7-M"; 
  
  try {
    // 2. Generar PDF Temporal
    const archivoPlantilla = DriveApp.getFileById(ID_PLANTILLA);
    const copiaArchivo = archivoPlantilla.makeCopy("Temp_Preview_" + alumno.nombre);
    const idCopia = copiaArchivo.getId();
    const presentacion = SlidesApp.openById(idCopia);
    
    // Reemplazar variables
    presentacion.replaceAllText("{{Nombre}}", alumno.nombre);
    presentacion.replaceAllText("{{Curso}}", nombreCurso);
    presentacion.replaceAllText("{{Nota}}", alumno.promedio.toString());
    presentacion.replaceAllText("{{Fecha}}", new Date().toLocaleDateString());
    
    presentacion.saveAndClose();
    
    // 3. Convertir a Base64 para mostrar en el navegador
    const pdfBlob = copiaArchivo.getAs(MimeType.PDF);
    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    
    // 4. Borrar archivo temporal inmediatamente
    copiaArchivo.setTrashed(true);
    
    return { success: true, base64: base64Data };

  } catch (e) {
    return { success: false, message: "Error generando vista previa: " + e.message };
  }
}
// --- NUEVA SEGURIDAD: LOGIN ALUMNO CON CÓDIGO (2FA) ---

// 1. PASO 1: Generar código y enviar correo
function enviarCodigoAccesoAlumno(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  if (!shP) return { success: false, message: "Error de sistema." };

  const data = shP.getDataRange().getValues();
  const emailLimpio = email.toString().trim().toLowerCase();
  
  // Buscar si el correo existe (buscamos todas las filas de ese alumno)
  // Usamos un array para guardar las filas donde aparece, aunque el token lo guardaremos en la primera coincidencia o en todas.
  // Para simplificar, buscaremos la PRIMERA vez que aparece el alumno para guardar su token maestro.
  let filaAlumno = -1;
  let nombreAlumno = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().toLowerCase().trim() === emailLimpio) {
      filaAlumno = i + 1;
      nombreAlumno = data[i][1];
      break; 
    }
  }

  if (filaAlumno === -1) {
    return { success: false, message: "Este correo no está inscrito en nuestros cursos." };
  }

  // Generar Token (6 dígitos)
  const token = Math.floor(100000 + Math.random() * 900000).toString();
  // Expiración (10 minutos)
  const expiracion = new Date(new Date().getTime() + 10 * 60000); 

  // Guardar en la hoja Participantes.
  // Asumimos que las columnas F (6) y G (7) están libres o se usan para esto.
  // Estructura: A(0), B(1), C(2), D(3), E(4=Diploma), F(5=Token), G(6=Expiracion)
  shP.getRange(filaAlumno, 6).setValue(token);
  shP.getRange(filaAlumno, 7).setValue(expiracion);

  // Enviar Correo
  try {
    MailApp.sendEmail({
      to: emailLimpio,
      subject: "🔐 Tu Código de Acceso - Portal Alumno",
      htmlBody: `
        <div style="font-family:sans-serif; padding:20px; border:1px solid #e2e8f0; border-radius:10px; text-align:center; max-width:500px; margin:0 auto;">
           <h2 style="color:#2563eb;">Código de Verificación</h2>
           <p>Hola <strong>${nombreAlumno}</strong>,</p>
           <p>Estás intentando ingresar al Portal del Estudiante. Usa este código:</p>
           <div style="background:#f1f5f9; padding:15px; font-size:2.5rem; letter-spacing:5px; font-weight:bold; color:#1e293b; margin:20px 0; border-radius:8px;">
              ${token}
           </div>
           <p style="color:#64748b; font-size:0.9rem;">Este código expira en 10 minutos.</p>
           <hr style="border:0; border-top:1px solid #e2e8f0; margin:20px 0;">
           <p style="font-size:0.8rem; color:#94a3b8;">Si no solicitaste este acceso, ignora este correo.</p>
        </div>
      `
    });
  } catch(e) {
    return { success: false, message: "Error enviando correo: " + e.message };
  }

  return { success: true, message: "Código enviado a " + emailLimpio };
}

// 2. PASO 2: Validar código y devolver datos (Reemplaza la lógica anterior)
function validarCodigoYEntrar(email, codigoInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const lastRow = shP.getLastRow();
  if (lastRow < 2) return { success: false, message: "No hay participantes registrados." };
  
  const data = shP.getRange(1, 1, lastRow, 7).getValues();
  const codigoLimpio = codigoInput.toString().trim();

  if (!codigoLimpio) return { success: false, message: "Ingresa un código válido." };

  let filaAlumno = -1;
  let idParticipante = "";
  let nombreAlumno = "";
  let emailAlumno = "";

  // BUSCAR POR TOKEN EN TODAS LAS FILAS (no depender del email del frontend)
  for (let i = 1; i < data.length; i++) {
    const tokenGuardado = data[i][5]; // Columna F
    
    if (tokenGuardado && tokenGuardado.toString().trim() === codigoLimpio) {
      // Token encontrado, verificar expiración
      const expiracion = new Date(data[i][6]); // Columna G
      const ahora = new Date();
      
      if (ahora > expiracion) {
        return { success: false, message: "El código ha expirado. Solicita uno nuevo." };
      }
      
      // ¡CÓDIGO CORRECTO! Limpiar token
      shP.getRange(i+1, 6).setValue("");
      shP.getRange(i+1, 7).setValue("");
      
      filaAlumno = i;
      idParticipante = data[i][0];
      nombreAlumno = data[i][1];
      emailAlumno = data[i][2] ? data[i][2].toString().toLowerCase().trim() : "";
      break;
    }
  }

  if (filaAlumno === -1) {
    registrarLog(email || "Alumno", "Intento Fallido", "Seguridad", "Código 2FA incorrecto o expirado");
    return { success: false, message: "Código incorrecto o expirado. Verifica e intenta de nuevo." };
  }

  // --- AUDITORÍA: ÉXITO ---
  registrarLog(emailAlumno, "Acceso", "Seguridad", "Alumno inició sesión con código 2FA: " + nombreAlumno);

  // Si llegamos aquí, es válido. Recuperamos sus datos académicos.
  const datosAcademicos = obtenerKardexAlumno(idParticipante);
  
  if (!datosAcademicos.success) {
      return { success: false, message: "Error recuperando historial." };
  }

  return { 
    success: true, 
    nombre: nombreAlumno, 
    datos: datosAcademicos 
  };
}
// --- DESCARGA DE DIPLOMA (PORTAL ALUMNO) ---
function descargarDiplomaAlumno(idEvento, emailAlumno) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Validar que el alumno realmente aprobó ese curso
  // Reutilizamos la lógica del libro de calificaciones para seguridad
  const libro = getLibroCalificacionesV2(idEvento);
  const alumno = libro.alumnos.find(a => a.email.toLowerCase().trim() === emailAlumno.toLowerCase().trim());

  if (!alumno) return { success: false, message: "No se encontraron tus datos en este curso." };
  if (alumno.estado !== "APROBADO") return { success: false, message: "Aún no tienes el estatus de APROBADO." };

  // 2. Datos del Evento
  const ev = ss.getSheetByName("Eventos").getDataRange().getValues().find(e => e[0].toString() === idEvento.toString());
  const nombreCurso = ev ? ev[1] : "Curso";

  // 3. ID DE LA PLANTILLA (Asegúrate que sea el correcto)
  const ID_PLANTILLA = "1UwxbrCQJe992sMOI3CFCJJGzrTU2FA-LcH9BIfQH7-M"; 

  try {
    // Generar PDF Temporal
    const archivoPlantilla = DriveApp.getFileById(ID_PLANTILLA);
    const copiaArchivo = archivoPlantilla.makeCopy("Diploma_" + alumno.nombre);
    const idCopia = copiaArchivo.getId();
    const presentacion = SlidesApp.openById(idCopia);
    
    // Reemplazar textos
    presentacion.replaceAllText("{{Nombre}}", alumno.nombre);
    presentacion.replaceAllText("{{Curso}}", nombreCurso);
    presentacion.replaceAllText("{{Nota}}", alumno.promedio.toString());
    presentacion.replaceAllText("{{Fecha}}", new Date().toLocaleDateString());
    
    presentacion.saveAndClose();
    
    // Obtener Blob y Base64
    const pdfBlob = copiaArchivo.getAs(MimeType.PDF);
    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    
    // Borrar temporal
    copiaArchivo.setTrashed(true);
    
    return { success: true, base64: base64Data, fileName: "Diploma - " + nombreCurso + ".pdf" };

  } catch (e) {
    return { success: false, message: "Error al generar documento: " + e.message };
  }
}
// --- CENTRO DE REPORTES CENTRALIZADO ---
function getReporteCentralizado(tipoReporte) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shE = ss.getSheetByName("Eventos");
  const shP = ss.getSheetByName("Participantes");
  const shA = ss.getSheetByName("Asistencia");
  const shC = ss.getSheetByName("Calificaciones");
  
  if(!shE || !shP) return { error: true, message: "Faltan hojas de datos." };

  const dataE = shE.getDataRange().getValues(); dataE.shift();
  const dataP = shP.getDataRange().getValues(); dataP.shift();
  const dataA = shA ? shA.getDataRange().getValues() : []; dataA.shift();
  const dataC = shC ? shC.getDataRange().getValues() : []; dataC.shift();

  let filas = [];
  let columnas = [];
  let titulo = "";

  try {
    switch (tipoReporte) {
      
      // 1. REPORTE DE DESEMPEÑO DOCENTE (CORREGIDO)
      case 'DOCENTES':
        titulo = "Desempeño y Carga Docente";
        columnas = ["Docente", "Cursos", "Alumnos", "Asist. %", "Nota Promedio"];
        
        const docentesMap = {};
        
        // Pre-procesamos calificaciones para búsqueda rápida: { "ID_PARTICIPANTE": Nota }
        const notasMap = {};
        dataC.forEach(row => {
           // Asumimos col 0 = ID_Part, col 2 = Nota Final
           notasMap[row[0].toString()] = parseFloat(row[2]) || 0; 
        });

        dataE.forEach(e => {
           const idEv = e[0].toString();
           const doc = e[4].toString().toLowerCase(); // Columna Docente
           if(!doc) return;
           
           if(!docentesMap[doc]) docentesMap[doc] = { cursos: 0, alumnos: 0, sumaAsist: 0, sumaNotas: 0, nNotas: 0 };
           
           // Alumnos inscritos en este curso
           const inscritos = dataP.filter(p => p[3].toString() === idEv);
           const nInscritos = inscritos.length;
           
           // Cálculo de Asistencia
           const asistencias = dataA.filter(a => a[1].toString() === idEv && (a[4]==="Presente"||a[4]==="Justificado")).length;
           const tasaAsist = nInscritos > 0 ? (asistencias / nInscritos) : 0; 
           
           // Cálculo de Notas (Buscamos la nota de cada inscrito)
           let sumaNotasCurso = 0;
           let conNota = 0;
           inscritos.forEach(p => {
              const idPart = p[0].toString();
              if(notasMap[idPart]) {
                 sumaNotasCurso += notasMap[idPart];
                 conNota++;
              }
           });

           docentesMap[doc].cursos++;
           docentesMap[doc].alumnos += nInscritos;
           docentesMap[doc].sumaAsist += tasaAsist;
           docentesMap[doc].sumaNotas += sumaNotasCurso;
           docentesMap[doc].nNotas += conNota;
        });

        // Convertir a Array Final
        for (const [doc, stats] of Object.entries(docentesMap)) {
           // Promedios finales
           let promAsist = stats.alumnos > 0 ? Math.round((stats.sumaAsist / stats.cursos) * 10) : 0; 
           if(promAsist > 100) promAsist = 100;

           let promNota = stats.nNotas > 0 ? (stats.sumaNotas / stats.nNotas).toFixed(1) : "0.0";

           filas.push([ doc, stats.cursos, stats.alumnos, promAsist + "%", promNota ]);
        }
        break;

      // 2. REPORTE DE OCUPACIÓN
      case 'OCUPACION':
        titulo = "Ocupación y Demanda de Talleres";
        columnas = ["ID", "Taller / Curso", "Cupo Máx", "Inscritos", "% Ocupación", "Estado"];
        
        dataE.forEach(e => {
           const id = e[0];
           const nombre = e[1];
           const cupo = parseInt(e[5]) || 0;
           const inscritos = dataP.filter(p => p[3].toString() === id.toString()).length;
           const porc = cupo > 0 ? Math.round((inscritos/cupo)*100) : 0;
           
           let estado = "🟢 Disponible";
           if(porc >= 80) estado = "🟡 Llenándose";
           if(porc >= 100) estado = "🔴 Lleno";

           filas.push([id, nombre, cupo, inscritos, porc + "%", estado]);
        });
        break;

      // 3. AUDITORÍA DE DIPLOMAS
      case 'DIPLOMAS':
        titulo = "Auditoría de Certificación";
        columnas = ["Alumno", "Correo", "Curso Aprobado", "Fecha", "Estado Diploma", "Acción"];
        
        // Buscamos alumnos con diploma pendiente (suponemos aprobado si está en esta lista filtrada logicamente, 
        // para este ejemplo rápido listamos todos y marcamos pendientes)
        dataP.forEach(p => {
           const fechaDiploma = p[4]; // Columna E
           if(!fechaDiploma) {
              const evento = dataE.find(e => e[0].toString() === p[3].toString());
              const nomEvento = evento ? evento[1] : "Desconocido";
              filas.push([ p[1], p[2], nomEvento, "-", "⚠️ PENDIENTE", "Ver Gradebook" ]);
           }
        });
        // Limitamos a 50 para no saturar si hay muchos
        filas = filas.slice(0, 50);
        break;

      // 4. ALUMNOS EN RIESGO (Lógica Simplificada para velocidad)
      case 'RIESGO':
        titulo = "Alerta Temprana - Alumnos en Riesgo";
        columnas = ["Alumno", "Curso", "Asistencias", "Faltas", "Riesgo"];
        
        // Agrupar asistencia por alumno/curso
        const mapRiesgo = {};
        dataA.forEach(a => {
           const key = a[2] + "|" + a[1]; // ID_Part | ID_Evento
           if(!mapRiesgo[key]) mapRiesgo[key] = { presentes: 0, faltas: 0 };
           
           if(a[4] === "Presente" || a[4] === "Justificado") mapRiesgo[key].presentes++;
           else mapRiesgo[key].faltas++;
        });

        for (const [key, stats] of Object.entries(mapRiesgo)) {
           const total = stats.presentes + stats.faltas;
           const rate = total > 0 ? (stats.presentes / total) : 0;
           
           if(rate < 0.70) { // Menos del 70% asistencia
              const [idPart, idEv] = key.split("|");
              const part = dataP.find(p => p[0].toString() === idPart);
              const ev = dataE.find(e => e[0].toString() === idEv);
              
              if(part && ev) {
                 filas.push([ part[1], ev[1], stats.presentes, stats.faltas, "🚨 ASISTENCIA BAJA ("+Math.round(rate*100)+"%)" ]);
              }
           }
        }
        break;
        
      case 'TOP':
        titulo = "Cuadro de Honor (Mejores Promedios)";
        columnas = ["#", "Alumno", "Curso", "Promedio Final", "Insignia"];
        
        // Esto requeriría procesar el gradebook completo. Para este ejemplo, simularemos con datos de muestra
        // Idealmente aquí llamas a getLibroCalificacionesV2 para cada curso activo.
        filas.push(["1", "Ana Gomez", "Python Básico", "10.0", "🥇 Oro"]);
        filas.push(["2", "Carlos Ruiz", "Excel Avanzado", "9.8", "🥈 Plata"]);
        filas.push(["3", "Maria Lopez", "Liderazgo", "9.5", "🥉 Bronce"]);
        break;
    }

    return { success: true, titulo: titulo, columnas: columnas, filas: filas };

  } catch (e) {
    return { success: false, message: "Error generando reporte: " + e.message };
  }
}
/**
 * Registra la asistencia mediante QR y genera el log de auditoría.
 * @param {string} codigoParticipante ID del alumno (ej: P-001 o P-001|EV-01)
 * @param {string} idEventoActual ID del evento seleccionado en la UI
 */
function registrarAsistenciaQR(codigoParticipante, idEventoActual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  const shA = ss.getSheetByName("Asistencia");
  
  try {
    // 1. LIMPIEZA DE DATOS
    // Si el QR viene con formato "ID_ALUMNO|ID_EVENTO", separamos el ID del alumno
    let idParticipante = codigoParticipante;
    if (codigoParticipante.toString().includes("|")) {
        idParticipante = codigoParticipante.toString().split("|")[0].trim();
    }

    // 2. VALIDACIÓN DE EXISTENCIA DEL ALUMNO
    const participantes = shP.getDataRange().getValues();
    let nombreAlumno = "";
    let existe = false;

    // Buscamos en la hoja Participantes (ID en Columna A, Nombre en Columna B)
    for (let i = 1; i < participantes.length; i++) {
      if (String(participantes[i][0]) === String(idParticipante)) {
        nombreAlumno = participantes[i][1]; 
        existe = true;
        break;
      }
    }

    if (!existe) {
      registrarLog("QR Scanner", "Error QR", "Asistencia", `ID no registrado intentó marcar: ${idParticipante}`);
      return { success: false, message: "❌ El código escaneado no existe en el sistema." };
    }

    // 3. VALIDACIÓN DE DUPLICADOS (Evitar doble marca el mismo día para el mismo evento)
    const hoyStr = Utilities.formatDate(new Date(), "GMT-6", "yyyy-MM-dd");
    const asistencias = shA.getDataRange().getValues();
    
    const yaRegistro = asistencias.some(r => {
      let fRegistro = "";
      if (r[3] instanceof Date) {
        fRegistro = Utilities.formatDate(r[3], "GMT-6", "yyyy-MM-dd");
      }
      // Columna B: ID Evento, Columna C: ID Participante, Columna D: Fecha
      return String(r[1]) === String(idEventoActual) && 
             String(r[2]) === String(idParticipante) && 
             fRegistro === hoyStr;
    });

    if (yaRegistro) {
       return { success: false, message: `⚠️ ${nombreAlumno} ya registró asistencia para este evento hoy.` };
    }

    // 4. REGISTRO EN TABLA ASISTENCIA
    // Formato: [ID_Log, ID_Evento, ID_Participante, Fecha_Hora, Estado, Metodo]
    const idUnicoAsis = "ASIS-" + Utilities.getUuid().substring(0, 5).toUpperCase();
    shA.appendRow([
      idUnicoAsis, 
      idEventoActual, 
      idParticipante, 
      new Date(), 
      "Presente", 
      "QR"
    ]);

    // 5. REGISTRO EN TABLA AUDITORIA (Lo que soluciona tu problema de logs)
    // El orden en Auditoria es: [ID_Log, Fecha, Usuario, Acción, Detalle]
    registrarLog("QR Scanner", "Asistencia", "Participantes", `Registro exitoso: ${nombreAlumno} en evento ${idEventoActual}`);

    return { 
      success: true, 
      mensaje: `${nombreAlumno} - Presente` 
    };

  } catch (error) {
    // Si algo falla, lo auditamos también (Punto extra en seguridad)
    registrarLog("QR Scanner", "Error", "Participantes", `Fallo en registro QR: ${error.message}`);
    return { success: false, message: "Error técnico: " + error.toString() };
  }
}
function PROBAR_SEPARACION_DATOS() {
  // 1. Simulamos exactamente el dato que te da problemas
  var codigoEscaneado = "P-61610|EV-TEST-8068";
  
  Logger.log("--- INICIANDO PRUEBA DE DIAGNÓSTICO ---");
  Logger.log("Dato original recibido del escáner: " + codigoEscaneado);
  
  // 2. Aplicamos la lógica de separación
  var idLimpio = codigoEscaneado;
  
  if (codigoEscaneado.toString().includes("|")) {
      var partes = codigoEscaneado.toString().split("|");
      idLimpio = partes[0].trim(); // Tomamos la parte izquierda
      Logger.log("✅ Separador '|' detectado.");
      Logger.log("Parte 1 (ID Alumno): " + partes[0]);
      Logger.log("Parte 2 (Evento): " + partes[1]);
  } else {
      Logger.log("⚠️ No se detectó el separador '|'.");
  }
  
  Logger.log("-------------------------------------");
  Logger.log("RESULTADO FINAL: El sistema buscará en la base de datos el ID: '" + idLimpio + "'");
  
  if (idLimpio === "P-61610") {
    Logger.log("🎉 PRUEBA EXITOSA: El código funciona correctamente.");
  } else {
    Logger.log("❌ PRUEBA FALLIDA: Sigue mostrando el texto largo.");
  }
}
// --- FUNCIÓN DE APOYO PARA AUDITORÍA ---


// --- PORTAL DEL ALUMNO (BACKEND) ---
// Movido desde js.txt para ejecución en servidor
function loginAlumno(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName("Participantes");
  
  if (!shP) return { success: false, message: "Error crítico: No se encuentra la hoja de Participantes." };

  const data = shP.getDataRange().getValues();
  // Validar que el email venga limpio
  if(!email) return { success: false, message: "Correo vacío." };
  
  const emailLimpio = email.toString().trim().toLowerCase();
  
  // Buscar si el correo existe en la base de datos (Columna C -> índice 2)
  // Buscamos la primera coincidencia para identificar al alumno
  const alumnoEncontrado = data.find(r => r[2] && r[2].toString().toLowerCase().trim() === emailLimpio);
  
  if (!alumnoEncontrado) {
    return { success: false, message: "Este correo no está inscrito en ningún curso." };
  }
  
  // Si existe, usamos su ID (Columna A -> índice 0) para generar el Kardex completo
  const idParticipante = alumnoEncontrado[0];
  
  // Llamamos a la función interna para obtener sus notas
  const datosAcademicos = obtenerKardexAlumno(idParticipante);
  
  if (!datosAcademicos.success) {
      return { success: false, message: "Error recuperando el historial académico." };
  }

  return { 
    success: true, 
    nombre: alumnoEncontrado[1], // Nombre del alumno
    datos: datosAcademicos 
  };
}

// ==========================================
// ===  AUDITORÍA (SOLICITADA EN FASE 15) ===
// ==========================================

function registrarLog(usuario, modulo, accion, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Auditoria");
  if (!sheet) {
    sheet = ss.insertSheet("Auditoria");
    sheet.appendRow(["Marca Temporal", "Usuario", "Módulo", "Acción", "Detalle"]);
  }
  sheet.appendRow([new Date(), usuario, modulo, accion, detalle]);
}

function getLogsServidor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Auditoria");
    
    // Si no existe la hoja, devolver array vacío
    if (!sheet) {
      Logger.log("Hoja 'Log' no encontrada, devolviendo array vacío");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Si solo hay header o está vacía
    if (data.length <= 1) {
      Logger.log("Hoja 'Log' vacía o solo con encabezados");
      return [];
    }
    
    data.shift(); // Quitar header
    
    // Mapear y ordenar por fecha descendente (más reciente primero)
    const logs = data.map(r => ({
      fecha: r[0] instanceof Date ? r[0].toLocaleString('es-ES') : String(r[0]),
      usuario: String(r[1] || ''),
      modulo: String(r[2] || ''),
      accion: String(r[3] || ''),
      detalle: String(r[4] || '')
    }));
    
    Logger.log(`getLogsServidor: Devolviendo ${logs.length} registros`);
    return logs.reverse(); // Asumiendo que se agregan al final, reverse los pone reciente primero
  } catch (e) {
    Logger.log("ERROR en getLogsServidor: " + e.toString());
    return [];
  }
}
// ==========================================
// ===  GESTIÓN DE USUARIOS (FASE 19)   ===
// ==========================================

function getUsuarios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Usuarios");
    
    if (!sheet) {
      Logger.log("Hoja 'Usuarios' no encontrada");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log("Hoja 'Usuarios' vacía o solo con encabezados");
      return [];
    }
    
    data.shift(); // Quitar header
    
    // Mapear usuarios (excluir password)
    const usuarios = data.map(r => ({
      nombre: String(r[0] || ''),
      email: String(r[1] || ''),
      // password en r[2] - NO se envía al frontend
      rol: String(r[3] || '')
    }));
    
    Logger.log(`getUsuarios: Devolviendo ${usuarios.length} usuarios`);
    return usuarios;
  } catch (e) {
    Logger.log("ERROR en getUsuarios: " + e.toString());
    return [];
  }
}

function crearUsuario(nombre, email, password, rol) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Usuarios");
    
    if (!sheet) {
      return { success: false, message: "Error: Hoja 'Usuarios' no encontrada." };
    }
    
    // Validaciones
    if (!nombre || !email || !password || !rol) {
      return { success: false, message: "Todos los campos son obligatorios." };
    }
    
    if (!['root', 'admin', 'docente'].includes(rol)) {
      return { success: false, message: "Rol inválido. Use: root, admin o docente." };
    }
    
    // Verificar si el email ya existe
    const data = sheet.getDataRange().getValues();
    const emailExiste = data.some((row, idx) => 
      idx > 0 && row[1].toString().toLowerCase().trim() === email.toLowerCase().trim()
    );
    
    if (emailExiste) {
      return { success: false, message: "Ya existe un usuario con ese correo electrónico." };
    }
    
    // Hash de password
    const passHash = hashPassword(password);
    
    // Agregar usuario
    sheet.appendRow([nombre, email, passHash, rol]);
    
    // Registrar en auditoría
    registrarLog(Session.getActiveUser().getEmail(), "Usuarios", "Crear", `Nuevo usuario: ${email} (${rol})`);
    
    return { success: true, message: "Usuario creado exitosamente." };
  } catch (e) {
    Logger.log("ERROR en crearUsuario: " + e.toString());
    return { success: false, message: "Error al crear usuario: " + e.toString() };
  }
}

function editarUsuario(emailOriginal, nuevoNombre, nuevoRol) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Usuarios");
    
    if (!sheet) {
      return { success: false, message: "Error: Hoja 'Usuarios' no encontrada." };
    }
    
    // Validaciones
    if (!emailOriginal || !nuevoNombre || !nuevoRol) {
      return { success: false, message: "Todos los campos son obligatorios." };
    }
    
    if (!['root', 'admin', 'docente'].includes(nuevoRol)) {
      return { success: false, message: "Rol inválido. Use: root, admin o docente." };
    }
    
    const data = sheet.getDataRange().getValues();
    const emailBuscado = emailOriginal.toLowerCase().trim();
    
    // Buscar el usuario
    let filaEncontrada = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase().trim() === emailBuscado) {
        filaEncontrada = i + 1; // +1 porque las filas en Sheets son 1-indexed
        break;
      }
    }
    
    if (filaEncontrada === -1) {
      return { success: false, message: "Usuario no encontrado." };
    }
    
    // Actualizar nombre y rol
    sheet.getRange(filaEncontrada, 1).setValue(nuevoNombre);
    sheet.getRange(filaEncontrada, 4).setValue(nuevoRol);
    
    // Registrar en auditoría
    registrarLog(Session.getActiveUser().getEmail(), "Usuarios", "Editar", `Usuario modificado: ${emailOriginal} → ${nuevoNombre} (${nuevoRol})`);
    
    return { success: true, message: "Usuario actualizado exitosamente." };
  } catch (e) {
    Logger.log("ERROR en editarUsuario: " + e.toString());
    return { success: false, message: "Error al editar usuario: " + e.toString() };
  }
}

function eliminarUsuario(email, usuarioActual) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Usuarios");
    
    if (!sheet) {
      return { success: false, message: "Error: Hoja 'Usuarios' no encontrada." };
    }
    
    const emailBuscado = email.toLowerCase().trim();
    const emailActual = usuarioActual.toLowerCase().trim();
    
    // Validación: No puede eliminarse a sí mismo
    if (emailBuscado === emailActual) {
      return { success: false, message: "No puedes eliminar tu propia cuenta." };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Buscar el usuario
    let filaEncontrada = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase().trim() === emailBuscado) {
        filaEncontrada = i + 1;
        break;
      }
    }
    
    if (filaEncontrada === -1) {
      return { success: false, message: "Usuario no encontrado." };
    }
    
    // Eliminar la fila
    sheet.deleteRow(filaEncontrada);
    
    // Registrar en auditoría
    registrarLog(usuarioActual, "Usuarios", "Eliminar", `Usuario eliminado: ${email}`);
    
    return { success: true, message: "Usuario eliminado exitosamente." };
  } catch (e) {
    Logger.log("ERROR en eliminarUsuario: " + e.toString());
    return { success: false, message: "Error al eliminar usuario: " + e.toString() };
  }
}
// ==========================================
// ===  BACKUP BD EN EXCEL (FASE 20)    ===
// ==========================================

function generarBackupExcel() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
    const nombreBackup = `Backup_Sistema_Eventos_${fecha}`;
    
    // Crear una copia del spreadsheet
    const archivo = DriveApp.getFileById(ssId);
    const copia = archivo.makeCopy(nombreBackup);
    
    // Convertir a Excel usando la URL de exportación
    const urlExcel = `https://docs.google.com/spreadsheets/d/${copia.getId()}/export?format=xlsx`;
    
    // Registrar en auditoría
    registrarLog(Session.getActiveUser().getEmail(), "Sistema", "Backup", `Generó backup: ${nombreBackup}`);
    
    return { 
      success: true, 
      url: urlExcel,
      urlSheet: copia.getUrl(),
      nombre: nombreBackup + '.xlsx',
      id: copia.getId()
    };
  } catch (e) {
    Logger.log("ERROR en generarBackupExcel: " + e.toString());
    return { success: false, message: "Error al generar backup: " + e.toString() };
  }
}

function getDataJSON() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const exportData = {};

    sheets.forEach(sheet => {
      const name = sheet.getName();
      // Skip config sheets if needed, e.g. "Config"
      const data = sheet.getDataRange().getValues();
      exportData[name] = data;
    });

    return {
      success: true,
      json: JSON.stringify(exportData, null, 2),
      fileName: `Backup_Sistema_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}.json`
    };

  } catch (e) {
    registrarLog("Sistema", "Error Backup", "Sistema", `Fallo al generar backup JSON: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function restaurarDatosJSON(jsonContent) {
  try {
    const data = JSON.parse(jsonContent);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Safety Backup before wipe
    const safetyName = `SAFETY_BACKUP_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmmss")}`;
    DriveApp.getFileById(ss.getId()).makeCopy(safetyName);

    // Iterar sobre las hojas del JSON
    for (const [sheetName, rows] of Object.entries(data)) {
      if (!Array.isArray(rows) || rows.length === 0) continue;
      
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        // Crear hoja si no existe (opcional, por robustez)
        sheet = ss.insertSheet(sheetName);
      }
      
      // Limpiar contenido actual
      sheet.clearContents();
      
      // Escribir nuevos datos
      // rows es array de arrays. Apps Script requiere rectangular grid.
      // Normalizamos longitud de filas por si acaso
      const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 0);
      const rowsNormalized = rows.map(r => {
        while(r.length < maxCols) r.push("");
        return r;
      });
      
      if(rowsNormalized.length > 0) {
        sheet.getRange(1, 1, rowsNormalized.length, maxCols).setValues(rowsNormalized);
      }
    }
    
    registrarLog("Sistema", "Restauración", "Datos", "Se restauró la base de datos desde JSON.");
    return { success: true };
    
  } catch (e) {
    registrarLog("Sistema", "Error Restore", "Sistema", `Fallo crítico al restaurar JSON: ${e.message}`);
    return { success: false, message: "Error al restaurar: " + e.message };
  }
}

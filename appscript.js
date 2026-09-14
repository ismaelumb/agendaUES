// ✅ ID DE TU GOOGLE SHEET - ACTUALIZADO
const ID_HOJA = '1jl9zsfRKMqAo51wDNaf5QB14bRPaEZpBo3t0VfZhO2E';

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const accion = data.accion;
  const sheet = SpreadsheetApp.openById(ID_HOJA);
  
  try {
    if (accion === 'login') return login(data, sheet);
    if (accion === 'crear_cita') return crearCita(data, sheet);
    if (accion === 'editar_cita') return editarCita(data, sheet);
    if (accion === 'obtener_citas') return obtenerCitas(data, sheet);
    if (accion === 'crear_usuario') return crearUsuario(data, sheet);
    
    return respuestaExito({ mensaje: "Acción no válida" }, false);
  } catch (error) {
    return respuestaExito({ error: error.toString() }, false);
  }
}

// ✅ FUNCIÓN LOGIN (LA QUE FALTABA)
function login(data, sheet) {
  const hojaUsuarios = sheet.getSheetByName('Usuarios');
  const datos = hojaUsuarios.getDataRange().getValues();
  
  // Recorrer todos los usuarios (saltando el encabezado)
  for (let i = 1; i < datos.length; i++) {
    const usuario = datos[i][1];       // Columna B (Usuario)
    const contrasena = datos[i][2];    // Columna C (Contraseña)
    const rol = datos[i][3];           // Columna D (Rol)
    const idUsuario = datos[i][0];     // Columna A (ID_Usuario)
    
    // Verificar credenciales
    if (usuario === data.usuario && contrasena === data.contrasena) {
      registrarBitacora(sheet, idUsuario, "LOGIN", `Login exitoso - ${data.usuario}`);
      return respuestaExito({
        mensaje: "Login exitoso",
        idUsuario: idUsuario,
        usuario: usuario,
        rol: rol
      }, true);
    }
  }
  
  registrarBitacora(sheet, "DESCONOCIDO", "LOGIN_FALLIDO", `Intento fallido - Usuario: ${data.usuario}`);
  return respuestaExito({ mensaje: "Usuario o contraseña incorrectos" }, false);
}

// 1. FUNCIÓN PARA CREAR CITA
function crearCita(data, sheet) {
  const hojaCitas = sheet.getSheetByName('Citas');
  const idCita = Utilities.getUuid();
  
  hojaCitas.appendRow([idCita, data.idUsuario, data.titulo, data.fechaHora, data.detalles]);
  registrarBitacora(sheet, data.idUsuario, "CREAR_CITA", `Cita creada: ${data.titulo} para ${data.fechaHora}`);
  
  return respuestaExito({ mensaje: "Cita creada con éxito", idCita: idCita }, true);
}

// 2. FUNCIÓN PARA EDITAR CITA (Con regla de 1 hora y sin borrado)
function editarCita(data, sheet) {
  const hojaCitas = sheet.getSheetByName('Citas');
  const datos = hojaCitas.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === data.idCita && datos[i][1] === data.idUsuario) {
      const fechaCita = new Date(datos[i][3]); // Fecha agendada
      const fechaActual = new Date();
      const diferenciaHoras = (fechaCita - fechaActual) / (1000 * 60 * 60);
      
      // Regla: Máximo 1 hora antes
      if (diferenciaHoras < 1) {
        return respuestaExito({ mensaje: "No se puede editar. Falta menos de 1 hora para la actividad." }, false);
      }
      
      // Editar campos (Actualizar Fila)
      hojaCitas.getRange(i + 1, 3).setValue(data.titulo);
      hojaCitas.getRange(i + 1, 4).setValue(data.fechaHora);
      hojaCitas.getRange(i + 1, 5).setValue(data.detalles);
      
      registrarBitacora(sheet, data.idUsuario, "EDITAR_CITA", `Cita ${data.idCita} editada a ${data.fechaHora}`);
      return respuestaExito({ mensaje: "Cita editada correctamente" }, true);
    }
  }
  return respuestaExito({ mensaje: "Cita no encontrada o no tienes permisos" }, false);
}

// 3. OBTENER CITAS (Admin ve todas, Usuario ve las suyas)
function obtenerCitas(data, sheet) {
  const hojaCitas = sheet.getSheetByName('Citas');
  const datos = hojaCitas.getDataRange().getValues();
  let resultados = [];
  
  for (let i = 1; i < datos.length; i++) {
    if (data.rol === 'admin' || datos[i][1] === data.idUsuario) {
      resultados.push({
        idCita: datos[i][0],
        idUsuario: datos[i][1],
        titulo: datos[i][2],
        fechaHora: datos[i][3],
        detalles: datos[i][4]
      });
    }
  }
  return respuestaExito({ citas: resultados }, true);
}

// 4. CREAR USUARIO (Solo Admin)
function crearUsuario(data, sheet) {
  if (data.rolActual !== 'admin') return respuestaExito({ mensaje: "Sin permisos" }, false);
  
  const hojaUsuarios = sheet.getSheetByName('Usuarios');
  const idNuevo = Utilities.getUuid();
  hojaUsuarios.appendRow([idNuevo, data.nuevoUsuario, data.nuevaContrasena, 'user']);
  
  registrarBitacora(sheet, data.idUsuario, "CREAR_USUARIO", `Usuario ${data.nuevoUsuario} creado.`);
  return respuestaExito({ mensaje: "Usuario generado" }, true);
}

// --- FUNCIONES DE SOPORTE ---

function registrarBitacora(sheet, idUsuario, accion, detalles) {
  const hojaBitacora = sheet.getSheetByName('Bitacora');
  const fecha = new Date();
  hojaBitacora.appendRow([fecha, idUsuario, accion, detalles]);
}

function respuestaExito(objeto, status) {
  objeto.success = status;
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
/*
 * RESERVAS DE TABLETS
 * Backend unico para Apps Script (Google Sheets).
 *
 * - doGet: entrega el horario semanal como CSV al frontend.
 * - doPost: registra la nueva solicitud en 'Solicitudes' y notifica al admin.
 * - enviarCorreoAlEditar: trigger onEdit, envia correo aprobacion/rechazo.
 */

const MAX_TABLETS = 60;
const CONFIG_EMAIL_ADMIN = "TU_CORREO_ADMIN@COLEGIO.CL";
const DIAS_NOMBRES = ["Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado"];
const MESES_NOMBRES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.accion === 'proximaSemana') {
      return obtenerSolicitudesProximaSemana();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Horario semanal") || ss.getSheets()[0];
    var range = sheet.getDataRange();
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();

    var headerRow = -1, colHora = -1;
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var val = values[r][c] ? String(values[r][c]).toLowerCase().trim() : '';
        if (val === 'hora') { headerRow = r; colHora = c; break; }
      }
      if (headerRow !== -1) break;
    }
    if (headerRow === -1) {
      return ContentService.createTextOutput('Error: no header')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    var WHITE = ['#ffffff', '#fff'];
    var csvRows = [];

    for (var r = 0; r < values.length; r++) {
      var row = [];
      for (var c = 0; c < values[r].length; c++) {
        var v = values[r][c] !== null ? String(values[r][c]) : '';
        var bg = backgrounds[r][c] ? backgrounds[r][c].toLowerCase() : '#ffffff';

        if (r > headerRow && c > colHora && v.trim() !== '' && bg !== '#ffffff' && bg !== '#fff') {
          v = '\u00A7' + v;
        }

        if (v.indexOf(',') !== -1 || v.indexOf('"') !== -1 || v.indexOf('\n') !== -1 || v.indexOf('\r') !== -1) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        row.push(v);
      }
      csvRows.push(row.join(','));
    }

    return ContentService.createTextOutput(csvRows.join('\n'))
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return responderJSON({
      status: 'error',
      message: 'El servidor esta ocupado. Intentalo de nuevo en unos segundos.'
    });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se recibieron datos validos en el envio.");
    }

    var data = JSON.parse(e.postData.contents);
    var fechaRaw = String(data.fecha || "").trim();
    var cursoSol = String(data.curso || "").trim();
    var profSol = String(data.profesor || "").trim();
    var cantSol = parseInt(data.tablets) || 0;
    var horaInicioSol = String(data.horaInicio || "").trim();
    var horaFinSol = String(data.horaFin || "").trim();

    if (!fechaRaw) throw new Error("Debes seleccionar una Fecha.");
    if (cantSol <= 0) throw new Error("La cantidad de tablets debe ser mayor a 0.");
    if (!horaInicioSol || !horaFinSol) throw new Error("Debes indicar 'Hora Inicio' y 'Hora Fin' validas.");

    var partesFecha = fechaRaw.split('-');
    if (partesFecha.length !== 3) throw new Error("Formato de fecha invalido.");

    var anio = parseInt(partesFecha[0], 10);
    var mes = parseInt(partesFecha[1], 10) - 1;
    var dia = parseInt(partesFecha[2], 10);
    var fechaObj = new Date(anio, mes, dia);

    var diasNombres = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"];
    var numDiaSemana = fechaObj.getDay();

    if (numDiaSemana === 0 || numDiaSemana === 6) {
      throw new Error("No se pueden hacer reservas para fines de semana.");
    }

    var diaSol = diasNombres[numDiaSemana];
    var fechaTexto = String(dia).padStart(2, '0') + '/' + String(mes + 1).padStart(2, '0') + '/' + anio;

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheetHorario = ss.getSheetByName("Horario semanal") || ss.getSheets()[0];
    var valuesHorario = sheetHorario.getDataRange().getValues();

    var headerRow = -1;
    var colHora = -1;

    for (var r = 0; r < valuesHorario.length; r++) {
      for (var c = 0; c < valuesHorario[r].length; c++) {
        var val = valuesHorario[r][c] ? String(valuesHorario[r][c]).toLowerCase().trim() : '';
        if (val === 'hora') {
          headerRow = r;
          colHora = c;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1) {
      throw new Error("No se encontro la celda 'Hora' en la pestana 'Horario semanal'.");
    }

    var colDia = -1;
    var diaSolClean = diaSol.toLowerCase();

    for (var c = colHora; c < valuesHorario[headerRow].length; c++) {
      var headerVal = valuesHorario[headerRow][c] ? String(valuesHorario[headerRow][c]).toLowerCase().trim() : '';
      if (headerVal && (diaSolClean.includes(headerVal) || headerVal.includes(diaSolClean))) {
        colDia = c;
        break;
      }
    }

    if (colDia === -1) throw new Error("No se encontro la columna para el dia '" + diaSol + "'.");

    var minInicioSol = horaAMinutos(horaInicioSol);
    var minFinSol = horaAMinutos(horaFinSol);

    if (minInicioSol === null || minFinSol === null || minInicioSol >= minFinSol) {
      throw new Error("El rango de horas seleccionado (" + horaInicioSol + " a " + horaFinSol + ") no es valido.");
    }

    for (var r = headerRow + 1; r < valuesHorario.length; r++) {
      var horaCeldaRaw = valuesHorario[r][colHora] ? String(valuesHorario[r][colHora]).trim() : '';

      if (!horaCeldaRaw || horaCeldaRaw.toLowerCase().includes('almuerzo')) continue;

      var partesBloque = horaCeldaRaw.split('-');
      if (partesBloque.length < 2) continue;

      var minBloqueInicio = horaAMinutos(partesBloque[0]);
      var minBloqueFin = horaAMinutos(partesBloque[1]);

      if (minBloqueInicio === null || minBloqueFin === null) continue;

      var haySolapamiento = (minBloqueInicio < minFinSol) && (minBloqueFin > minInicioSol);

      if (haySolapamiento) {
        var textoBloqueCompleto = "";
        var rowIdx = r;

        while (rowIdx < valuesHorario.length) {
          var hFila = valuesHorario[rowIdx][colHora] ? String(valuesHorario[rowIdx][colHora]).trim() : '';
          if (rowIdx > r && hFila !== '') break;

          var esAlm = valuesHorario[rowIdx].some(function(cell) {
            return cell && String(cell).toLowerCase().includes('almuerzo');
          });
          if (esAlm) break;

          var contenidoCelda = valuesHorario[rowIdx][colDia] ? String(valuesHorario[rowIdx][colDia]) : '';
          if (contenidoCelda.trim() !== '') {
            textoBloqueCompleto += "\n" + contenidoCelda;
          }

          rowIdx++;
        }

        var ocupadas = 0;
        var coincidencias = textoBloqueCompleto.match(/(\d+)\s*TABLET/gi);
        if (coincidencias) {
          coincidencias.forEach(function(item) {
            var num = parseInt(item.match(/\d+/)[0]);
            if (!isNaN(num)) ocupadas += num;
          });
        }

        var disponibles = MAX_TABLETS - ocupadas;

        if (cantSol > disponibles) {
          return responderJSON({
            status: 'error',
            message: 'Sin cupo suficiente el dia ' + diaSol + ' (' + horaCeldaRaw + '). Pediste ' + cantSol + ' tablets, pero solo quedan ' + disponibles + ' disponibles.'
          });
        }
      }
    }

    var sheetSolicitudes = ss.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
    if (!sheetSolicitudes) throw new Error("No se encontro la pestana 'Solicitudes'.");

    var colAValues = sheetSolicitudes.getRange("A1:A" + sheetSolicitudes.getMaxRows()).getValues();
    var primeraFilaLibre = 2;

    for (var i = 1; i < colAValues.length; i++) {
      if (!colAValues[i][0] || String(colAValues[i][0]).trim() === "") {
        primeraFilaLibre = i + 1;
        break;
      }
    }

    var ultimaColumna = sheetSolicitudes.getLastColumn();

    if (primeraFilaLibre > 2) {
      var rangoOrigen = sheetSolicitudes.getRange(primeraFilaLibre - 1, 1, 1, ultimaColumna);
      var rangoDestino = sheetSolicitudes.getRange(primeraFilaLibre, 1, 1, ultimaColumna);
      rangoOrigen.copyTo(rangoDestino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      rangoDestino.clearContent();
    }

    var datosNuevos = [
      profSol,
      cantSol,
      fechaTexto,
      cursoSol,
      horaInicioSol,
      horaFinSol
    ];

    sheetSolicitudes.getRange(primeraFilaLibre, 1, 1, datosNuevos.length).setValues([datosNuevos]);

    // --- Notificar al admin de la nueva solicitud ---
    try {
      var linkHoja = ss.getUrl();
      var asuntoAdmin = "Nueva reserva de Tablets: " + profSol + " - " + fechaTexto;
      var cuerpoAdmin = "Nueva solicitud de reserva de Tablets.\n\n"
        + "Profesor: " + profSol + "\n"
        + "Cantidad de tablets: " + cantSol + "\n"
        + "Curso: " + cursoSol + "\n"
        + "Fecha: " + fechaTexto + " (" + diaSol + ")\n"
        + "Horario: " + horaInicioSol + " a " + horaFinSol + "\n"
        + "Estado: Pendiente de aprobacion\n\n"
        + "Para aprobar o rechazar, abre la hoja de Solicitudes:\n" + linkHoja;
      MailApp.sendEmail(CONFIG_EMAIL_ADMIN, asuntoAdmin, cuerpoAdmin);
    } catch (errNotif) {
      Logger.log("Error al notificar al admin: " + errNotif.message);
    }

    return responderJSON({
      status: 'success',
      message: 'Solicitud registrada con exito para el dia ' + diaSol + ' (' + fechaTexto + ').'
    });

  } catch (err) {
    return responderJSON({
      status: 'error',
      message: 'Error en el servidor: ' + err.message
    });
  } finally {
    lock.releaseLock();
  }
}

function horaAMinutos(horaStr) {
  if (!horaStr) return null;
  var partes = String(horaStr).replace(/\s/g, '').split(':');
  if (partes.length < 2) return null;
  var h = parseInt(partes[0], 10);
  var m = parseInt(partes[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function responderJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerSolicitudesProximaSemana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
  if (!sheet) throw new Error("No se encontro la pestana 'Solicitudes'.");

  var hoy = new Date();
  var d = hoy.getDay();
  var diasAlLunes = (d === 0) ? 1 : 8 - d;
  var lunes = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() + diasAlLunes);
  var viernes = new Date(lunes.getFullYear(), lunes.getMonth(), lunes.getDate() + 4);

  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return responderJSON([]);

  var datos = sheet.getRange(2, 1, ultimaFila - 1, 7).getValues();
  var resultados = [];

  for (var i = 0; i < datos.length; i++) {
    var fechaCelda = datos[i][2];
    var fecha = (fechaCelda instanceof Date) ? fechaCelda : parseFechaTexto(String(fechaCelda));
    if (!fecha) continue;
    fecha = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
    if (fecha.getTime() >= lunes.getTime() && fecha.getTime() <= viernes.getTime()) {
      resultados.push({
        dia: DIAS_NOMBRES[fecha.getDay()],
        fecha: String(fecha.getDate()).padStart(2,'0') + " de " + MESES_NOMBRES[fecha.getMonth()],
        horaInicio: formatearHora(datos[i][4]),
        horaFin: formatearHora(datos[i][5]),
        profesor: String(datos[i][0] || "").trim(),
        curso: String(datos[i][3] || "").trim(),
        cantidad: datos[i][1],
        estado: String(datos[i][6] || "").trim()
      });
    }
  }

  resultados.sort(function(a, b) {
    var idx = { Lunes:1, Martes:2, Miercoles:3, Jueves:4, Viernes:5 };
    var d1 = idx[a.dia] || 99, d2 = idx[b.dia] || 99;
    if (d1 !== d2) return d1 - d2;
    return horaAMinutos(a.horaInicio) - horaAMinutos(b.horaInicio);
  });

  return responderJSON(resultados);
}

function parseFechaTexto(str) {
  if (!str) return null;
  var p = String(str).trim().split('/');
  if (p.length !== 3) return null;
  return new Date(parseInt(p[2],10), parseInt(p[1],10) - 1, parseInt(p[0],10));
}

function formatearHora(valor) {
  if (valor instanceof Date) {
    var h = valor.getHours();
    var m = valor.getMinutes();
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  if (typeof valor === 'number' && valor > 0 && valor < 1) {
    var totalMin = Math.round(valor * 24 * 60);
    var h = Math.floor(totalMin / 60);
    var m = totalMin % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  return String(valor || "").trim().replace('.', ':');
}


// ========================================================
// GESTION DE CORREOS AL EDITAR LA HOJA (onEdit) Y HORARIO
// ========================================================

/**
 * @OnlyCurrentDoc
 * Este script se activa al editar la hoja. Si se marca una casilla en la columna de envÃ­o,
 * busca el correo de un profesor y envÃ­a una notificaciÃ³n.
 * Si hay mÃºltiples coincidencias, muestra un selector. Si hay un error, lo registra.
 */

// --- CONFIGURACIÃ“N DE HOJAS Y COLUMNAS ---
const NOMBRE_HOJA_SOLICITUDES = "Solicitudes";
const NOMBRE_HOJA_PROFESORES = "Profes";
const COLUMNAS_SOLICITUDES = {
  NOMBRE_PROFE: 1,
  CANTIDAD_TABLET: 2,
  FECHA: 3,
  CURSO: 4,
  HORA_PRINCIPIO: 5,
  HORA_FIN: 6,
  ESTADO: 7,
  COMENTARIOS: 8,
  CHECKBOX: 10,
  ERRORES: 11
};
const COLUMNAS_PROFESORES = {
  NOMBRE: 1,
  APELLIDO: 2,
  CORREO: 3
};
// --- FIN DE LA CONFIGURACIÃ“N ---

/**
 * Carga el diccionario de profesores desde la hoja 'Profes'.
 * El diccionario tendrÃ¡ el formato "Nombre Apellido": "correo" o "Nombre": "correo".
 * @returns {Object} El diccionario de profesores.
 */
function cargarDiccionarioDesdeProfesores() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaProfesores = spreadsheet.getSheetByName(NOMBRE_HOJA_PROFESORES);

  if (!hojaProfesores) {
    throw new Error(`La hoja "${NOMBRE_HOJA_PROFESORES}" no se encontrÃ³.`);
  }

  // Obtiene todos los datos de la hoja, omitiendo la fila de encabezados.
  const valores = hojaProfesores.getRange(2, 1, hojaProfesores.getLastRow() - 1, hojaProfesores.getLastColumn()).getValues();

  const diccionario = {};
  for (let i = 0; i < valores.length; i++) {
    const nombre = valores[i][COLUMNAS_PROFESORES.NOMBRE - 1].toString().trim();
    const apellido = valores[i][COLUMNAS_PROFESORES.APELLIDO - 1].toString().trim();
    const correo = valores[i][COLUMNAS_PROFESORES.CORREO - 1].toString().trim();

    if (nombre && correo) {
      // Crea la clave completa con nombre y apellido
      const nombreCompleto = apellido ? `${nombre} ${apellido}` : nombre;
      diccionario[nombreCompleto.toLowerCase()] = {
        nombreCompleto: nombreCompleto,
        correo: correo
      };
      
      // Si hay un apellido, tambiÃ©n agrega una entrada solo con el nombre
      // Esto permite que la bÃºsqueda sea mÃ¡s flexible
      if (apellido) {
        diccionario[nombre.toLowerCase()] = {
          nombreCompleto: nombreCompleto,
          correo: correo
        };
      }
    }
  }
  return diccionario;
}

/**
 * FunciÃ³n principal activada por la ediciÃ³n de la hoja.
 */
function enviarCorreoAlEditar(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();

  if (sheet.getName() !== NOMBRE_HOJA_SOLICITUDES) return;
  if (col !== COLUMNAS_SOLICITUDES.CHECKBOX) return;
  if (row <= 1) return;

  var valor = range.getValue();
  if (valor !== true) return;

  var celdaEstado = sheet.getRange(row, COLUMNAS_SOLICITUDES.ERRORES);
  var estadoActual = String(celdaEstado.getValue() || "").trim().toLowerCase();
  if (estadoActual === "procesando..." || estadoActual.indexOf("enviado") !== -1) return;

  celdaEstado.setValue("Procesando...").setFontColor("#E67E22").setBackground(null);

  try {
    var nombreProfe = sheet.getRange(row, COLUMNAS_SOLICITUDES.NOMBRE_PROFE).getValue();
    var diccionarioProfesores = cargarDiccionarioDesdeProfesores();
    var coincidencias = buscarCoincidencias(nombreProfe, diccionarioProfesores);

    if (coincidencias.length === 0) {
    
      var ui= SpreadsheetApp.getUi();
      var resp= ui.alert("Profesor no Encontrado ", "No se encontrÃ³ correo para " + nombreProfe +
                        ".\n Â¿Deseas agregarlo ahora?", ui.ButtonSet.YES_NO );
      if (resp != ui.Button.YES){
        reescribirHorario(row);
        celdaEstado.setValue("Agendado sin correo").setFontColor("#E67E22").setBackground("#FEF3C7");
        return;

      }
       var inputNombre = ui.prompt('Nombre del profesor', 'Ej: Juan', ui.ButtonSet.OK_CANCEL);
    if (inputNombre.getSelectedButton() !== ui.Button.OK) { 
      reescribirHorario(row);
      celdaEstado.setValue("Agendado sin correo").setFontColor("#E67E22").setBackground("#FEF3C7");
      return; 
      }
    var nombre = inputNombre.getResponseText().trim();

    var inputApellido = ui.prompt('Apellido del profesor', 'Ej: PÃ©rez', ui.ButtonSet.OK_CANCEL);
    if (inputApellido.getSelectedButton() !== ui.Button.OK){ 
      reescribirHorario(row);
      celdaEstado.setValue("Agendado sin correo").setFontColor("#E67E22").setBackground("#FEF3C7");
      return; 
    }
    var apellido = inputApellido.getResponseText().trim();

    var inputCorreo = ui.prompt('Correo electrÃ³nico', 'Ej: juan@colegio.cl', ui.ButtonSet.OK_CANCEL);
    if (inputCorreo.getSelectedButton() !== ui.Button.OK) 
    { 
      reescribirHorario(row);
      celdaEstado.setValue("Agendado sin correo").setFontColor("#E67E22").setBackground("#FEF3C7");
      return; 
     }
    var correo = inputCorreo.getResponseText().trim();

    if (!correo || correo.indexOf('@') === -1) {
      reescribirHorario(row);
      celdaEstado.setValue("Agendado sin correo").setFontColor("#E67E22").setBackground("#FEF3C7");
      ui.alert('Error', 'El correo ingresado no es vÃ¡lido. La reserva se agendÃ³ sin envÃ­o de correo.', ui.ButtonSet.OK);
      return;
    }

    agregarProfesor(nombre, apellido, correo);
    enviarEmail(correo, row, celdaEstado);

  } else if (coincidencias.length === 1) {
    enviarEmail(coincidencias[0].correo, row, celdaEstado);
  } else {
  mostrarSelectorUI(coincidencias, row);


    }

  } catch (error) {
    celdaEstado.setValue("Error: " + error.message).setFontColor("#DC2626").setBackground(null);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

/**
 * Busca coincidencias parciales de nombres en el diccionario.
 * @param {string} nombreBuscado El nombre del profesor de la hoja.
 * @param {Object} diccionario El diccionario de profesores.
 * @returns {Array<Object>} Un array de objetos con nombre y correo.
 */
function buscarCoincidencias(nombreBuscado, diccionario) {
  const coincidencias = [];
  const nombreNormalizado = nombreBuscado.toString().trim().toLowerCase();
  
  for (const clave in diccionario) {
    if (clave.includes(nombreNormalizado)) {
      coincidencias.push({
        nombre: diccionario[clave].nombreCompleto,
        correo: diccionario[clave].correo
      });
    }
  }

  // Eliminar duplicados basÃ¡ndose en el correo electrÃ³nico.
  const unicas = {};
  return coincidencias.filter(item => {
    // Si el correo ya existe en el objeto 'unicas', es un duplicado.
    if (unicas.hasOwnProperty(item.correo)) {
      return false;
    }
    // Si no, lo agregamos y lo mantenemos en el resultado.
    unicas[item.correo] = true;
    return true;
  });
}

/**
 * Muestra el selector de opciones al usuario.
 * @param {Array<Object>} opciones
 * @param {number} rowNum
 */
function mostrarSelectorUI(opciones, rowNum) {
  const plantilla = HtmlService.createTemplateFromFile('SelectorUI');
  plantilla.opciones = opciones;
  plantilla.rowNum = rowNum;
  const html = plantilla.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Elige el profesor');
}

/**
 * Esta funciÃ³n es llamada desde el HTML para procesar la selecciÃ³n del usuario.
 * @param {string} correoSeleccionado El correo que el usuario ha elegido.
 * @param {number} rowNum El nÃºmero de fila del registro a procesar.
 */
function procesarSeleccion(correoSeleccionado, rowNum) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
    var celdaEstado = sheet.getRange(rowNum, COLUMNAS_SOLICITUDES.ERRORES);

    enviarEmail(correoSeleccionado, rowNum, celdaEstado);
    return "Correo enviado a " + correoSeleccionado;
  } catch (error) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_SOLICITUDES);
    sheet.getRange(rowNum, COLUMNAS_SOLICITUDES.ERRORES).setValue("Error: " + error.message).setFontColor("#DC2626").setBackground(null);
    return "Error: " + error.message;
  }
}

/**
 * FunciÃ³n que encapsula toda la lÃ³gica de obtenciÃ³n de datos y envÃ­o de correo.
 * @param {string} email El correo del destinatario.
 * @param {number} rowNum El nÃºmero de fila de donde se obtienen los datos.
 * @param {Object} celdaEstado La celda de columna K para escribir el resultado.
 */
function enviarEmail(email, rowNum, celdaEstado) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
  const rowData = sheet.getRange(rowNum, 1, 1, 10).getValues()[0];
  const nombreProfe = rowData[COLUMNAS_SOLICITUDES.NOMBRE_PROFE - 1];
  const estado = rowData[COLUMNAS_SOLICITUDES.ESTADO - 1].toString().trim().toLowerCase();

  const cantidadTablet = rowData[COLUMNAS_SOLICITUDES.CANTIDAD_TABLET - 1];
  const fecha = Utilities.formatDate(new Date(rowData[COLUMNAS_SOLICITUDES.FECHA - 1]), spreadsheet.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  const curso = rowData[COLUMNAS_SOLICITUDES.CURSO - 1];
  const horaPrincipio = Utilities.formatDate(new Date(rowData[COLUMNAS_SOLICITUDES.HORA_PRINCIPIO - 1]), spreadsheet.getSpreadsheetTimeZone(), "HH:mm");
  const horaFin = Utilities.formatDate(new Date(rowData[COLUMNAS_SOLICITUDES.HORA_FIN - 1]), spreadsheet.getSpreadsheetTimeZone(), "HH:mm");
  const comentarios = rowData[COLUMNAS_SOLICITUDES.COMENTARIOS - 1] || "Sin comentarios.";

  let asunto = "";
  let cuerpo = "";

  if (estado === 'sÃ­' || estado === 'si') {
    asunto = `Solicitud de Tablets APROBADA - ${curso}`;
    cuerpo = `Hola ${nombreProfe},\n\nTu solicitud de tablets ha sido APROBADA.\n\nDetalles:\n- Curso: ${curso}\n- Cantidad: ${cantidadTablet}\n- Fecha: ${fecha}\n- Horario: De ${horaPrincipio} a ${horaFin}\n\nComentarios: ${comentarios}\n\nSaludos,\nEquipo de TI`;
  } else if (estado === 'no') {
    asunto = `Solicitud de Tablets RECHAZADA - ${curso}`;
    cuerpo = `Hola ${nombreProfe},\n\nLamentamos informarte que tu solicitud de tablets ha sido RECHAZADA.\n\nDetalles:\n- Curso: ${curso}\n- Cantidad: ${cantidadTablet}\n- Fecha: ${fecha}\n- Horario: De ${horaPrincipio} a ${horaFin}\n\nMotivo: ${comentarios}\n\nSaludos,\nEquipo de TI`;
  } else {
    throw new Error("El estado en la columna 'G' debe ser 'SÃ­' o 'No'.");
  }

  MailApp.sendEmail(email, asunto, cuerpo, { name: 'RESERVA TABLETS' });

  if (estado === 'sÃ­' || estado === 'si') {
    reescribirHorario(rowNum);
  }

  celdaEstado.setValue("Enviado âœ“").setFontColor("#16A34A").setBackground("#DCFCE7").setNote("Enviado el " + new Date());
  spreadsheet.toast("Correo enviado a " + nombreProfe, "Ã‰xito", 5);
}

/**
 * FunciÃ³n para manejar los errores y actualizar la hoja de cÃ¡lculo.
 * @param {string} message Mensaje de error.
 * @param {Object} celdaEstado La celda de columna K para escribir el error.
 * @param {Object} spreadsheet El objeto de la hoja de cÃ¡lculo.
 */
function handleError(message, celdaEstado, spreadsheet) {
  celdaEstado.setValue(message).setFontColor("#DC2626").setBackground(null);
  spreadsheet.toast(message, "Error", 10);
}

// --- CONFIGURACIÃ“N DE HOJA "HORARIO SEMANAL" ---
const NOMBRE_HOJA_HORARIO = "Horario semanal";
const FILA_HEADER_HORARIO = 5;
const COL_HORA_HORARIO = 2;
const MAPEO_DIAS = { 1: 3, 2: 4, 3: 5, 4: 6, 5: 7 };

function reescribirHorario(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSol = ss.getSheetByName(NOMBRE_HOJA_SOLICITUDES);
  var sheetHorario = ss.getSheetByName(NOMBRE_HOJA_HORARIO);
  if (!sheetHorario) { Logger.log("No se encontrÃ³ la hoja '" + NOMBRE_HOJA_HORARIO + "'"); return; }

  var rowData = sheetSol.getRange(rowNum, 1, 1, 7).getValues()[0];
  var nombreProfe = String(rowData[0] || "").trim();
  var cantidad = rowData[1];
  var fechaRaw = rowData[2];
  var curso = String(rowData[3] || "").trim();
  var horaInicioRaw = rowData[4];
  var horaFinRaw = rowData[5];

  var fecha = new Date(fechaRaw);
  var numDia = fecha.getDay();
  var colDia = MAPEO_DIAS[numDia];
  if (!colDia) { Logger.log("Fin de semana, no se escribe en horario."); return; }

  var horaInicioDate = new Date(horaInicioRaw);
  var horaFinDate = new Date(horaFinRaw);
  var minInicio = horaInicioDate.getHours() * 60 + horaInicioDate.getMinutes();
  var minFin = horaFinDate.getHours() * 60 + horaFinDate.getMinutes();

  if (minInicio >= minFin) { Logger.log("Rango de horas invÃ¡lido."); return; }

  var cantNum = parseInt(cantidad);
  var textoReserva;
  if (curso) {
    textoReserva = (nombreProfe + " " + curso + " " + (isNaN(cantNum) ? cantidad : cantNum) + " TABLETS").toUpperCase();
    
  } else {
    textoReserva = (nombreProfe + " " + (isNaN(cantNum) ? cantidad : cantNum) + " TABLETS").toUpperCase();
  }
  var lastRow = sheetHorario.getLastRow();
  if (lastRow <= FILA_HEADER_HORARIO) return;
  var data = sheetHorario.getRange(FILA_HEADER_HORARIO + 1, COL_HORA_HORARIO, lastRow - FILA_HEADER_HORARIO, 6).getValues();

  var bloques = [];
  for (var r = 0; r < data.length; r++) {
    var celdaHora = data[r][0];
    if (!celdaHora) continue;
    var textoHora = String(celdaHora).trim();
    if (textoHora === "" || textoHora.toLowerCase().indexOf("almuerzo") !== -1) continue;

    var partes = textoHora.split("-");
    if (partes.length < 2) continue;

    var minBloqueInicio =HoraAMinutosApp(partes[0].trim());
    var minBloqueFin = HoraAMinutosApp(partes[1].trim());
    if (minBloqueInicio === null || minBloqueFin === null) continue;

    var finBloque = r + 1;
    while (finBloque < data.length) {
      var nextHora = data[finBloque][0];
      if (!nextHora) { finBloque++; continue; }
      var nextTexto = String(nextHora).trim();
      if (nextTexto === "") { finBloque++; continue; }
      if (nextTexto.toLowerCase().indexOf("almuerzo") !== -1) break;
      if (nextTexto.indexOf("-") !== -1) break;
      finBloque++;
    }

    bloques.push({
      filaSheet: FILA_HEADER_HORARIO + 1 + r,
      filasHasta: FILA_HEADER_HORARIO + 1 + finBloque,
      minInicio: minBloqueInicio,
      minFin: minBloqueFin
    });
  }

  var escritas = 0;
  for (var b = 0; b < bloques.length; b++) {
    var bloque = bloques[b];
    if (minInicio >= bloque.minFin || minFin <= bloque.minInicio) continue;

    var escritaEnBloque = false;
    for (var rf = bloque.filaSheet; rf < bloque.filasHasta; rf++) {
      var celda = sheetHorario.getRange(rf, colDia);
      var valor = celda.getValue();
      if (!valor || String(valor).trim() === "") {
        celda.setValue(textoReserva);
        celda.setFontFamily("Roboto");
        celda.setFontWeight("bold");
        celda.setFontSize(11);
        celda.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        escritaEnBloque = true;
        escritas++;
        break;
      }
    }
    if (!escritaEnBloque) {
      var ultimaFila = sheetHorario.getRange(bloque.filasHasta - 1, colDia);
      var existente = String(ultimaFila.getValue() || "").trim();
      if (existente) {
        ultimaFila.setValue(existente + "\n" + textoReserva);
      } else {
        ultimaFila.setValue(textoReserva);
      }
      ultimaFila.setFontFamily("Roboto");
      ultimaFila.setFontWeight("bold");
      ultimaFila.setFontSize(11);
      ultimaFila.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      escritas++;
    }
  }

  if (escritas > 0) {
    ss.toast("Reserva escrita en " + escritas + " bloque(s) del horario", "Ã‰xito", 5);
  } else {
    var diasSemana = ["Domingo", "Lunes", "Martes", "MiÃ©rcoles", "Jueves", "Viernes", "SÃ¡bado"];
    var msg = "No se encontraron bloques horarios solapados para " + diasSemana[numDia] + " " + horaInicioDate.getHours() + ":" + ("0" + horaInicioDate.getMinutes()).slice(-2) + " - " + horaFinDate.getHours() + ":" + ("0" + horaFinDate.getMinutes()).slice(-2);
    ss.toast(msg, "Aviso", 8);
  }
}

function HoraAMinutosApp(horaStr) {
  if (!horaStr) return null;
  var partes = String(horaStr).replace(/\s/g, "").split(":");
  if (partes.length < 2) return null;
  var h = parseInt(partes[0], 10);
  var m = parseInt(partes[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function agregarProfesor(nombre, apellido, correo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_HOJA_PROFESORES);
  if (!sheet) throw new Error('No se encontrÃ³ la hoja "' + NOMBRE_HOJA_PROFESORES + '".');
  sheet.appendRow([nombre, apellido, correo]);
}

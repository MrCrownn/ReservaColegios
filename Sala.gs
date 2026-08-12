// === CONFIGURACION ===
const NOMBRE_HOJA_SOLICITUDES = "Solicitudes";
const NOMBRE_HOJA_HORARIO = "Horario";
const NOMBRE_HOJA_PROFESORES = "Profes";
const DIAS_BUSCAR = ['lunes', 'martes', 'miercoles', 'miércoles', 'jueves', 'viernes'];
const CONFIG_EMAIL_ADMIN = "TU_CORREO_ADMIN@COLEGIO.CL";
const DIAS_NOMBRES = ["Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado"];
const MESES_NOMBRES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const COLUMNAS_SOLICITUDES = {
  PROFESOR: 1,        // A
  CURSO: 2,           // B
  FECHA: 3,           // C
  HORA_INICIO: 4,     // D
  HORA_FIN: 5,        // E
  RESERV_ACEPTADA: 6, // F
  OBSERVACIONES: 7    // G
};

function buscarHoja(ss, nombreBuscado) {
  var hojas = ss.getSheets();
  var nombreNorm = nombreBuscado.trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  for (var i = 0; i < hojas.length; i++) {
    var nombreHoja = hojas[i].getName().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (nombreHoja === nombreNorm) return hojas[i];
  }
  return null;
}

// === doGet: Genera CSV del horario ===
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.accion === 'proximaSemana') {
      return obtenerSolicitudesProximaSemana();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = buscarHoja(ss, NOMBRE_HOJA_HORARIO);
    if (!sheet) {
      var nombres = ss.getSheets().map(function(h) { return h.getName(); }).join(', ');
      return ContentService.createTextOutput('Error: no se encontro la hoja "' + NOMBRE_HOJA_HORARIO + '". Hojas disponibles: ' + nombres)
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var range = sheet.getDataRange();
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();

    var headerRowDays = -1;
    for (var r = 0; r < values.length; r++) {
      var dayCount = 0;
      for (var c = 0; c < values[r].length; c++) {
        var val = values[r][c] ? String(values[r][c]).toLowerCase().trim() : '';
        for (var d = 0; d < DIAS_BUSCAR.length; d++) {
          if (val === DIAS_BUSCAR[d] || val.indexOf(DIAS_BUSCAR[d]) === 0) { dayCount++; break; }
        }
      }
      if (dayCount >= 3) { headerRowDays = r; break; }
    }

    if (headerRowDays === -1) {
      return ContentService.createTextOutput('Error: no header')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    var headerRowDates = headerRowDays - 1;
    var dataStartRow = headerRowDays + 1;
    var colHora = 1;
    var colDiasInicio = 2;

    var csvRows = [];

    for (var r = 0; r < dataStartRow; r++) {
      var row = [];
      for (var c = 0; c < values[r].length; c++) {
        var v = values[r][c] !== null ? String(values[r][c]) : '';
        if (r === headerRowDays && c === colHora && (v === '' || v.toLowerCase().indexOf('hora') === -1)) {
          v = 'Hora';
        }
        if (v.indexOf(',') !== -1 || v.indexOf('"') !== -1 || v.indexOf('\n') !== -1) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        row.push(v);
      }
      csvRows.push(row.join(','));
    }

    for (var r = dataStartRow; r < values.length; r++) {
      var celdaHora = values[r][colHora] ? String(values[r][colHora]).trim() : '';
      if (celdaHora === '' || celdaHora.toLowerCase().indexOf('almuerzo') !== -1) {
        var row = [ '', celdaHora ];
        for (var c = colDiasInicio; c < values[r].length; c++) {
          var v = values[r][c] !== null ? String(values[r][c]) : '';
          row.push(v);
        }
        csvRows.push(row.join(','));
        continue;
      }

      var row = [ '', celdaHora ];
      for (var c = colDiasInicio; c < values[r].length; c++) {
        var v = values[r][c] !== null ? String(values[r][c]) : '';
        var bg = backgrounds[r][c] ? backgrounds[r][c].toLowerCase() : '#ffffff';

        if (v.trim() !== '' && (bg === '#b6d7a8' || bg === '#0f9d58')) {
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

// === doPost: Recibe reserva del formulario ===
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se recibieron datos validos en el envio.");
    }

    var data = JSON.parse(e.postData.contents);
    var profesor = String(data.profesor || "").trim();
    var curso = String(data.curso || "").trim();
    var fechaRaw = String(data.fecha || "").trim();
    var horaInicio = String(data.horaInicio || "").trim();
    var horaFin = String(data.horaFin || "").trim();
    var observaciones = String(data.observaciones || "").trim();

    if (!profesor) throw new Error("Debes ingresar el nombre del profesor.");
    if (!curso) throw new Error("Debes ingresar el curso.");
    if (!fechaRaw) throw new Error("Debes seleccionar una fecha.");
    if (!horaInicio || !horaFin) throw new Error("Debes indicar hora de inicio y fin.");

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

    var fechaTexto = String(dia).padStart(2, '0') + '/' + String(mes + 1).padStart(2, '0') + '/' + anio;

    var minSolicitados = horaAMinutos(horaInicio);
    var minFinSolicitados = horaAMinutos(horaFin);
    if (minSolicitados === null || minFinSolicitados === null || minSolicitados >= minFinSolicitados) {
      throw new Error("El rango de horas seleccionado no es valido.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Verificar solapamiento en "Solicitudes" ---
    var sheetSolicitudes = buscarHoja(ss, NOMBRE_HOJA_SOLICITUDES);
    if (!sheetSolicitudes) throw new Error("No se encontro la pestana 'Solicitudes'.");

    var ultimaFila = sheetSolicitudes.getLastRow();
    if (ultimaFila >= 2) {
      var datosSol = sheetSolicitudes.getRange(2, 1, ultimaFila - 1, 7).getValues();
      var displayHorasIni = sheetSolicitudes.getRange(2, COLUMNAS_SOLICITUDES.HORA_INICIO, ultimaFila - 1, 1).getDisplayValues();
      var displayHorasFin = sheetSolicitudes.getRange(2, COLUMNAS_SOLICITUDES.HORA_FIN, ultimaFila - 1, 1).getDisplayValues();
      for (var i = 0; i < datosSol.length; i++) {
        var fechaExistente = formatearFecha(datosSol[i][COLUMNAS_SOLICITUDES.FECHA - 1]);
        var horaIniExistente = formatearHora(displayHorasIni[i][0]);
        var horaFinExistente = formatearHora(displayHorasFin[i][0]);

        if (fechaExistente !== fechaTexto) continue;
        if (!horaIniExistente || !horaFinExistente) continue;

        var existIni = horaAMinutos(horaIniExistente);
        var existFin = horaAMinutos(horaFinExistente);
        if (existIni === null || existFin === null) continue;

        if (minSolicitados < existFin && minFinSolicitados > existIni) {
          var profExistente = datosSol[i][COLUMNAS_SOLICITUDES.PROFESOR - 1] || "Otro profesor";
          throw new Error("La sala ya esta reservada el " + fechaTexto + " de " + horaIniExistente + " a " + horaFinExistente + " por " + profExistente + ".");
        }
      }
    }

    // --- Verificar solapamiento en "Planificación diaria" (reservas permanentes) ---
    var sheetHorario = buscarHoja(ss, NOMBRE_HOJA_HORARIO);
    if (sheetHorario) {
      var rangoH = sheetHorario.getDataRange();
      var valoresH = rangoH.getValues();
      var fondosH = rangoH.getBackgrounds();

      var headerRowDays = -1;
      for (var r = 0; r < valoresH.length; r++) {
        var dayCount = 0;
        for (var c = 0; c < valoresH[r].length; c++) {
          var v = valoresH[r][c] ? String(valoresH[r][c]).toLowerCase().trim() : '';
          for (var d = 0; d < DIAS_BUSCAR.length; d++) {
            if (v === DIAS_BUSCAR[d] || v.indexOf(DIAS_BUSCAR[d]) === 0) { dayCount++; break; }
          }
        }
        if (dayCount >= 3) { headerRowDays = r; break; }
      }

      if (headerRowDays !== -1) {
        var colDia = -1;
        for (var c = 2; c < valoresH[headerRowDays].length; c++) {
          var nombreDia = valoresH[headerRowDays][c] ? String(valoresH[headerRowDays][c]).trim().toLowerCase() : '';
          var numDiaCol = -1;
          if (nombreDia.indexOf('lunes') !== -1) numDiaCol = 1;
          else if (nombreDia.indexOf('martes') !== -1) numDiaCol = 2;
          else if (nombreDia.indexOf('miercoles') !== -1 || nombreDia.indexOf('miércoles') !== -1) numDiaCol = 3;
          else if (nombreDia.indexOf('jueves') !== -1) numDiaCol = 4;
          else if (nombreDia.indexOf('viernes') !== -1) numDiaCol = 5;
          if (numDiaCol === numDiaSemana) { colDia = c; break; }
        }

        if (colDia !== -1) {
          for (var r = headerRowDays + 1; r < valoresH.length; r++) {
            var textoHora = valoresH[r][1] ? String(valoresH[r][1]).trim() : '';
            if (textoHora === '' || textoHora.toLowerCase().indexOf('almuerzo') !== -1) continue;
            var partes = textoHora.replace(/\s/g, '').split('-');
            if (partes.length < 2) continue;

            var bloqIni = horaAMinutos(partes[0].trim());
            var bloqFin = horaAMinutos(partes[1].trim());
            if (bloqIni === null || bloqFin === null) continue;

            if (minSolicitados < bloqFin && minFinSolicitados > bloqIni) {
              var celdaFondo = fondosH[r][colDia] ? fondosH[r][colDia].toLowerCase() : '#ffffff';
              var esVerde = celdaFondo === '#b6d7a8' || celdaFondo === '#0f9d58';
              var tieneContenido = valoresH[r][colDia] && String(valoresH[r][colDia]).trim() !== '';
              if (esVerde || tieneContenido) {
                var textoCelda = valoresH[r][colDia] ? String(valoresH[r][colDia]).trim() : '';
                throw new Error("La sala tiene una reserva el " + diasNombres[numDiaSemana] + " de " + partes[0].trim() + " a " + partes[1].trim() + (textoCelda ? " (" + textoCelda + ")" : "") + ".");
              }
            }
          }
        }
      }
    }

    // --- Escribir la reserva ---
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
      profesor,
      curso,
      fechaObj,
      horaInicio,
      horaFin,
      "Pendiente",
      observaciones
    ];

    sheetSolicitudes.getRange(primeraFilaLibre, 1, 1, datosNuevos.length).setValues([datosNuevos]);

    // --- Notificar al admin de la nueva solicitud ---
    try {
      var linkHoja = ss.getUrl();
      var asuntoAdmin = "Nueva reserva de Sala: " + profesor + " - " + fechaTexto;
      var cuerpoAdmin = "Nueva solicitud de reserva de Sala.\n\n"
        + "Profesor: " + profesor + "\n"
        + "Curso: " + curso + "\n"
        + "Fecha: " + fechaTexto + " (" + diasNombres[numDiaSemana] + ")\n"
        + "Horario: " + horaInicio + " a " + horaFin + "\n"
        + "Estado: Pendiente de aprobacion\n\n"
        + "Para aprobar o rechazar, abre la hoja de Solicitudes:\n" + linkHoja;
      MailApp.sendEmail(CONFIG_EMAIL_ADMIN, asuntoAdmin, cuerpoAdmin);
    } catch (errNotif) {
      Logger.log("Error al notificar al admin: " + errNotif.message);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Reserva de Sala registrada para el dia ' + diasNombres[numDiaSemana] + ' (' + fechaTexto + ') de ' + horaInicio + ' a ' + horaFin + '.'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Error: ' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseFechaDisplay(str) {
  if (!str) return null;
  var partes = str.split('/');
  if (partes.length === 3) {
    return new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
  }
  return null;
}

function horaAMinutos(horaStr) {
  if (!horaStr) return null;
  var partes = String(horaStr).replace(/\s/g, '').replace('.', ':').split(':');
  if (partes.length < 2) return null;
  var h = parseInt(partes[0], 10);
  var m = parseInt(partes[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function formatearFecha(valor) {
  if (valor instanceof Date) {
    var d = valor.getDate();
    var m = valor.getMonth() + 1;
    var y = valor.getFullYear();
    if (y > 1900) return String(d).padStart(2, '0') + '/' + String(m).padStart(2, '0') + '/' + y;
    return String(valor);
  }
  return String(valor || "").trim();
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

function buscarCorreoProfesor(nombreBuscado, ss) {
  var sheetProfes = buscarHoja(ss, NOMBRE_HOJA_PROFESORES);
  if (!sheetProfes) return null;

  var datos = sheetProfes.getDataRange().getValues();
  var nombreNorm = nombreBuscado.toString().trim().toLowerCase();

  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][0] || "").trim().toLowerCase();
    var apellido = String(datos[i][1] || "").trim().toLowerCase();
    var correo = String(datos[i][2] || "").trim();

    if (!correo) continue;

    var nombreCompleto = apellido ? nombre + " " + apellido : nombre;
    if (nombreCompleto.indexOf(nombreNorm) !== -1 || nombre.indexOf(nombreNorm) !== -1) {
      return correo;
    }
  }
  return null;
}

// === onEdit: Envio de correo cuando el admin acepta/rechaza ===
function onEditSala(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSolicitudes = buscarHoja(ss, NOMBRE_HOJA_SOLICITUDES);
  if (!sheetSolicitudes || sheet.getName() !== sheetSolicitudes.getName()) return;
  if (col !== COLUMNAS_SOLICITUDES.RESERV_ACEPTADA) return;
  if (row <= 1) return;

  var valor = String(range.getValue() || "").trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (valor !== "si" && valor !== "no") return;

  var celdaEstado = sheet.getRange(row, COLUMNAS_SOLICITUDES.OBSERVACIONES);
  var obsActual = String(celdaEstado.getValue() || "").trim();
  if (obsActual.toLowerCase().indexOf('enviado') !== -1) return;

  var filaData = sheet.getRange(row, 1, 1, 7).getValues()[0];
  var profesor = String(filaData[COLUMNAS_SOLICITUDES.PROFESOR - 1] || "").trim();
  var curso = String(filaData[COLUMNAS_SOLICITUDES.CURSO - 1] || "").trim();
  var fechaDisplay = sheet.getRange(row, COLUMNAS_SOLICITUDES.FECHA).getDisplayValue();
  var horaInicioRaw = sheet.getRange(row, COLUMNAS_SOLICITUDES.HORA_INICIO).getDisplayValue();
  var horaFinRaw = sheet.getRange(row, COLUMNAS_SOLICITUDES.HORA_FIN).getDisplayValue();

  var fecha = formatearFecha(fechaDisplay);
  var horaInicio = formatearHora(horaInicioRaw);
  var horaFin = formatearHora(horaFinRaw);

  var mensajeObs = "";

  if (valor === "si") {
    mensajeObs = reescribirHorarioSala(row);
  } else {
    mensajeObs = "Rechazado";
  }

  var correo = buscarCorreoProfesor(profesor, ss);
  if (!correo) {
    Logger.log("No se encontro correo para: " + profesor);
    mensajeObs += " - No se encontró el docente";
    celdaEstado.setValue(mensajeObs.trim());
    return;
  }

  var asunto, cuerpo;
  if (valor === "si") {
    asunto = "Reserva Sala APROBADA - " + fecha + " (" + horaInicio + " a " + horaFin + ")";
    cuerpo = "Hola " + profesor + ",\n\nTu reserva de la Sala ha sido APROBADA.\n\nDetalles:\n- Curso: " + curso + "\n- Fecha: " + fecha + "\n- Horario: " + horaInicio + " a " + horaFin + (obsActual ? "\n- Observaciones: " + obsActual : "") + "\n\nSaludos,\nEquipo de TI";
  } else {
    asunto = "Reserva Sala RECHAZADA - " + fecha + " (" + horaInicio + " a " + horaFin + ")";
    cuerpo = "Hola " + profesor + ",\n\nLamentamos informarte que tu reserva de la Sala ha sido RECHAZADA.\n\nDetalles:\n- Curso: " + curso + "\n- Fecha: " + fecha + "\n- Horario: " + horaInicio + " a " + horaFin + (obsActual ? "\n- Motivo: " + obsActual : "") + "\n\nSaludos,\nEquipo de TI";
  }

  try {
    MailApp.sendEmail(correo, asunto, cuerpo, { name: 'RESERVA SALA' });
    mensajeObs += " - Enviado ✓";
  } catch (err) {
    Logger.log("Error al enviar correo: " + err.message);
    mensajeObs += " - Error al enviar correo";
  }

  celdaEstado.setValue(mensajeObs.trim());
}

function reescribirHorarioSala(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSol = buscarHoja(ss, NOMBRE_HOJA_SOLICITUDES);
  if (!sheetSol) { Logger.log("SALA: No se encontró Solicitudes"); return "Error: No se encontró Solicitudes"; }
  var sheetHorario = buscarHoja(ss, NOMBRE_HOJA_HORARIO);
  if (!sheetHorario) { Logger.log("SALA: No se encontró " + NOMBRE_HOJA_HORARIO); return "Error: No se encontró la hoja Horario"; }

  var rowData = sheetSol.getRange(rowNum, 1, 1, 7).getValues()[0];
  var curso = String(rowData[COLUMNAS_SOLICITUDES.CURSO - 1] || "").trim();
  var fechaDisplay = sheetSol.getRange(rowNum, COLUMNAS_SOLICITUDES.FECHA).getDisplayValue();
  var horaInicioStr = formatearHora(sheetSol.getRange(rowNum, COLUMNAS_SOLICITUDES.HORA_INICIO).getDisplayValue());
  var horaFinStr = formatearHora(sheetSol.getRange(rowNum, COLUMNAS_SOLICITUDES.HORA_FIN).getDisplayValue());

  Logger.log("SALA: curso=" + curso + " fechaDisplay=" + fechaDisplay + " horaIni=" + horaInicioStr + " horaFin=" + horaFinStr);

  var fechaDate = parseFechaDisplay(fechaDisplay);
  if (!fechaDate) { Logger.log("SALA: Fecha inválida: " + fechaDisplay); return "Error: Fecha inválida"; }
  var numDia = fechaDate.getDay();
  var MAPEO_DIAS = { 1: 3, 2: 4, 3: 5, 4: 6, 5: 7 };
  var colDia = MAPEO_DIAS[numDia];
  if (!colDia) { Logger.log("SALA: Día no válido (fin de semana): " + numDia); return "Error: Día no válido (fin de semana)"; }

  var minInicio = horaAMinutos(horaInicioStr);
  var minFin = horaAMinutos(horaFinStr);
  if (minInicio === null || minFin === null || minInicio >= minFin) { Logger.log("SALA: Horas inválidas: " + horaInicioStr + "-" + horaFinStr); return "Error: Horas inválidas"; }

  Logger.log("SALA: minInicio=" + minInicio + " minFin=" + minFin + " colDia=" + colDia);

  var values = sheetHorario.getDataRange().getValues();
  Logger.log("SALA: Total filas horario=" + values.length);
  var headerRowDays = -1;
  for (var r = 0; r < values.length; r++) {
    var dayCount = 0;
    for (var c = 0; c < values[r].length; c++) {
      var val = values[r][c] ? String(values[r][c]).toLowerCase().trim() : '';
      for (var d = 0; d < DIAS_BUSCAR.length; d++) {
        if (val === DIAS_BUSCAR[d] || val.indexOf(DIAS_BUSCAR[d]) === 0) { dayCount++; break; }
      }
    }
    if (dayCount >= 3) { headerRowDays = r; Logger.log("SALA: headerRowDays=" + r + " dayCount=" + dayCount); break; }
  }
  if (headerRowDays === -1) { Logger.log("SALA: No se encontró fila de días"); return "Error: No se encontró fila de días en el horario"; }

  var profesor = String(rowData[COLUMNAS_SOLICITUDES.PROFESOR - 1] || "").trim();
  var textoReserva = (profesor + " " + curso).toUpperCase();

  for (var r = headerRowDays + 1; r < values.length; r++) {
    var celdaHora = values[r][1] ? String(values[r][1]).trim() : '';
    if (!celdaHora || celdaHora.toLowerCase().indexOf('almuerzo') !== -1) continue;

    var partes = celdaHora.replace(/\s/g, '').replace(/\./g, ':').split('-');
    if (partes.length < 2) continue;

    var minBloIni = horaAMinutos(partes[0]);
    var minBloFin = horaAMinutos(partes[1]);
    if (minBloIni === null || minBloFin === null) continue;

    if (minInicio < minBloFin && minFin > minBloIni) {
      var celda = sheetHorario.getRange(r + 1, colDia);
      var valor = String(celda.getValue() || "").trim();
      if (!valor) {
        celda.setValue(textoReserva);
        celda.setFontFamily("Roboto");
        celda.setFontSize(11);
        celda.setFontWeight("bold");
        celda.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        ss.toast("Reserva escrita en horario", "Éxito", 5);
        return "Reserva escrita en horario";
      }
      return "Error: El horario ya está ocupado";
    }
  }
  return "Error: No se encontró bloque horario que coincida";
}

function obtenerSolicitudesProximaSemana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = buscarHoja(ss, NOMBRE_HOJA_SOLICITUDES);
  if (!sheet) throw new Error("No se encontro la hoja 'Solicitudes'.");

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
    var fecha = (fechaCelda instanceof Date) ? fechaCelda : parseFechaDisplay(String(fechaCelda));
    if (!fecha) continue;
    fecha = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
    if (fecha.getTime() >= lunes.getTime() && fecha.getTime() <= viernes.getTime()) {
      resultados.push({
        dia: DIAS_NOMBRES[fecha.getDay()],
        fecha: String(fecha.getDate()).padStart(2,'0') + " de " + MESES_NOMBRES[fecha.getMonth()],
        horaInicio: formatearHora(datos[i][3]),
        horaFin: formatearHora(datos[i][4]),
        profesor: String(datos[i][0] || "").trim(),
        curso: String(datos[i][1] || "").trim(),
        estado: String(datos[i][5] || "Pendiente").trim()
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

function responderJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

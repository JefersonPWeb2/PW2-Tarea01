// Función para enviar recordatorios por correo electrónico
function enviarRecordatorios() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datos = hoja.getDataRange().getValues();
  
  // Obtener la fila actual y la fecha actual
  var filaActual = 1; // Empezamos en la fila 1 (omitir encabezados)
  var fechaActual = new Date();
  
  // Recorrer cada fila de datos
  for (var i = 1; i < datos.length; i++) {
    var tarea = datos[i][0];
    var responsable = datos[i][1];
    var fechaLimite = new Date(datos[i][2]);
    var horalimite = datos[i][3];
    var estado = datos[i][4];
    var correoResponsable = datos[i][5];
    
    // Verificar tareas pendientes próximas a la fecha límite
    if (estado.toLowerCase() === 'pendiente' && fechaLimite <= fechaActual) {
      // Enviar correo electrónico de recordatorio
      var asunto = 'RECORDATORIO: TAREA PENDIENTE';
      var mensaje = 'Hola ' + responsable + ',\n\nLa tarea "' + tarea + '" está pendiente y vence a las '+ horalimite +'. Por favor, tómese un momento para completarla.\n\nGracias.';
      MailApp.sendEmail(correoResponsable, asunto, mensaje);
      // Actualizar el estado de la tarea en la hoja de cálculo
      hoja.getRange(i + 1, 5).setValue('Sin terminar'); // Columna E (Estado)
    }
  }
}

// FUNCION PARA CALCULAR EL TIEMPO ENTRE HORA COMIENZO Y HORA FINAL
function calcularTiempo() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Consolidado");

  let lastRow = hoja.getLastRow();
  let horasComienzo = hoja.getRange("B2:B" + lastRow).getValues();
  let horasFinal = hoja.getRange("H2:H" + lastRow).getValues();
  let tiempos = [];

  for (let i = 0; i < horasComienzo.length; i++) {
    if (horasComienzo[i][0] && horasFinal[i][0]) {
      let comienzo = new Date(horasComienzo[i][0]);
      let fin = new Date(horasFinal[i][0]);
      let diferencia = (fin - comienzo) / 1000; // Diferencia en segundos

      let horas = Math.floor(diferencia / 3600);
      let minutos = Math.floor((diferencia % 3600) / 60);
      let segundos = Math.floor(diferencia % 60);

      // Formato con ceros a la izquierda
      let horasFormateado = String(horas).padStart(2, '0');
      let minutosFormateado = String(minutos).padStart(2, '0');
      let segundosFormateado = String(segundos).padStart(2, '0');

      tiempos.push([`${horasFormateado}:${minutosFormateado}:${segundosFormateado}`]);
    } else {
      tiempos.push([""]);
    }
  }

  hoja.getRange("I2:I" + lastRow).setValues(tiempos);
}

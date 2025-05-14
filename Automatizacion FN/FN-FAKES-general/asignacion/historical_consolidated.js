function identify_months() {
  let fnFakes = SpreadsheetApp.getActiveSpreadsheet();
  let consolidadoFnFakes = fnFakes.getSheetByName("Consolidado");

  let fechaRange = consolidadoFnFakes.getRange("A:A"); // obtener todos los valores de la columna A
  let fechas = fechaRange.getValues();

  let meses = {}; // objeto para contar los meses únicos

  for (let i = 0; i < fechas.length; i++) {
    let fecha = fechas[i][0]; // obtener la fecha de cada celda
    let date = new Date(fecha); // convertir la fecha a un objeto Date
    let mes = date.getMonth() + 1; // obtener el mes (0-11, así que sumamos 1)
    let mesNombre = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"][mes - 1]; // obtener el nombre del mes

    if (!meses[mesNombre]) {
      meses[mesNombre] = 1; // si el mes no existe, agregarlo con un conteo de 1
    } else {
      meses[mesNombre]++; // si el mes ya existe, incrementar el conteo
    }
  }

  // mostrar los resultados
  let mensaje = "Hay " + Object.keys(meses).length + " meses únicos:\n";
  for (let mes in meses) {
    mensaje += "- " + mes + ": " + meses[mes] + " veces\n";
  }
  Logger.log(mensaje);
  return meses;
}

function consolidadoHistorico() {
  // obtener la hoja de cálculo activa
  let fnFakes = SpreadsheetApp.getActiveSpreadsheet();

  // obtener la hoja de cálculo "Consolidado"
  let consolidadoFnFakes = fnFakes.getSheetByName("Consolidado");

  // obtener el rango de datos en la hoja "Consolidado"
  let dataRange = consolidadoFnFakes.getDataRange();

  // obtener los datos como un array de arrays de valores
  let data = dataRange.getValues();

  // obtener los meses únicos
  let meses = identify_months();

  // obtener los 2 primeros meses más antiguos
  let mesesAntiguos = Object.keys(meses).slice(0, 3);

  // filtrar los datos para solo incluir los registros de los 2 primeros meses más antiguos
  let filteredData = data.filter(function (row) {
    let fecha = row[0];
    let date = new Date(fecha);
    let mes = date.getMonth() + 1;
    let mesNombre = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"][mes - 1];
    return mesesAntiguos.includes(mesNombre);
  });

  // ID del archivo de hojas de cálculo donde se va a guardar el consolidado histórico
  let fileId = "1jvx_9NahDjlCwxdVkaF-j-vEg4exhVqimbyRnr7xbAk";

  // abrir el archivo de hojas de cálculo
  let historicoSpreadsheet = SpreadsheetApp.openById(fileId);

  // obtener la hoja de cálculo activa en el archivo
  let historicoSheet = historicoSpreadsheet.getActiveSheet();

  // agregar los datos al consolidado histórico
  for (let i = 0; i < filteredData.length; i++) {
    let row = filteredData[i];
    historicoSheet.appendRow(row);
  }
}

function eliminarMesesAntiguos() {
  let fnFakes = SpreadsheetApp.getActiveSpreadsheet();
  let consolidadoFnFakes = fnFakes.getSheetByName("Consolidado");

  let meses = identify_months();
  let mesesAntiguos = Object.keys(meses).slice(0, 3);

  let dataRange = consolidadoFnFakes.getDataRange();
  let data = dataRange.getValues();

  let newData = data.filter(function (row) {
    let fecha = row[0];
    let date = new Date(fecha);
    let mes = date.getMonth() + 1;
    let mesNombre = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"][mes - 1];
    return !mesesAntiguos.includes(mesNombre);
  });

  consolidadoFnFakes.clearContents();
  consolidadoFnFakes.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function historical_consolidated() {
  identify_months();
  consolidadoHistorico();
  eliminarMesesAntiguos();
}
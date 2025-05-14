function transferirDatos() {

  // Abre la hoja de cálculo de destino
  var archivoDestino = SpreadsheetApp.openById('1E2WTFRs0fpCSA4FWT56cc-DsHdDcLotaN3ZfJiUP_Ew');
  var hojaDestino = archivoDestino.getSheets()[0]; 

  // Abre la primera hoja de cálculo de origen 
  var archivoOrigen1 = SpreadsheetApp.openById('1TG0_xhcIoXU4syz929vbDjArDIhgEyEjpYjLmw6Krz8');
  var hojaOrigen1 = archivoOrigen1.getSheetByName('Calidad');  
  
  //Abre la segunda hoja de cálculo de origen 
  var archivoOrigen2 = SpreadsheetApp.openById('1PD4RidiaP4UV66cT2UoC43xwM1SDTYzjaDBYxxaJvsE');
  var hojaOrigen2 = archivoOrigen2.getSheetByName('Calidad');  
 
  // Obtén todos los datos de la hoja de origen 1 (comienza desde la fila 2)
  var datos1 = hojaOrigen1.getRange(2, 1, hojaOrigen1.getLastRow() - 1, hojaOrigen1.getLastColumn()).getDisplayValues();
  
  //Obtén todos los datos de la hoja de origen 2 (comienza desde la fila 2)
  var datos2 = hojaOrigen2.getRange(2, 1, hojaOrigen2.getLastRow() - 1, hojaOrigen2.getLastColumn()).getDisplayValues();

  // Unimos los datos de ambas hojas de origen
  var datosCombinados = datos1.concat(datos2);
  
  // Filtramos las filas que contienen "Correcto", "Incorrectos" o "Excluido QA" en la columna T (índice 19)
  var filasParaTransferir = [];
  for (var i = 0; i < datosCombinados.length; i++) {
    var valorColumnaT = datosCombinados[i][19];  // Columna T (índice 19, ya que las filas comienzan en el índice 0)
    
    if (valorColumnaT == "Correcto" || valorColumnaT == "Incorrectos" || valorColumnaT == "Excluido QA") {
      filasParaTransferir.push(datosCombinados[i]);
    }
  }
  
  // Si hay filas para transferir, las agregamos al archivo de destino a partir de la fila 2
  if (filasParaTransferir.length > 0) {
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, filasParaTransferir.length, filasParaTransferir[0].length).setValues(filasParaTransferir);
  } else {
    Logger.log("No se encontraron filas para transferir.");
  }
  console.log(filasParaTransferir.length)
}

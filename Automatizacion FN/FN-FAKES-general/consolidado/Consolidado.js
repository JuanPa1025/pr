function origenDedataCL2() {
  // HOJA ORIGEN 1
  let libro1 = SpreadsheetApp.openById("1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk");
  let hoja1 = libro1.getSheetByName("Calidad");
  let datosHoja1 = [];
  if (hoja1) {
    let fila1 = hoja1.getLastRow();
    let columna1 = Math.min(26, hoja1.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila1 > 1 && columna1 > 0) {
      datosHoja1 = hoja1.getRange(2, 1, fila1 - 1, columna1).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  }

  // HOJA ORIGEN 2
  let libro2 = SpreadsheetApp.openById("16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8");
  let hoja2 = libro2.getSheetByName("Calidad");
  let datosHoja2 = [];
  if (hoja2) {
    let fila2 = hoja2.getLastRow();
    let columna2 = Math.min(26, hoja2.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila2 > 1 && columna2 > 0) {
      datosHoja2 = hoja2.getRange(2, 1, fila2 - 1, columna2).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  } else {
    Logger.log("No se encontró la hoja 2.");
  }

  // Combinamos todos los datos
  let datosCombinados = datosHoja1.concat(datosHoja2);

  Logger.log(`Total de registros combinados: ${datosCombinados.length}`);
  return datosCombinados;
}


// CONFIGURA LOS MENÚS DESPLEGABLES
function configurarColumnas2(hoja) {
  const valoresAnalistaQA = ['Juan Beloso', 'Delfina Moreno', 'Agustina Molina'];
  const valoresAnalisisOk = ['Correcto', 'Incorrectos'];
  const valoresMotivo = ['PDP', 'Error fake/clean', 'Error en atributo', 'Error en precio ', 'Error en descripcion ', 'Lista de precios ', 'No sacó etiqueta ', 'Error en enviar a editar', 'Referencia incorrecta / No puso ref'];

  // Obtener el número de filas con datos (desde la columna A)
  const numFilas = hoja.getLastRow();

  // Definir las reglas de validación de datos
  const reglaAnalistaQA = SpreadsheetApp.newDataValidation().requireValueInList(valoresAnalistaQA).setAllowInvalid(false).build();
  const reglaAnalisisOk = SpreadsheetApp.newDataValidation().requireValueInList(valoresAnalisisOk).setAllowInvalid(false).build();
  const reglaMotivo = SpreadsheetApp.newDataValidation().requireValueInList(valoresMotivo).setAllowInvalid(false).build();

  // Aplicar las validaciones dinámicamente según el número de filas con datos
  hoja.getRange(`T2:T${numFilas}`).setDataValidation(reglaAnalistaQA);
  hoja.getRange(`U2:U${numFilas}`).setDataValidation(reglaAnalisisOk);
  hoja.getRange(`V2:V${numFilas}`).setDataValidation(reglaMotivo);
}

// CONVIERTE UNA FECHA DE "d/m/aaaa" A UN OBJETO DATE
function convertirFecha2(fechaStr) {
  let partes = fechaStr.split("/");
  return new Date(partes[2], partes[1] - 1, partes[0]); // Año, Mes (0 indexado), Día
}

// CREA O ACTUALIZA EL DOCUMENTO POR MES
function creacionDocumento2() {
  let hojaCLData = origenDedataCL2();

  if (hojaCLData.length === 0) {
    console.error("No hay datos para procesar.");
    return;
  }

  let fechaActual = new Date();
  let mesActual = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'MMMM'); // Ejemplo: "Septiembre"
  let anioActual = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy'); // Ejemplo: "2024"

  // Carpeta destino
  let folder = DriveApp.getFolderById("1qU1rCOh50WVXhh4XOhExWWWkN8SH8HVb");

  // Verificar si el archivo ya existe
  let fileName = `${mesActual}-${anioActual}.xlsx`;
  let files = folder.getFilesByName(fileName);
  let librocreado;

  if (files.hasNext()) {
    // Si el archivo ya existe, abrirlo
    let file = files.next();
    librocreado = SpreadsheetApp.open(file);
  } else {
    // Si no existe, crear un nuevo archivo
    librocreado = SpreadsheetApp.create(fileName);

    // Mover el nuevo archivo a la carpeta especificada
    DriveApp.getFileById(librocreado.getId()).moveTo(folder);
    // Define los encabezados en la nueva hoja
    const encabezados = ["x", "Hora comienzo", "ID del caso", "Publicacion", " ", "Analista", "Fecha final", "Hora final", "Tiempo", "Titulo Infracción", "precio ", "Descripción Infracción", "Infracción Imagen", "Tipo de infracción Imagen", "Imagen URL", " ", " ", " ", " ", "Analista QA", "Analisis Ok?", "Motivo", "Comentario", "Devolución Analista", "mes", "semana"];
    let hojaNueva = librocreado.getActiveSheet();
    hojaNueva.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }

  // Filtrar los datos por el mes actual
  // let datosFiltrados = hojaCLData.filter(fila => {
  //   let fechaRegistro = convertirFecha2(fila[0]);  // Convierte el formato "d/m/aaaa" a objeto Date
  //   let mesRegistro = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), 'MMMM');
  //   let anioRegistro = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), 'yyyy');

  //   return mesRegistro === mesActual && anioRegistro === anioActual;
  // });

  if (hojaCLData.length > 0) {
    // Encuentra la primera fila vacía en la hoja activa del archivo
    let hojaActiva = librocreado.getActiveSheet();
    let ultimaFila = hojaActiva.getLastRow();

    // Inserta los nuevos datos bajo los encabezados
    hojaActiva.getRange(ultimaFila + 1, 1, hojaCLData.length, hojaCLData[0].length).setValues(hojaCLData);
  } else {
    console.error("No hay datos para el mes actual.");
  }
}

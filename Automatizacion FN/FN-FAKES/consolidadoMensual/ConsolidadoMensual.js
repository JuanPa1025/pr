// Función que combina los datos de tres hojas diferentes.
function origenDedataCL2() {
  // HOJA ORIGEN 1
  let libro1 = SpreadsheetApp.openById("1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s");
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
  let libro2 = SpreadsheetApp.openById("1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A");
  let hoja2 = libro2.getSheetByName("Calidad");
  let datosHoja2 = [];
  if (hoja2) {
    let fila2 = hoja2.getLastRow();
    let columna2 = Math.min(26, hoja2.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila2 > 1 && columna2 > 0) {
      datosHoja2 = hoja2.getRange(2, 1, fila2 - 1, columna2).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  }

  // HOJA ORIGEN 3
  let libro3 = SpreadsheetApp.openById("1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI");
  let hoja3 = libro3.getSheetByName("Calidad");
  let datosHoja3 = [];
  if (hoja3) {
    let fila3 = hoja3.getLastRow();
    let columna3 = Math.min(26, hoja3.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila3 > 1 && columna3 > 0) {
      datosHoja3 = hoja3.getRange(2, 1, fila3 - 1, columna3).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  }

  // HOJA ORIGEN 4
  let libro4 = SpreadsheetApp.openById("1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0");
  let hoja4 = libro4.getSheetByName("Calidad");
  let datosHoja4 = [];
  if (hoja4) {
    let fila4 = hoja4.getLastRow();
    let columna4 = Math.min(26, hoja4.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila4 > 1 && columna4 > 0) {
      datosHoja4 = hoja4.getRange(2, 1, fila4 - 1, columna4).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  }

  // Combinamos todos los datos
  let datosCombinados = datosHoja1.concat(datosHoja2, datosHoja3, datosHoja4);
  
  return datosCombinados;
}

// CONFIGURA LOS MENÚS DESPLEGABLES
function configurarColumnas2(hoja) {
  //const valoresAnalistaQA = ['Malena Diaz', 'Jeremías Murguia', 'Melisa Martinez'];
  const valoresAnalisisOk = ['Correcto', 'Incorrectos'];
  const valoresMotivo = ['Error fakes /clean', 'Error en precio', 'Error en foto', 'Error en caracteristicas', 'Error en descripción', 'Error en titulo', 'Error en atributo', 'PDP'];
  const numFilas = hoja.getMaxRows();

  //const reglaAnalistaQA = SpreadsheetApp.newDataValidation().requireValueInList(valoresAnalistaQA).setAllowInvalid(false).build();
  const reglaAnalisisOk = SpreadsheetApp.newDataValidation().requireValueInList(valoresAnalisisOk).setAllowInvalid(false).build();
  const reglaMotivo = SpreadsheetApp.newDataValidation().requireValueInList(valoresMotivo).setAllowInvalid(false).build();

  //hoja.getRange(`S2:S${numFilas}`).setDataValidation(reglaAnalistaQA);
  hoja.getRange(`T2:T${numFilas}`).setDataValidation(reglaAnalisisOk);
  hoja.getRange(`U2:U${numFilas}`).setDataValidation(reglaMotivo);
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
  let folder = DriveApp.getFolderById("1m4R8sB82Ckw5bcxVHPTALwnC53Uhvm1l");
  
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
    const encabezados = ["Fecha comienzo", "Hora comienzo", "ID del caso", "Publicacion", "dominio", "Analista", "Fecha final", "Hora final", "Tiempo", "Titulo Infracción", "Descripción Infracción", "Infracción Imagen", "Tipo de infracción Imagen", "Imagen URL", "Tag MS", "Tag ML", "Preguntas y respuestas infracción", "Fecha QA", "Analista QA", "Analisis Ok?", "Motivo", "Comentario", "Comentario/Devolución Analista", "Devolución del analista a la contestación", "Mes", "Semana"];
    let hojaNueva = librocreado.getActiveSheet();
    hojaNueva.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }

  // Filtrar los datos por el mes actual
  let datosFiltrados = hojaCLData.filter(fila => {
    let fechaRegistro = convertirFecha2(fila[0]);  // Convierte el formato "d/m/aaaa" a objeto Date
    let mesRegistro = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), 'MMMM');
    let anioRegistro = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), 'yyyy');

    return mesRegistro === mesActual && anioRegistro === anioActual;
  });

  if (datosFiltrados.length > 0) {
    // Encuentra la primera fila vacía en la hoja activa del archivo
    let hojaActiva = librocreado.getActiveSheet();
    let ultimaFila = hojaActiva.getLastRow();
    
    // Inserta los nuevos datos bajo los encabezados
    hojaActiva.getRange(ultimaFila + 1, 1, datosFiltrados.length, datosFiltrados[0].length).setValues(datosFiltrados);
  } else {
    console.error("No hay datos para el mes actual.");
  }
}

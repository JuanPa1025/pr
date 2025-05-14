// Función que combina los datos de dos hojas diferentes.
function origenDedataCL2() {
  // HOJA ORIGEN 1
  let libro1 = SpreadsheetApp.openById("1TG0_xhcIoXU4syz929vbDjArDIhgEyEjpYjLmw6Krz8");
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
  let libro2 = SpreadsheetApp.openById("1PD4RidiaP4UV66cT2UoC43xwM1SDTYzjaDBYxxaJvsE");
  let hoja2 = libro2.getSheetByName("Calidad");
  let datosHoja2 = [];
  if (hoja2) {
    let fila2 = hoja2.getLastRow();
    let columna2 = Math.min(26, hoja2.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila2 > 1 && columna2 > 0) {
      datosHoja2 = hoja2.getRange(2, 1, fila2 - 1, columna2).getDisplayValues();  // Excluye la primera fila (encabezados)
    }
  }

  // Combinamos los datos de ambos libros
  let datosCombinados = datosHoja1.concat(datosHoja2);
  
  return datosCombinados;
}

// CONFIGURA LOS MENÚS DESPLEGABLES
function configurarColumnas2(hoja) {
  //const valoresAnalistaQA = ['Leila', 'Nahuel'];
  const valoresAnalisisOk = ['Correcto', 'Incorrectos', 'Excluido QA'];
  const valoresMotivo = ['Falta Infracción', 'No aplica Infracción']
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

// BORRAR LOS DATOS DE LAS HOJAS DE ORIGEN
function borrarDatosOrigen() {
  // BORRAR DATOS DE LA HOJA ORIGEN 1
  let libro1 = SpreadsheetApp.openById("1TG0_xhcIoXU4syz929vbDjArDIhgEyEjpYjLmw6Krz8");
  let hoja1 = libro1.getSheetByName("Calidad");
  if (hoja1) {
    let fila1 = hoja1.getLastRow();
    let columna1 = Math.min(26, hoja1.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila1 > 1 && columna1 > 0) {
      hoja1.getRange(2, 1, fila1 - 1, columna1).clearContent();  // Borra el contenido a partir de la fila 2 (sin tocar los encabezados)
    }
  }

  // BORRAR DATOS DE LA HOJA ORIGEN 2
  let libro2 = SpreadsheetApp.openById("1PD4RidiaP4UV66cT2UoC43xwM1SDTYzjaDBYxxaJvsE");
  let hoja2 = libro2.getSheetByName("Calidad");
  if (hoja2) {
    let fila2 = hoja2.getLastRow();
    let columna2 = Math.min(26, hoja2.getLastColumn());  // Limita a 26 columnas (A-Z)
    if (fila2 > 1 && columna2 > 0) {
      hoja2.getRange(2, 1, fila2 - 1, columna2).clearContent();  // Borra el contenido a partir de la fila 2 (sin tocar los encabezados)
    }
  }
}

function creacionDocumento2() {
  let hojaCLData = origenDedataCL2();

  if (hojaCLData.length === 0) {
    console.error("No hay datos para procesar.");
    return;
  }

  let fechaActual = new Date();
  let mesActual = fechaActual.toLocaleDateString("es-ES", { month: "long" });
  mesActual = mesActual.charAt(0).toUpperCase() + mesActual.slice(1);
  let anioActual = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy');

  // Carpeta destino
  let folder = DriveApp.getFolderById("16NubD5puPRLnY8I7QO1b8YWkm_GDVsls");
  let fileName = `${mesActual}-${anioActual}.xlsx`;
  let files = folder.getFilesByName(fileName);
  let librocreado;

  if (files.hasNext()) {
    librocreado = SpreadsheetApp.open(files.next());
  } else {
    librocreado = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(librocreado.getId()).moveTo(folder);

    const encabezados = ["Fecha comienzo", "Hora comienzo", "ID del caso", "Publicacion", "dominio", "Analista", "Fecha final", "Hora final", "Tiempo", "Titulo Infracción", "Descripción Infracción", "Infracción Imagen", "Tipo de infracción Imagen", "Imagen URL", "Tag MS", "Tag ML", "Preguntas y respuestas infracción", "Fecha QA", "Analista QA", "Analisis Ok?", "Motivo", "Comentario", "Comentario/Devolución Analista", "Devolución del analista a la contestación", "Mes", "Semana"];
    let hojaNueva = librocreado.getActiveSheet();
    hojaNueva.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }

  // Filtrar los datos por el mes actual
  let datosFiltrados = hojaCLData.filter(fila => {
    let valorFecha = fila[0]; // Suponiendo que la fecha está en la columna A
    if (!valorFecha || typeof valorFecha !== 'string') return false;

    try {
      let fechaRegistro = convertirFecha2(valorFecha);
      let mesRegistro = fechaRegistro.toLocaleDateString("es-ES", { month: "long" });
      mesRegistro = mesRegistro.charAt(0).toUpperCase() + mesRegistro.slice(1);
      let anioRegistro = Utilities.formatDate(fechaRegistro, Session.getScriptTimeZone(), 'yyyy');

      return mesRegistro === mesActual && anioRegistro === anioActual;
    } catch (e) {
      console.warn("Error al convertir la fecha:", valorFecha);
      return false;
    }
  });

  if (datosFiltrados.length > 0) {
    let hojaActiva = librocreado.getActiveSheet();
    let ultimaFila = hojaActiva.getLastRow();
    hojaActiva.getRange(ultimaFila + 1, 1, datosFiltrados.length, datosFiltrados[0].length).setValues(datosFiltrados);

    borrarDatosOrigen(); // Solo si se insertaron datos
  } else {
    console.error("No hay datos para el mes actual.");
  }
}













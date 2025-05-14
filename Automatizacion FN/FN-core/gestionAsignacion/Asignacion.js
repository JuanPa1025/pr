// FILTRO: FUNCIONA CON FORMATO 8/9/2024 d/m/aaaa

function aplicarFiltros1() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A:Q');

  var filtroExistente = rango.getFilter();
  if (filtroExistente) {
    filtroExistente.remove();
  }

  var filtro = rango.createFilter();

  var hoy = new Date();
  var ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);

  // Convertir la fecha de ayer a formato m/d/yyyy
  var dia = ayer.getDate().toString();
  var mes = (ayer.getMonth() + 1).toString();
  var anio = ayer.getFullYear();
  var fechaAyerTextoDMY = dia + '/' + mes + '/' + anio;

  // Convertir la fecha de ayer a formato dd/mm/yyyy
  // var diaDD = ('0' + ayer.getDate()).slice(-2);
  // var mesDD = ('0' + (ayer.getMonth() + 1)).slice(-2);
  // var fechaAyerTextoDMY = diaDD + '/' + mesDD + '/' + anio;

  var columnaFecha = 1;
  var criterioFechaDMY = SpreadsheetApp.newFilterCriteria()
    .whenTextContains(fechaAyerTextoDMY)
    .build();

  var columnaNoVacio = 4; // Columna D
  var criterioNoVacio = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();

  // Aplicar filtros para ambos formatos de fecha
  filtro.setColumnFilterCriteria(columnaFecha, criterioFechaDMY);
  filtro.setColumnFilterCriteria(columnaNoVacio, criterioNoVacio);

  rango.sort({ column: 3, ascending: true });
}

// PASAR DATOS Analista 1 ------------------------------------------------------------------------------------------------
function pasarDatosAnalista1() {
  // Obtén la hoja activa
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hojaActiva.getRange('A2:O' + hojaActiva.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datosActivos = rango.getDisplayValues(); // Usa getDisplayValues() para preservar formato de hora

  // Obtener los registros de hojaNahuel (solo los de ayer)
  var hojaNahuel = SpreadsheetApp.openById('1PD4RidiaP4UV66cT2UoC43xwM1SDTYzjaDBYxxaJvsE').getSheetByName('Calidad');
  var registrosDestino2 = hojaNahuel.getRange(2, 1, hojaNahuel.getLastRow(), hojaNahuel.getLastColumn()).getDisplayValues(); // También usamos getDisplayValues()

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener registros de una fecha específica
  function obtenerDatosParaFecha(datos, fechaObjetivo) {
    return datos.filter(function(row) {
      var fechaCelda = row[0]; // Suponiendo que la fecha está en la columna 0 (A)

      // Convertir la fecha si es una cadena en formato d/m/aaaa
      var fechaConvertida = (typeof fechaCelda === 'string' && fechaCelda.includes('/')) 
                            ? convertirTextoAFecha(fechaCelda) 
                            : (fechaCelda instanceof Date ? fechaCelda : null);

      if (fechaConvertida) {
        fechaConvertida.setHours(0, 0, 0, 0); // Ajustar la hora para comparar solo la fecha
        return fechaConvertida.getTime() === fechaObjetivo.getTime();
      }
      return false; // Si no se puede convertir la fecha, la descartamos
    });
  }

  // Función para obtener la última fecha registrada (de atrás hacia adelante)
  function obtenerUltimaFechaRegistros() {
    var hoy = new Date();
    var fechaObjetivo = new Date(hoy);
    fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

    // Buscar registros de fechas hacia atrás hasta encontrar registros
    var datosFiltrados;
    while (datosFiltrados === undefined || datosFiltrados.length === 0) {
      fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
      datosFiltrados = obtenerDatosParaFecha(datosActivos, fechaObjetivo);
      if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
        // Detener si la fecha llega a una fecha límite (puedes ajustar esto según sea necesario)
        Logger.log('No se encontraron datos en fechas anteriores.');
        return null;
      }
    }
    return fechaObjetivo; // Devolver la última fecha con datos encontrados
  }

  // Obtener la última fecha registrada
  var ultimaFecha = obtenerUltimaFechaRegistros();
  if (!ultimaFecha) {
    Logger.log('No se encontraron datos para ninguna fecha.');
    return;
  }

  // Filtrar los registros de la última fecha registrada
  var datosActivosUltimaFecha = obtenerDatosParaFecha(datosActivos, ultimaFecha);
  var datosDestino2UltimaFecha = obtenerDatosParaFecha(registrosDestino2, ultimaFecha);

  // Función para verificar si un registro ya existe en los datos de hojaDestino2
  function yaExistente(registro, registrosDestino2) {
    return registrosDestino2.some(function(destino) {
      return destino[2] === registro[2] && destino[3] === registro[3];
    });
  }

  // Filtrar los registros de la hoja activa de la última fecha que no están en hojaDestino2
  var datosUnicos = datosActivosUltimaFecha.filter(row => !yaExistente(row, datosDestino2UltimaFecha));

  // Si no hay datos únicos, no hacemos nada
  if (datosUnicos.length === 0) {
    Logger.log('No hay datos nuevos para copiar. No hay datos para la última fecha registrada o Todos los datos ya existen en la hoja de Leila.');
    return;
  }

  // Función para seleccionar aleatoriamente N registros de una lista
  function seleccionarRegistrosAleatorios(lista, n) {
    var seleccionados = [];
    while (seleccionados.length < n && lista.length > 0) {
      var indiceAleatorio = Math.floor(Math.random() * lista.length);
      seleccionados.push(lista[indiceAleatorio]);
      lista.splice(indiceAleatorio, 1); // Eliminar el registro ya seleccionado
    }
    return seleccionados;
  }

  // Registros aleatorios de ayer
  var datosAleatorios = seleccionarRegistrosAleatorios(datosUnicos, 600);

  // Abrimos la hoja de destino (hojaLeila)
  var hojaLeila = SpreadsheetApp.openById('1TG0_xhcIoXU4syz929vbDjArDIhgEyEjpYjLmw6Krz8').getSheetByName('Calidad');

  // Copiar los datos a hojaLeila
  if (datosAleatorios.length > 0) {
    hojaLeila.getRange(hojaLeila.getLastRow() + 1, 1, datosAleatorios.length, datosAleatorios[0].length).setValues(datosAleatorios);
    Logger.log('Se copiaron ' + datosAleatorios.length + ' datos nuevos aleatorios de ayer a hojaLeila.');
  } else {
    Logger.log('No se copiaron datos porque no hay datos nuevos de ayer para agregar.');
  }
}

// PASAR DATOS Analista 2 ------------------------------------------------------------------------------------------------
function pasarDatosAnalista2() {
  // Obtén la hoja activa
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hojaActiva.getRange('A2:O' + hojaActiva.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datosActivos = rango.getDisplayValues(); // Usa getDisplayValues() para preservar formato de hora

  // Obtener los registros de hojaLeila (solo los de la última fecha registrada)
  var hojaLeila = SpreadsheetApp.openById('1TG0_xhcIoXU4syz929vbDjArDIhgEyEjpYjLmw6Krz8').getSheetByName('Calidad');
  var registrosDestino2 = hojaLeila.getRange(2, 1, hojaLeila.getLastRow(), hojaLeila.getLastColumn()).getDisplayValues(); // También usamos getDisplayValues()

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener registros de una fecha específica
  function obtenerDatosParaFecha(datos, fechaObjetivo) {
    return datos.filter(function(row) {
      var fechaCelda = row[0]; // Suponiendo que la fecha está en la columna 0 (A)

      // Convertir la fecha si es una cadena en formato d/m/aaaa
      var fechaConvertida = (typeof fechaCelda === 'string' && fechaCelda.includes('/')) 
                            ? convertirTextoAFecha(fechaCelda) 
                            : (fechaCelda instanceof Date ? fechaCelda : null);

      if (fechaConvertida) {
        fechaConvertida.setHours(0, 0, 0, 0); // Ajustar la hora para comparar solo la fecha
        return fechaConvertida.getTime() === fechaObjetivo.getTime();
      }
      return false; // Si no se puede convertir la fecha, la descartamos
    });
  }

  // Función para obtener la última fecha registrada (de atrás hacia adelante)
  function obtenerUltimaFechaRegistros() {
    var hoy = new Date();
    var fechaObjetivo = new Date(hoy);
    fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

    // Buscar registros de fechas hacia atrás hasta encontrar registros
    var datosFiltrados;
    while (datosFiltrados === undefined || datosFiltrados.length === 0) {
      fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
      datosFiltrados = obtenerDatosParaFecha(datosActivos, fechaObjetivo);
      if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
        // Detener si la fecha llega a una fecha límite (puedes ajustar esto según sea necesario)
        Logger.log('No se encontraron datos en fechas anteriores.');
        return null;
      }
    }
    return fechaObjetivo; // Devolver la última fecha con datos encontrados
  }

  // Obtener la última fecha registrada
  var ultimaFecha = obtenerUltimaFechaRegistros();
  if (!ultimaFecha) {
    Logger.log('No se encontraron datos para ninguna fecha.');
    return;
  }

  // Filtrar los registros de la última fecha registrada
  var datosActivosUltimaFecha = obtenerDatosParaFecha(datosActivos, ultimaFecha);
  var datosDestino2UltimaFecha = obtenerDatosParaFecha(registrosDestino2, ultimaFecha);

  // Función para verificar si un registro ya existe en los datos de hojaDestino2
  function yaExistente(registro, registrosDestino2) {
    return registrosDestino2.some(function(destino) {
      return destino[2] === registro[2] && destino[3] === registro[3];
    });
  }

  // Filtrar los registros de la hoja activa de la última fecha que no están en hojaDestino2
  var datosUnicos = datosActivosUltimaFecha.filter(row => !yaExistente(row, datosDestino2UltimaFecha));

  // Si no hay datos únicos, no hacemos nada
  if (datosUnicos.length === 0) {
    Logger.log('No hay datos nuevos para copiar. No hay datos para la última fecha registrada o Todos los datos ya existen en la hoja de Leila.');
    return;
  }

  // Función para seleccionar aleatoriamente N registros de una lista
  function seleccionarRegistrosAleatorios(lista, n) {
    var seleccionados = [];
    while (seleccionados.length < n && lista.length > 0) {
      var indiceAleatorio = Math.floor(Math.random() * lista.length);
      seleccionados.push(lista[indiceAleatorio]);
      lista.splice(indiceAleatorio, 1); // Eliminar el registro ya seleccionado
    }
    return seleccionados;
  }

  // Registros aleatorios de la última fecha registrada
  var datosAleatorios = seleccionarRegistrosAleatorios(datosUnicos, 400);

  // Abrimos la hoja de destino (hojaNahuel)
  var hojaNahuel = SpreadsheetApp.openById('1PD4RidiaP4UV66cT2UoC43xwM1SDTYzjaDBYxxaJvsE').getSheetByName('Calidad');

  // Copiar los datos a hojaNahuel
  if (datosAleatorios.length > 0) {
    hojaNahuel.getRange(hojaNahuel.getLastRow() + 1, 1, datosAleatorios.length, datosAleatorios[0].length).setValues(datosAleatorios);
    Logger.log('Se copiaron ' + datosAleatorios.length + ' datos nuevos aleatorios de la última fecha registrada a hojaNahuel.');
  } else {
    Logger.log('No se copiaron datos porque no hay datos nuevos de la última fecha registrada para agregar.');
  }
}

// ELIMINAR FILTROS
function eliminarFiltros1() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define el rango donde se aplicaron los filtros
  var rango = hoja.getRange('A:Q');

  // Verifica si hay un filtro aplicado en el rango
  var filtroExistente = rango.getFilter();

  // Si hay un filtro, quítalo
  if (filtroExistente) {
    filtroExistente.remove();
  }
}

// MENU
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Asignacion QA')
    .addItem('Aplicar Filtros', 'aplicarFiltros1')
    .addItem('Asignación de Datos Leila', 'pasarDatosAnalista1')
    .addItem('Asignación de Datos Nahuel', 'pasarDatosAnalista2')
    .addSeparator()
    .addItem('Eliminar Filtros', 'eliminarFiltros1')
    .addToUi();
}

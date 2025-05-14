// FILTRO: FUNCIONA CON FORMATO 8/9/2024 d/m/aaaa

function aplicarFiltros1() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A:O');

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
// PASAR DATOS ANALISTA 1 
function pasarDatosAnalista1() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:Q' + hoja.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datos = rango.getValues();

  // Obtener registros ya asignados de las hojas de otros analistas
  var hojaAnalista2 = SpreadsheetApp.openById('1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A').getSheetByName('Calidad');
  var hojaAnalista3 = SpreadsheetApp.openById('1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI').getSheetByName('Calidad');
  var hojaAnalista4 = SpreadsheetApp.openById('1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0').getSheetByName('Calidad');

  var registrosAsignadosA2 = hojaAnalista2.getRange(2, 1, hojaAnalista2.getLastRow(), hojaAnalista2.getLastColumn()).getValues();
  var registrosAsignadosA3 = hojaAnalista3.getRange(2, 1, hojaAnalista3.getLastRow(), hojaAnalista3.getLastColumn()).getValues();
  var registrosAsignadosA4 = hojaAnalista4.getRange(2, 1, hojaAnalista4.getLastRow(), hojaAnalista4.getLastColumn()).getValues();
  

  var registrosAsignados = registrosAsignadosA2.concat(registrosAsignadosA3, registrosAsignadosA4);

  // Función para verificar si un registro ya ha sido asignado
  function yaAsignado(registro, registrosAsignados) {
    return registrosAsignados.some(function (asignado) {
      // Compara las columnas C y D
      return asignado[2] === registro[2] && asignado[3] === registro[3];
    });
  }

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener los registros para una fecha específica
  function obtenerDatosParaFecha(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3];
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  // Identifica la fecha de ayer
  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

  // Buscar registros de fechas hacia atrás hasta encontrar registros
  var datosFiltrados;
  while (datosFiltrados === undefined || datosFiltrados.length === 0) {
    fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
      // Detener si la fecha llega a una fecha límite (puedes ajustar esto según sea necesario)
      Logger.log('No se encontraron datos en fechas anteriores.');
      return;
    }
  }

  // Identifica los correos únicos en la columna F
  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F (índice 5)
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) {
        correos[correo] = [];
      }
      // Solo añade registros que no han sido asignados ya
      if (!yaAsignado(row, registrosAsignados)) {
        correos[correo].push(row);
      }
    }
  });

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

  // Distribuye los registros entre los analistas
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];

    // Verificar si el correo tiene 700 registros para asignar 70
    if (registros.length >= 700) {
      // Selecciona 70 registros aleatorios entre todos los disponibles
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, 70);
      if (!bloquesPorAnalista[0]) {
        bloquesPorAnalista[0] = [];
      }
      bloquesPorAnalista[0] = bloquesPorAnalista[0].concat(registrosAleatorios);
      // En tal caso de ser menos de 700 sacar el 10% de los registros y asignarlos y verificar que sea mayor a 0
    } else if (registros.length < 700 && registros.length > 0) {
      let porcentaje = 0.10;
      let totalRegistro = registros.length * porcentaje;
      totalRegistro = parseInt(totalRegistro);
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, totalRegistro)
      if (!bloquesPorAnalista[0]) {
        bloquesPorAnalista[0] = [];
      }
      bloquesPorAnalista[0] = bloquesPorAnalista[0].concat(registrosAleatorios);
    } else {
      Logger.log('El correo ' + correo + ' no tiene registros, no será asignado.');
    }
  }

  // Abre la hoja de destino donde se copiarán los datos
  var hojaAnalista1 = SpreadsheetApp.openById('1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s').getSheetByName('Calidad');

  // Copia los datos filtrados a la hoja de destino
  var bloques = bloquesPorAnalista[0] || []; // Usa el primer bloque para el analista 1
  if (bloques.length > 0) {
    hojaAnalista1.getRange(2, 1, bloques.length, bloques[0].length).setValues(bloques); // Comienza desde la fila 2
    Logger.log('Datos de la fecha ' + Utilities.formatDate(fechaObjetivo, Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' copiados con éxito a Analista 1.');
  } else {
    Logger.log('No se encontraron bloques suficientes para Analista 1.');
  }
}


// PASAR DATOS Analista 2 --------------------------------------------------------------------
function pasarDatosAnalista2() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:Q' + hoja.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datos = rango.getValues();

  // Obtener registros ya asignados de las hojas de otros analistas
  var hojaAnalista1 = SpreadsheetApp.openById('1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s').getSheetByName('Calidad');
  var hojaAnalista3 = SpreadsheetApp.openById('1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI').getSheetByName('Calidad');
  var hojaAnalista4 = SpreadsheetApp.openById('1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0').getSheetByName('Calidad');

  var registrosAsignadosA1 = hojaAnalista1.getRange(2, 1, hojaAnalista1.getLastRow(), hojaAnalista1.getLastColumn()).getValues();
  var registrosAsignadosA3 = hojaAnalista3.getRange(2, 1, hojaAnalista3.getLastRow(), hojaAnalista3.getLastColumn()).getValues();
  var registrosAsignadosA4 = hojaAnalista4.getRange(2, 1, hojaAnalista4.getLastRow(), hojaAnalista4.getLastColumn()).getValues();  

  // Crear un Set para almacenar registros asignados
  var registrosAsignados = new Set();

  // Agregar registros asignados a un Set para mejorar la búsqueda usando columnas C y D
  registrosAsignadosA1.concat(registrosAsignadosA3, registrosAsignadosA4).forEach(function (asignado) {
    // Usar una representación en cadena de las columnas C y D para evitar duplicados
    registrosAsignados.add(asignado[2] + '|' + asignado[3]); // C y D son índices 2 y 3
  });

  // Función para verificar si un registro ya ha sido asignado (solo C y D)
  function yaAsignado(registro) {
    return registrosAsignados.has(registro[2] + '|' + registro[3]); // C y D son índices 2 y 3
  }

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener los registros para una fecha específica
  function obtenerDatosParaFecha(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3];
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  // Identifica la fecha de ayer
  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

  // Buscar registros de fechas hacia atrás hasta encontrar registros
  var datosFiltrados;
  while (datosFiltrados === undefined || datosFiltrados.length === 0) {
    fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
      Logger.log('No se encontraron datos en fechas anteriores.');
      return;
    }
  }

  // Identifica los correos únicos en la columna F
  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F (índice 5)
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) {
        correos[correo] = [];
      }
      // Solo añade registros que no han sido asignados ya
      if (!yaAsignado(row)) {
        correos[correo].push(row);
      }
    }
  });

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

  // Distribuye los registros entre los analistas
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];

    // Verificar si el correo tiene 700 registros para asignar 70
    if (registros.length >= 700) {
      // Selecciona 70 registros aleatorios entre todos los disponibles
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, 70);
      if (!bloquesPorAnalista[1]) {
        bloquesPorAnalista[1] = [];
      }
      bloquesPorAnalista[1] = bloquesPorAnalista[1].concat(registrosAleatorios);
      // En tal caso de ser menos de 700 sacar el 10% de los registros y asignarlos y verificar que sea mayor a 0
    } else if (registros.length < 700 && registros.length > 0) {
      let porcentaje = 0.10;
      let totalRegistro = registros.length * porcentaje;
      totalRegistro = parseInt(totalRegistro);
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, totalRegistro)
      if (!bloquesPorAnalista[1]) {
        bloquesPorAnalista[1] = [];
      }
      bloquesPorAnalista[1] = bloquesPorAnalista[1].concat(registrosAleatorios);
    } else {
      Logger.log('El correo ' + correo + ' no tiene registros, no será asignado.');
    }
  }

  // Abre la hoja de destino donde se copiarán los datos
  var hojaAnalista2 = SpreadsheetApp.openById('1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A').getSheetByName('Calidad');

  // Copia los datos del segundo bloque a la hoja de destino 
  var bloques = bloquesPorAnalista[1] || []; // Usa el bloque seleccionado para el analista 2
  if (bloques.length > 0) {
    hojaAnalista2.getRange(2, 1, bloques.length, bloques[0].length).setValues(bloques); // Comienza desde la fila 2
    Logger.log('Datos de la fecha ' + Utilities.formatDate(fechaObjetivo, Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' copiados con éxito a Analista 2.');
  } else {
    Logger.log('No se encontraron bloques suficientes para Analista 2.');
  }
}


// PASAR DATOS Analista 3 --------------------------------------------------------------------
function pasarDatosAnalista3() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:Q' + hoja.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datos = rango.getValues();

  // Obtener registros ya asignados de las hojas de otros analistas
  var hojaAnalista1 = SpreadsheetApp.openById('1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s').getSheetByName('Calidad');
  var hojaAnalista2 = SpreadsheetApp.openById('1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A').getSheetByName('Calidad');
  var hojaAnalista4 = SpreadsheetApp.openById('1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0').getSheetByName('Calidad');

  var registrosAsignadosA1 = hojaAnalista1.getRange(2, 1, hojaAnalista1.getLastRow(), hojaAnalista1.getLastColumn()).getValues();
  var registrosAsignadosA2 = hojaAnalista2.getRange(2, 1, hojaAnalista2.getLastRow(), hojaAnalista2.getLastColumn()).getValues();
  var registrosAsignadosA4 = hojaAnalista4.getRange(2, 1, hojaAnalista4.getLastRow(), hojaAnalista4.getLastColumn()).getValues();

  // Crear un Set para almacenar registros asignados usando columnas C y D
  var registrosAsignados = new Set();

  // Agregar registros asignados a un Set
  registrosAsignadosA1.concat(registrosAsignadosA2, registrosAsignadosA4).forEach(function (asignado) {
    registrosAsignados.add(asignado[2] + '|' + asignado[3]); // C y D son índices 2 y 3
  });

  // Función para verificar si un registro ya ha sido asignado (solo C y D)
  function yaAsignado(registro) {
    return registrosAsignados.has(registro[2] + '|' + registro[3]); // C y D son índices 2 y 3
  }

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener los registros para una fecha específica
  function obtenerDatosParaFecha(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3];
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  // Identifica la fecha de ayer
  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

  // Buscar registros de fechas hacia atrás hasta encontrar registros
  var datosFiltrados;
  while (datosFiltrados === undefined || datosFiltrados.length === 0) {
    fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
      Logger.log('No se encontraron datos en fechas anteriores.');
      return;
    }
  }

  // Identifica los correos únicos en la columna F
  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F (índice 5)
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) {
        correos[correo] = [];
      }
      // Solo añade registros que no han sido asignados ya
      if (!yaAsignado(row)) {
        correos[correo].push(row);
      }
    }
  });

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

  // Distribuye los registros entre los analistas
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];

    // Verificar si el correo tiene 700 registros para asignar 70
    if (registros.length >= 700) {
      // Selecciona 70 registros aleatorios entre todos los disponibles
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, 70);
      if (!bloquesPorAnalista[2]) {
        bloquesPorAnalista[2] = [];
      }
      bloquesPorAnalista[2] = bloquesPorAnalista[2].concat(registrosAleatorios);
      // En tal caso de ser menos de 700 sacar el 10% de los registros y asignarlos y verificar que sea mayor a 0
    } else if (registros.length < 700 && registros.length > 0) {
      let porcentaje = 0.10;
      let totalRegistro = registros.length * porcentaje;
      totalRegistro = parseInt(totalRegistro);
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, totalRegistro)
      if (!bloquesPorAnalista[2]) {
        bloquesPorAnalista[2] = [];
      }
      bloquesPorAnalista[2] = bloquesPorAnalista[2].concat(registrosAleatorios);
    } else {
      Logger.log('El correo ' + correo + ' no tiene registros, no será asignado.');
    }
  }

  // Abre la hoja de destino donde se copiarán los datos
  var hojaAnalista3 = SpreadsheetApp.openById('1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI').getSheetByName('Calidad');

  // Copia los datos del tercer bloque a la hoja de destino 
  var bloques = bloquesPorAnalista[2] || []; // Usa el bloque seleccionado para el analista 3
  if (bloques.length > 0) {
    hojaAnalista3.getRange(2, 1, bloques.length, bloques[0].length).setValues(bloques); // Comienza desde la fila 2
    Logger.log('Datos de la fecha ' + Utilities.formatDate(fechaObjetivo, Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' copiados con éxito a Analista 3.');
  } else {
    Logger.log('No se encontraron bloques suficientes para Analista 3.');
  }
}

// PASAR DATOS Analista 4 --------------------------------------------------------------------
function pasarDatosAnalista4() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:Q' + hoja.getLastRow()); // Empieza desde la fila 2 para evitar encabezados
  var datos = rango.getValues();

  // Obtener registros ya asignados de las hojas de otros analistas
  var hojaAnalista1 = SpreadsheetApp.openById('1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s').getSheetByName('Calidad');
  var hojaAnalista2 = SpreadsheetApp.openById('1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A').getSheetByName('Calidad');
  var hojaAnalista3 = SpreadsheetApp.openById('1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI').getSheetByName('Calidad');

  var registrosAsignadosA1 = hojaAnalista1.getRange(2, 1, hojaAnalista1.getLastRow(), hojaAnalista1.getLastColumn()).getValues();
  var registrosAsignadosA2 = hojaAnalista2.getRange(2, 1, hojaAnalista2.getLastRow(), hojaAnalista2.getLastColumn()).getValues();
  var registrosAsignadosA3 = hojaAnalista3.getRange(2, 1, hojaAnalista3.getLastRow(), hojaAnalista3.getLastColumn()).getValues();

  // Crear un Set para almacenar registros asignados usando columnas C y D
  var registrosAsignados = new Set();

  // Agregar registros asignados a un Set
  registrosAsignadosA1.concat(registrosAsignadosA2, registrosAsignadosA3).forEach(function (asignado) {
    registrosAsignados.add(asignado[2] + '|' + asignado[3]); // C y D son índices 2 y 3
  });

  // Función para verificar si un registro ya ha sido asignado (solo C y D)
  function yaAsignado(registro) {
    return registrosAsignados.has(registro[2] + '|' + registro[3]); // C y D son índices 2 y 3
  }

  // Función para convertir una fecha en formato texto d/m/aaaa a objeto Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

  // Función para obtener los registros para una fecha específica
  function obtenerDatosParaFecha(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3];
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  // Identifica la fecha de ayer
  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setHours(0, 0, 0, 0); // Establecer las horas a 0 para comparar solo la fecha

  // Buscar registros de fechas hacia atrás hasta encontrar registros
  var datosFiltrados;
  while (datosFiltrados === undefined || datosFiltrados.length === 0) {
    fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (fechaObjetivo.getTime() < new Date('2000-01-01').getTime()) {
      Logger.log('No se encontraron datos en fechas anteriores.');
      return;
    }
  }

  // Identifica los correos únicos en la columna F
  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F (índice 5)
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) {
        correos[correo] = [];
      }
      // Solo añade registros que no han sido asignados ya
      if (!yaAsignado(row)) {
        correos[correo].push(row);
      }
    }
  });

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

  // Distribuye los registros entre los analistas
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];

    // Verificar si el correo tiene 700 registros para asignar 70
    if (registros.length >= 700) {
      // Selecciona 70 registros aleatorios entre todos los disponibles
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, 70);
      if (!bloquesPorAnalista[2]) {
        bloquesPorAnalista[2] = [];
      }
      bloquesPorAnalista[2] = bloquesPorAnalista[2].concat(registrosAleatorios);
      // En tal caso de ser menos de 700 sacar el 10% de los registros y asignarlos y verificar que sea mayor a 0
    } else if (registros.length < 700 && registros.length > 0) {
      let porcentaje = 0.10;
      let totalRegistro = registros.length * porcentaje;
      totalRegistro = parseInt(totalRegistro);
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, totalRegistro)
      if (!bloquesPorAnalista[2]) {
        bloquesPorAnalista[2] = [];
      }
      bloquesPorAnalista[2] = bloquesPorAnalista[2].concat(registrosAleatorios);
    } else {
      Logger.log('El correo ' + correo + ' no tiene registros, no será asignado.');
    }
  }

  // Abre la hoja de destino donde se copiarán los datos
  var hojaAnalista4 = SpreadsheetApp.openById('1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0').getSheetByName('Calidad');

  // Copia los datos del tercer bloque a la hoja de destino 
  var bloques = bloquesPorAnalista[2] || []; // Usa el bloque seleccionado para el analista 3
  if (bloques.length > 0) {
    hojaAnalista4.getRange(2, 1, bloques.length, bloques[0].length).setValues(bloques); // Comienza desde la fila 2
    Logger.log('Datos de la fecha ' + Utilities.formatDate(fechaObjetivo, Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' copiados con éxito a Analista 4.');
  } else {
    Logger.log('No se encontraron bloques suficientes para Analista 4.');
  }
}

// ELIMINAR FILTROS
function eliminarFiltros1() {
  // Obtén la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define el rango donde se aplicaron los filtros
  var rango = hoja.getRange('A:O');

  // Verifica si hay un filtro aplicado en el rango
  var filtroExistente = rango.getFilter();

  // Si hay un filtro, quítalo
  if (filtroExistente) {
    filtroExistente.remove();
  }
}

/// Mostrar el formulario de asignación
function mostrarFormulario() {
    var html = HtmlService.createHtmlOutputFromFile('Formulario')
        .setWidth(400)
        .setHeight(350);
    SpreadsheetApp.getUi().showModalDialog(html, 'Asignar Registros');
}

////////////////////Función para asignar registros a analistas
function asignarRegistros(correoSeleccionado, cantidad, diaSeleccionado, analistaSeleccionado) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var datos = hoja.getRange('A2:O' + hoja.getLastRow()).getValues();

    // Normaliza el correo
    correoSeleccionado = correoSeleccionado.trim().toLowerCase();

    // Obtener la fecha objetivo
    var hoy = new Date();
    var fechaObjetivo = new Date(hoy);
    switch (diaSeleccionado) {
        case 'domingo':
            fechaObjetivo.setDate(hoy.getDate() - 1);
            break;
        case 'sabado':
            fechaObjetivo.setDate(hoy.getDate() - 2);
            break;
        case 'viernes':
            fechaObjetivo.setDate(hoy.getDate() - 3);
            break;
        case 'jueves':
            fechaObjetivo.setDate(hoy.getDate() - 4);
            break;
        case 'miercoles':
            fechaObjetivo.setDate(hoy.getDate() - 5);
            break;
        default:
            SpreadsheetApp.getUi().alert('Día seleccionado no válido.');
            return;
    }
    fechaObjetivo.setHours(0, 0, 0, 0);

    // Filtrar registros por correo y fecha
    var registrosFiltrados = datos.filter(row => {
        var fechaCelda = row[0];
        var correo = row[5].trim().toLowerCase();
        if (correo !== correoSeleccionado) return false;

        // Comparar fechas
        if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
            var partes = fechaCelda.split('/');
            var fechaConvertida = new Date(partes[2], partes[1] - 1, partes[0]);
            fechaConvertida.setHours(0, 0, 0, 0);
            return fechaConvertida.getTime() === fechaObjetivo.getTime();
        }
        if (fechaCelda instanceof Date) {
            fechaCelda.setHours(0, 0, 0, 0);
            return fechaCelda.getTime() === fechaObjetivo.getTime();
        }
        return false;
    });

    Logger.log('Total de registros filtrados: ' + registrosFiltrados.length);
    
    // Definir la hoja del analista 
    var hojaAnalistaOpuestaId = analistaSeleccionado === 'Analista 1'
        ? '1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s' // Hoja del Analista 1 
        : analistaSeleccionado === 'Analista 2'
        ? '1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI' // Hoja del Analista 2
        : analistaSeleccionado === 'Analista 3'
        ? '1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A' // Hoja del Analista 3
        : '1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0'; // Hoja del Analista 4


    var hojaAnalistaOpuesta = SpreadsheetApp.openById(hojaAnalistaOpuestaId).getSheetByName('Calidad');
    var datosAnalistaOpuesta = hojaAnalistaOpuesta.getRange('A2:O' + hojaAnalistaOpuesta.getLastRow()).getValues();
    
    // Crear un conjunto para verificar duplicados
    var idsAsignadosOpuesto = new Set(datosAnalistaOpuesta.map(row => row[2] + '|' + row[3]));

    // Filtrar registros que no estén en el analista opuesto
    registrosFiltrados = registrosFiltrados.filter(row => {
        var idCombinado = row[2] + '|' + row[3];
        return !idsAsignadosOpuesto.has(idCombinado);
    });

    // Log para depuración
    Logger.log('Registros después de filtrar duplicados: ' + registrosFiltrados.length);

    // Verificar si hay suficientes registros disponibles
    if (registrosFiltrados.length < cantidad) {
        Logger.log('No hay suficientes registros disponibles para: ' + correoSeleccionado + ' en ' + diaSeleccionado);
        SpreadsheetApp.getUi().alert('No hay suficientes registros disponibles.');
        return;
    }

    // Seleccionar registros aleatorios
    var seleccionados = [];
    while (seleccionados.length < cantidad) {
        var indiceAleatorio = Math.floor(Math.random() * registrosFiltrados.length);
        seleccionados.push(registrosFiltrados[indiceAleatorio]);
        registrosFiltrados.splice(indiceAleatorio, 1);
    }

    // Determinar la hoja de destino según el analista seleccionado
    var hojaDestinoId = analistaSeleccionado === 'Analista 1'
        ? '1VIg2iMhmdYO5fCIZW_VTINrTfnn8MbAjgWAquYYFU6s' // Hoja del Analista 1 
        : '1WQZMfMmJvAt92Zm2a4-5PbM-sn3ekSwE6XrGd2W4awI' // Hoja del Analista 2
        ; '1ni4c_hKUTsf5zgDeu1T1Qm9KggDeXL6-F7sH2DgMn-A' // Hoja del Analista 3
        ; '1gtA7Jhr88tGj3HqLhiHjx7qa3riIcGBA0L55x_GTsC0' // Hoja del Analista 4

    var hojaDestino = SpreadsheetApp.openById(hojaDestinoId).getSheetByName('Calidad');
    var filaInicio = hojaDestino.getLastRow() + 1;

    // Copiar los registros a la hoja de destino
    hojaDestino.getRange(filaInicio, 1, seleccionados.length, seleccionados[0].length).setValues(seleccionados);

    SpreadsheetApp.getUi().alert('Registros asignados con éxito.');
}

// MENU
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Asignacion QA')
    .addItem('Aplicar Filtros', 'aplicarFiltros1')
    .addItem('Asignación de Datos Malena Diaz', 'pasarDatosAnalista1')
    .addItem('Asignación de Datos Jeremías Murguia  ', 'pasarDatosAnalista2')
    .addItem('Asignación de Datos Melisa Martinez  ', 'pasarDatosAnalista3')
    .addItem('Asignación de Datos Nahuel Peralta  ', 'pasarDatosAnalista4')
    .addSeparator()
    .addItem('Eliminar Filtros', 'eliminarFiltros1')
    .addSeparator()
    .addItem('Abrir Formulario', 'mostrarFormulario')
    .addToUi();
}

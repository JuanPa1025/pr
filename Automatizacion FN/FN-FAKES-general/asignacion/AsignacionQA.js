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

// PASAR DATOS DIA MARTES- VIERNES ---------------------------------------------------------------------------------

function pasarDatosAnalista1() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:O' + hoja.getLastRow()); // Desde la fila 2 para evitar encabezados
  var datos = rango.getValues();

  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1; // Los meses en JS son de 0 a 11
    var anio = parseInt(partes[2], 10);
    return new Date(anio, mes, dia);
  }

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

  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setDate(hoy.getDate() - 1); // Restar un día inicialmente
  fechaObjetivo.setHours(0, 0, 0, 0);

  var datosFiltrados = [];

  // Bucle para retroceder hasta encontrar registros
  while (datosFiltrados.length === 0 && fechaObjetivo >= new Date(hoy.getFullYear(), hoy.getMonth(), 1)) {
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (datosFiltrados.length === 0) {
      fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día si no hay registros
    }
  }

  if (datosFiltrados.length === 0) {
    Logger.log('No se encontraron datos para asignar.');
    return;
  }

  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F (índice 5)
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) {
        correos[correo] = [];
      }
      correos[correo].push(row);
    }
  });

  var idsAsignados = obtenerIdsAsignados(); // Verificar registros ya asignados

  function obtenerIdsAsignados() {
    var hojaAnalista2 = SpreadsheetApp.openById('16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8') // Hoja analista 2
      .getSheetByName('Calidad');
    var datosAnalista2 = hojaAnalista2.getRange('C2:D' + hojaAnalista2.getLastRow()).getValues();
    return new Set(datosAnalista2.map(row => row.join('-'))); // Combina ID de caso y publicación
  }

  function seleccionarRegistrosAleatorios(registros, cantidad) {
    var seleccionados = [];
    var disponibles = registros.filter(row => {
      var idUnico = row[2] + '-' + row[3]; // Combina ID de caso y publicación
      return !idsAsignados.has(idUnico); // Verifica que no esté asignado al Analista 2
    });

    while (seleccionados.length < cantidad && disponibles.length > 0) {
      var indiceAleatorio = Math.floor(Math.random() * disponibles.length);
      seleccionados.push(disponibles.splice(indiceAleatorio, 1)[0]);
    }
    return seleccionados;
  }

  var bloquesPorCorreo = {};
  for (var correo in correos) {
    var registros = correos[correo];

    if (registros.length >= 30) {
      var registrosAleatorios = seleccionarRegistrosAleatorios(registros, 30);
      bloquesPorCorreo[correo] = registrosAleatorios;
    } else {
      Logger.log('El correo ' + correo + ' tiene menos de 30 registros, no será asignado.');
    }
  }

  var hojaAnalista1 = SpreadsheetApp.openById('1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk') // hoja al que se le asignara
    .getSheetByName('Calidad');

  var registrosAsignados = [];
  for (var correo in bloquesPorCorreo) {
    registrosAsignados = registrosAsignados.concat(bloquesPorCorreo[correo]);
  }

  if (registrosAsignados.length > 0) {
    let ultimaFila = hojaAnalista1.getLastRow();
    hojaAnalista1.getRange(ultimaFila + 1, 1, registrosAsignados.length, registrosAsignados[0].length).setValues(registrosAsignados);
    Logger.log('Datos copiados con éxito para el Analista 1.');
  } else {
    Logger.log('No se encontraron registros suficientes para asignar al Analista 1.');
  }
}


// PASAR DATOS LUNES MIXEADO SIN MENU --------------------------------------------------------------------------------
function pasarDatosAnalista1Lunes() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:O' + hoja.getLastRow());
  var datos = rango.getValues();

  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    return new Date(partes[2], partes[1] - 1, partes[0]);
  }

  function obtenerDatosPorDia(fechaObjetivo) {
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

  var hoy = new Date();
  var fechasObjetivo = [];
  for (var i = 1; i <= 3; i++) {
    var fecha = new Date(hoy);
    fecha.setDate(hoy.getDate() - i);
    fecha.setHours(0, 0, 0, 0);
    fechasObjetivo.push(fecha);
  }

  var datosFiltradosPorDia = {};
  var correos = {};

  // Obtener registros para cada fecha
  for (var i = 0; i < fechasObjetivo.length; i++) {
    var fechaObjetivo = fechasObjetivo[i];
    var registrosDia = obtenerDatosPorDia(fechaObjetivo);

    // Agrupar registros por correo
    registrosDia.forEach(row => {
      var correo = row[5]; // Columna F (índice 5)
      if (correo && correo.trim() !== '') {
        if (!correos[correo]) {
          correos[correo] = { viernes: [], sabado: [], domingo: [] };
        }

        // Asignar los registros a los días correspondientes
        if (i === 0) correos[correo].viernes.push(row);
        if (i === 1) correos[correo].sabado.push(row);
        if (i === 2) correos[correo].domingo.push(row);
      }
    });
  }

  // Cargar registros ya asignados al Analista 2
  var hojaAnalista2 = SpreadsheetApp.openById('16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8').getSheetByName('Calidad');
  var datosAnalista2 = hojaAnalista2.getRange('A2:O' + hojaAnalista2.getLastRow()).getValues();
  var idsAsignadosAnalista2 = new Set(datosAnalista2.map(row => row[2] + '|' + row[3])); // Combina ID de caso y publicación

  // Obtener 10 registros del viernes, 10 del sábado, y 10 del domingo (total 30) de forma aleatoria
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];
    var bloques = [];

    // Seleccionar 10 registros aleatorios del viernes, sábado y domingo
    for (var dia in registros) {
      var registrosDia = registros[dia];
      // Filtrar registros no duplicados con el Analista 2
      var registrosFiltrados = registrosDia.filter(row => {
        var idCombinado = row[2] + '|' + row[3]; // ID de caso + publicación
        return !idsAsignadosAnalista2.has(idCombinado);
      });

      // Si hay suficientes registros, seleccionar aleatorios
      if (registrosFiltrados.length > 10) {
        var seleccionados = [];
        while (seleccionados.length < 10) {
          var indiceAleatorio = Math.floor(Math.random() * registrosFiltrados.length);
          seleccionados.push(registrosFiltrados[indiceAleatorio]);
          registrosFiltrados.splice(indiceAleatorio, 1); // Eliminar para evitar duplicados en la selección
        }
        bloques = bloques.concat(seleccionados);
      } else {
        Logger.log('No hay suficientes registros en ' + dia + ' para ' + correo);
      }
    }

    // Solo agregar si se obtuvieron registros
    if (bloques.length > 0) {
      bloquesPorAnalista[correo] = bloques;
    }
  }

  // Copiar los datos a la hoja de destino
  var hojaAnalista1 = SpreadsheetApp.openById('1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk').getSheetByName('Calidad');

  // Comenzar a copiar desde la fila 2
  var filaInicio = 2;
  for (var correo in bloquesPorAnalista) {
    var bloques = bloquesPorAnalista[correo];
    if (bloques.length > 0) {
      hojaAnalista1.getRange(filaInicio, 1, bloques.length, bloques[0].length).setValues(bloques);
      filaInicio += bloques.length; // Actualizar fila de inicio
    }
  }
  Logger.log('Datos de las fechas viernes, sábado y domingo copiados con éxito a Analista 1.');
}


// PASAR DATOS Analista 2 MARTES-VIERNES ------------------------------------------------------------------------------------
function pasarDatosAnalista2() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:O' + hoja.getLastRow()); // Empieza desde la fila 2
  var datos = rango.getValues();

  // Cargar los registros ya asignados al Analista 1
  var hojaAnalista1 = SpreadsheetApp.openById('1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk').getSheetByName('Calidad'); 
  var datosAnalista1 = hojaAnalista1.getRange('A2:O' + hojaAnalista1.getLastRow()).getValues();
  var idsAsignadosAnalista1 = new Set(datosAnalista1.map(row => row[2] + '|' + row[3])); // Combina ID de caso y publicación

  // Función para convertir fecha de texto a Date
  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    return new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
  }

  // Filtrar registros para una fecha específica
  function obtenerDatosParaFecha(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3]; // Verifica que la celda no esté vacía
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  // Obtener la fecha de ayer
  var hoy = new Date();
  var fechaObjetivo = new Date(hoy);
  fechaObjetivo.setDate(hoy.getDate() - 1);
  fechaObjetivo.setHours(0, 0, 0, 0);

  // Buscar registros para la fecha objetivo
  var datosFiltrados = [];

  // Bucle para retroceder hasta encontrar registros
  while (datosFiltrados.length === 0 && fechaObjetivo >= new Date(hoy.getFullYear(), hoy.getMonth(), 1)) {
    datosFiltrados = obtenerDatosParaFecha(fechaObjetivo);
    if (datosFiltrados.length === 0) {
      fechaObjetivo.setDate(fechaObjetivo.getDate() - 1); // Restar un día si no hay registros
    }
  }

  if (datosFiltrados.length === 0) {
    Logger.log('No se encontraron datos para asignar.');
    return;
  }

  // Agrupar registros por correo (columna F)
  var correos = {};
  datosFiltrados.forEach(row => {
    var correo = row[5]; // Columna F
    if (correo && correo.trim() !== '') {
      if (!correos[correo]) correos[correo] = [];
      correos[correo].push(row);
    }
  });

  // Función para seleccionar 30 registros aleatorios que no estén asignados al Analista 1
  function seleccionarRegistrosAleatorios(registros) {
    var registrosFiltrados = registros.filter(row => {
      var idCombinado = row[2] + '|' + row[3]; // ID de caso + publicación
      return !idsAsignadosAnalista1.has(idCombinado);
    });

    if (registrosFiltrados.length < 30) {
      Logger.log('No hay suficientes registros no duplicados para este correo.');
      return [];
    }

    // Seleccionar 30 registros aleatorios
    var seleccionados = [];
    while (seleccionados.length < 30) {
      var indiceAleatorio = Math.floor(Math.random() * registrosFiltrados.length);
      seleccionados.push(registrosFiltrados.splice(indiceAleatorio, 1)[0]);
    }
    return seleccionados;
  }

  // Preparar los registros para copiar a la hoja del Analista 2
  var registrosParaCopiar = [];
  for (var correo in correos) {
    var registros = correos[correo];
    var seleccionados = seleccionarRegistrosAleatorios(registros);
    if (seleccionados.length > 0) registrosParaCopiar = registrosParaCopiar.concat(seleccionados);
  }

  if (registrosParaCopiar.length === 0) {
    Logger.log('No se encontraron registros suficientes para Analista 2.');
    return;
  }

  // La hoja de destino del Analista 2
  var hojaAnalista2 = SpreadsheetApp.openById('16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8').getSheetByName('Calidad');
  hojaAnalista2.getRange(2, 1, registrosParaCopiar.length, registrosParaCopiar[0].length).setValues(registrosParaCopiar);

  Logger.log('Datos copiados con éxito a Analista 2.');
}

// PASAR DATOS ANALISTA 2 LUNES MIXEADO SIN MENU -----------------------------------------------------------------------------

function pasarDatosAnalista2Lunes() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getRange('A2:O' + hoja.getLastRow());
  var datos = rango.getValues();

  function convertirTextoAFecha(fechaTexto) {
    var partes = fechaTexto.split('/');
    return new Date(partes[2], partes[1] - 1, partes[0]);
  }

  function obtenerDatosPorDia(fechaObjetivo) {
    return datos.filter(row => {
      var fechaCelda = row[0];
      if (typeof fechaCelda === 'string' && fechaCelda.includes('/')) {
        var fechaConvertida = convertirTextoAFecha(fechaCelda);
        fechaConvertida.setHours(0, 0, 0, 0);
        return fechaConvertida.getTime() === fechaObjetivo.getTime() && row[3]; // Filtramos con la columna D (índice 3)
      }
      if (fechaCelda instanceof Date) {
        fechaCelda.setHours(0, 0, 0, 0);
        return fechaCelda.getTime() === fechaObjetivo.getTime() && row[3];
      }
      return false;
    });
  }

  var hoy = new Date();
  var fechasObjetivo = [];
  for (var i = 1; i <= 3; i++) {
    var fecha = new Date(hoy);
    fecha.setDate(hoy.getDate() - i);
    fecha.setHours(0, 0, 0, 0);
    fechasObjetivo.push(fecha);
  }

  var datosFiltradosPorDia = {};
  var correos = {};

  // Obtener registros para cada fecha
  for (var i = 0; i < fechasObjetivo.length; i++) {
    var fechaObjetivo = fechasObjetivo[i];
    var registrosDia = obtenerDatosPorDia(fechaObjetivo);
    
    // Agrupar registros por correo
    registrosDia.forEach(row => {
      var correo = row[5]; // Columna F (índice 5)
      if (correo && correo.trim() !== '') {
        if (!correos[correo]) {
          correos[correo] = { viernes: [], sabado: [], domingo: [] };
        }

        // Asignar los registros a los días correspondientes
        if (i === 0) correos[correo].viernes.push(row);
        if (i === 1) correos[correo].sabado.push(row);
        if (i === 2) correos[correo].domingo.push(row);
      }
    });
  }

  // Cargar registros ya asignados al Analista 1
  var hojaAnalista1 = SpreadsheetApp.openById('1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk').getSheetByName('Calidad');
  var datosAnalista1 = hojaAnalista1.getRange('A2:O' + hojaAnalista1.getLastRow()).getValues();
  var idsAsignadosAnalista1 = new Set(datosAnalista1.map(row => row[2] + '|' + row[3])); // Combina ID de caso y publicación

  // Obtener 10 registros del viernes, 10 del sábado, y 10 del domingo (total 30) de forma aleatoria
  var bloquesPorAnalista = {};
  for (var correo in correos) {
    var registros = correos[correo];
    var bloques = [];

    // Seleccionar 10 registros aleatorios del viernes, sábado y domingo
    for (var dia in registros) {
      var registrosDia = registros[dia];
      // Filtrar registros no duplicados con el Analista 1
      var registrosFiltrados = registrosDia.filter(row => {
        var idCombinado = row[2] + '|' + row[3]; // ID de caso + publicación
        return !idsAsignadosAnalista1.has(idCombinado);
      });

      // Si hay suficientes registros, seleccionar aleatorios
      if (registrosFiltrados.length > 10) {
        var seleccionados = [];
        while (seleccionados.length < 10) {
          var indiceAleatorio = Math.floor(Math.random() * registrosFiltrados.length);
          seleccionados.push(registrosFiltrados[indiceAleatorio]);
          registrosFiltrados.splice(indiceAleatorio, 1); // Eliminar para evitar duplicados en la selección
        }
        bloques = bloques.concat(seleccionados);
      } else {
        Logger.log('No hay suficientes registros en ' + dia + ' para ' + correo);
      }
    }

    // Solo agregar si se obtuvieron registros
    if (bloques.length > 0) {
      bloquesPorAnalista[correo] = bloques;
    }
  }

  // Copiar los datos a la hoja de destino
  var hojaAnalista2 = SpreadsheetApp.openById('16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8').getSheetByName('Calidad');

  // Comenzar a copiar desde la fila 2
  var filaInicio = 2;

  // Copia los datos filtrados a la hoja de destino
  for (var correo in bloquesPorAnalista) {
    var bloques = bloquesPorAnalista[correo] || []; // Usa el segundo bloque para el analista 2
    if (bloques.length > 0) {
      hojaAnalista2.getRange(filaInicio, 1, bloques.length, bloques[0].length).setValues(bloques); // Comienza desde la fila 2
      filaInicio += bloques.length; // Actualizar fila de inicio
      Logger.log('Datos copiados con éxito a Analista 2.');
    } else {
      Logger.log('No se encontraron bloques suficientes para Analista 2.');
    }
  }
}

// PASAR DATOS ANALISTA 1 y 2 LUNES MIXEADO CON MENU -----------------------------------------------------------------------
/// Mostrar el formulario de asignación
function mostrarFormulario() {
    var html = HtmlService.createHtmlOutputFromFile('Formulario')
        .setWidth(400)
        .setHeight(350);
    SpreadsheetApp.getUi().showModalDialog(html, 'Asignar Registros');
}

// Función para asignar registros a analistas
function asignarRegistros(correoSeleccionado, cantidad, diaSeleccionado, analistaSeleccionado) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var datos = hoja.getRange('A2:O' + hoja.getLastRow()).getDisplayValues();

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
    
    // Definir la hoja del analista opuesto
    var hojaAnalistaOpuestaId = analistaSeleccionado === 'Analista 1'
        ? '16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8' // Hoja del Analista 2
        : '1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk'; // Hoja del Analista 1

    var hojaAnalistaOpuesta = SpreadsheetApp.openById(hojaAnalistaOpuestaId).getSheetByName('Calidad');
    var datosAnalistaOpuesta = hojaAnalistaOpuesta.getRange('A2:O' + hojaAnalistaOpuesta.getLastRow()).getDisplayValues();
    
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
        ? '1AjWdcB21CbsTj7NaONRgJwTZE0pcYuU9Q6k0DSHX1vk' // Hoja del Analista 1
        : '16Xs7BHw90oK4qYMJIiwfeftviAef7HhUHi18ac5fbE8'; // Hoja del Analista 2

    var hojaDestino = SpreadsheetApp.openById(hojaDestinoId).getSheetByName('Calidad');
    var filaInicio = hojaDestino.getLastRow() + 1;

    // Copiar los registros a la hoja de destino
    hojaDestino.getRange(filaInicio, 1, seleccionados.length, seleccionados[0].length).setValues(seleccionados);

    SpreadsheetApp.getUi().alert('Registros asignados con éxito.');
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

// MENU
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Asignacion QA')
    .addItem('Aplicar Filtros', 'aplicarFiltros1')
    .addItem('Eliminar Filtros', 'eliminarFiltros1')
    .addSeparator()
    .addItem('Asignación de Datos Juan Beloso', 'pasarDatosAnalista1')
    .addItem('Asignación de Datos Delfina Moreno  ', 'pasarDatosAnalista2')
    .addSeparator()
    .addItem('Asignación de Datos Lunes Juan Beloso', 'pasarDatosAnalista1Lunes')
    .addItem('Asignación de Datos Lunes Delfina Moreno  ', 'pasarDatosAnalista2Lunes')
    .addSeparator()
    .addItem('Abrir Formulario', 'mostrarFormulario')
    .addToUi();
}

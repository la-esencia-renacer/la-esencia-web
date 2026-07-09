/**
 * Configura una planilla de Google Sheets para seguimiento operativo
 * de contactos recibidos por el formulario de La Esencia.
 *
 * Uso:
 * 1. Abrir la planilla de respuestas del Google Form.
 * 2. Ir a Extensiones -> Apps Script.
 * 3. Pegar este archivo.
 * 4. Guardar y ejecutar configurarSeguimientoLaEsencia().
 * 5. Autorizar permisos si Google lo solicita.
 */
function configurarSeguimientoLaEsencia() {
  const planilla = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRespuestas = obtenerHojaRespuestas_(planilla);

  const hojaSeguimiento = obtenerOCrearHoja_(planilla, "Seguimiento operativo");
  const hojaMensajes = obtenerOCrearHoja_(planilla, "Mensajes modelo");
  const hojaListas = obtenerOCrearHoja_(planilla, "Listas");

  configurarListas_(hojaListas);
  configurarHojaRespuestas_(hojaRespuestas, hojaListas);
  configurarMensajesModelo_(hojaMensajes);
  configurarSeguimientoOperativo_(hojaSeguimiento);

  SpreadsheetApp.flush();
}

const HOJAS_INTERNAS_LA_ESENCIA = [
  "Seguimiento operativo",
  "Mensajes modelo",
  "Listas"
];

const COLUMNAS_INTERNAS_LA_ESENCIA = [
  "Estado",
  "Tipo de contacto",
  "Prioridad",
  "Región",
  "Área validada",
  "Riesgo o sensibilidad",
  "Respuesta enviada",
  "Fecha de respuesta",
  "Próximo paso",
  "Responsable interno",
  "Observaciones"
];

const LISTAS_LA_ESENCIA = {
  "Estados": [
    "Nuevo",
    "Revisado",
    "Responder",
    "Respondido",
    "Agendar reunión",
    "En seguimiento",
    "Integrar a nodo futuro",
    "Derivar a Instituto / Observatorio",
    "Descartar / sin continuidad"
  ],
  "Tipos de contacto": [
    "Ciudadano/a",
    "Profesional",
    "Referente comunitario",
    "Institución",
    "Universidad",
    "Municipio",
    "Empresa / sponsor",
    "Actor territorial",
    "Otro"
  ],
  "Prioridad": [
    "Alta",
    "Media",
    "Baja",
    "Sensible"
  ],
  "Riesgo o sensibilidad": [
    "Bajo",
    "Medio",
    "Alto",
    "Sensible"
  ],
  "Respuesta enviada": [
    "Sí",
    "No"
  ],
  "Región": [
    "AMBA",
    "Buenos Aires Interior",
    "San Luis / Cuyo",
    "Centro",
    "Cuyo",
    "NEA",
    "NOA",
    "Patagonia",
    "Nacional",
    "Exterior",
    "Sin definir"
  ],
  "Área validada": [
    "Comunidad y prevención social",
    "Seguridad humana",
    "Educación",
    "Tecnología pública",
    "Cultura, juego y encuentro comunitario",
    "Justicia accesible",
    "Gestión de crisis",
    "Desarrollo productivo",
    "Instituto / Observatorio",
    "Nodos territoriales",
    "Institucional",
    "Sin definir",
    "Otra"
  ],
  "Próximo paso": [
    "Enviar respuesta general",
    "Enviar respuesta profesional/técnica",
    "Enviar respuesta institucional",
    "Agendar reunión",
    "Sumar a seguimiento",
    "Derivar a Instituto / Observatorio",
    "Evaluar nodo territorial futuro",
    "Esperar nueva información",
    "Sin continuidad"
  ],
  "Responsable interno": [
    "Equipo La Esencia",
    "Instituto / Observatorio",
    "Nodo territorial futuro",
    "Gestión institucional",
    "Pendiente de asignación"
  ]
};

function obtenerHojaRespuestas_(planilla) {
  const activa = planilla.getActiveSheet();
  if (activa && !esHojaInterna_(activa.getName())) {
    return activa;
  }

  const candidatas = planilla
    .getSheets()
    .filter((hoja) => !esHojaInterna_(hoja.getName()));

  const posibleHojaDeFormulario = candidatas.find((hoja) => {
    const nombre = hoja.getName().toLowerCase();
    return nombre.indexOf("respuesta") !== -1 || nombre.indexOf("form") !== -1;
  });

  return posibleHojaDeFormulario || candidatas[0] || activa;
}

function esHojaInterna_(nombre) {
  return HOJAS_INTERNAS_LA_ESENCIA.indexOf(nombre) !== -1;
}

function obtenerOCrearHoja_(planilla, nombre) {
  return planilla.getSheetByName(nombre) || planilla.insertSheet(nombre);
}

function configurarHojaRespuestas_(hoja, hojaListas) {
  if (!hoja) {
    throw new Error("No se encontró una hoja de respuestas para configurar.");
  }

  asegurarColumnasInternas_(hoja);
  hoja.setFrozenRows(1);
  activarFiltro_(hoja);
  aplicarValidaciones_(hoja, hojaListas);
  aplicarFormatoCondicional_(hoja);
  hoja.autoResizeColumns(1, hoja.getLastColumn());
}

function asegurarColumnasInternas_(hoja) {
  const ultimaColumna = Math.max(hoja.getLastColumn(), 1);
  const encabezados = hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0]
    .map((valor) => String(valor).trim());

  const faltantes = COLUMNAS_INTERNAS_LA_ESENCIA.filter((columna) => encabezados.indexOf(columna) === -1);
  if (!faltantes.length) {
    return;
  }

  const inicio = hoja.getLastColumn() + 1;
  hoja.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
  hoja.getRange(1, inicio, 1, faltantes.length).setFontWeight("bold").setBackground("#e7eef5");
}

function activarFiltro_(hoja) {
  const filas = Math.max(hoja.getLastRow(), 1);
  const columnas = Math.max(hoja.getLastColumn(), 1);
  const filtro = hoja.getFilter();

  if (filtro) {
    filtro.remove();
  }

  hoja.getRange(1, 1, filas, columnas).createFilter();
}

function configurarListas_(hojaListas) {
  hojaListas.clear();

  const nombresListas = Object.keys(LISTAS_LA_ESENCIA);
  nombresListas.forEach((nombre, indice) => {
    const columna = indice + 1;
    const valores = LISTAS_LA_ESENCIA[nombre];

    hojaListas.getRange(1, columna).setValue(nombre);
    hojaListas.getRange(2, columna, valores.length, 1).setValues(valores.map((valor) => [valor]));
  });

  hojaListas.setFrozenRows(1);
  hojaListas.getRange(1, 1, 1, nombresListas.length).setFontWeight("bold").setBackground("#e7eef5");
  hojaListas.autoResizeColumns(1, nombresListas.length);
}

function aplicarValidaciones_(hojaRespuestas, hojaListas) {
  const mapaColumnas = obtenerMapaColumnas_(hojaRespuestas);
  const filasValidacion = Math.max(hojaRespuestas.getMaxRows() - 1, 1);

  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Estado"], hojaListas, "Estados", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Tipo de contacto"], hojaListas, "Tipos de contacto", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Prioridad"], hojaListas, "Prioridad", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Región"], hojaListas, "Región", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Área validada"], hojaListas, "Área validada", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Riesgo o sensibilidad"], hojaListas, "Riesgo o sensibilidad", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Respuesta enviada"], hojaListas, "Respuesta enviada", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Próximo paso"], hojaListas, "Próximo paso", filasValidacion);
  aplicarValidacionDesdeLista_(hojaRespuestas, mapaColumnas["Responsable interno"], hojaListas, "Responsable interno", filasValidacion);
}

function aplicarValidacionDesdeLista_(hojaRespuestas, columnaDestino, hojaListas, nombreLista, filasValidacion) {
  if (!columnaDestino) {
    return;
  }

  const rangoLista = obtenerRangoLista_(hojaListas, nombreLista);
  if (!rangoLista) {
    return;
  }

  const regla = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(rangoLista, true)
    .setAllowInvalid(false)
    .build();

  hojaRespuestas.getRange(2, columnaDestino, filasValidacion, 1).setDataValidation(regla);
}

function obtenerRangoLista_(hojaListas, nombreLista) {
  const encabezados = hojaListas.getRange(1, 1, 1, hojaListas.getLastColumn()).getValues()[0];
  const indice = encabezados.indexOf(nombreLista);

  if (indice === -1) {
    return null;
  }

  const columna = indice + 1;
  const valores = LISTAS_LA_ESENCIA[nombreLista];
  return hojaListas.getRange(2, columna, valores.length, 1);
}

function aplicarFormatoCondicional_(hoja) {
  const mapaColumnas = obtenerMapaColumnas_(hoja);
  const filas = Math.max(hoja.getMaxRows() - 1, 1);
  const reglasExistentes = hoja.getConditionalFormatRules();
  const reglasNuevas = [];

  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Estado"], filas, "Nuevo", "#fff3cd");
  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Estado"], filas, "Respondido", "#d9ead3");
  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Estado"], filas, "Agendar reunión", "#d9eaf7");
  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Estado"], filas, "Descartar / sin continuidad", "#eeeeee");
  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Prioridad"], filas, "Alta", "#f4cccc");
  agregarReglaTexto_(reglasNuevas, hoja, mapaColumnas["Prioridad"], filas, "Sensible", "#fce5cd");

  hoja.setConditionalFormatRules(reglasExistentes.concat(reglasNuevas));
}

function agregarReglaTexto_(reglas, hoja, columna, filas, texto, color) {
  if (!columna) {
    return;
  }

  const rango = hoja.getRange(2, columna, filas, 1);
  const regla = SpreadsheetApp
    .newConditionalFormatRule()
    .whenTextEqualTo(texto)
    .setBackground(color)
    .setRanges([rango])
    .build();

  reglas.push(regla);
}

function configurarMensajesModelo_(hoja) {
  const mensajes = [
    [
      "A. Respuesta general ciudadana",
      "Hola, muchas gracias por acercarte a La Esencia.\n\n" +
        "Recibimos tu intención de contacto y participación inicial. Esta etapa es fundacional, cultural, social e institucional; no constituye afiliación partidaria formal ni inscripción legal a una organización política.\n\n" +
        "Estamos ordenando los aportes por territorio, áreas de interés y posibles líneas de colaboración. En función de la información recibida, podremos contactarte para futuras conversaciones, nodos territoriales, espacios de pensamiento o líneas vinculadas al Instituto / Observatorio en desarrollo.\n\n" +
        "Gracias por ser parte de esta primera etapa de construcción.\n\n" +
        "La Esencia\n" +
        "Movimiento Fundacional de Renacer Argentino\n" +
        "laesenciarenacer.com.ar"
    ],
    [
      "B. Respuesta para perfil profesional o técnico",
      "Hola, muchas gracias por acercarte a La Esencia.\n\n" +
        "Valoramos especialmente los aportes profesionales y técnicos en esta etapa fundacional. Estamos ordenando áreas de interés vinculadas a comunidad, educación, tecnología pública, seguridad humana, justicia accesible, cultura, desarrollo productivo, gestión de crisis e Instituto / Observatorio.\n\n" +
        "Tu contacto quedará registrado para futuras mesas de conversación, relevamiento de propuestas o posibles espacios de trabajo técnico.\n\n" +
        "Esta instancia no constituye afiliación partidaria formal ni compromiso institucional automático. Es una primera vía de contacto y construcción progresiva.\n\n" +
        "Muchas gracias.\n\n" +
        "La Esencia\n" +
        "Movimiento Fundacional de Renacer Argentino\n" +
        "laesenciarenacer.com.ar"
    ],
    [
      "C. Respuesta para institución u organización",
      "Estimado/a,\n\n" +
        "Muchas gracias por acercarse a La Esencia.\n\n" +
        "La Esencia se encuentra en etapa fundacional de construcción cultural, social e institucional. Estamos desarrollando una línea técnico-institucional orientada a articular territorio, evidencia, pilotos y cooperación ética para el bien común.\n\n" +
        "Nos interesa conocer mejor la posible línea de articulación y evaluar, de manera gradual y transparente, si existe una agenda compartida.\n\n" +
        "Si lo considera oportuno, podemos coordinar una reunión breve de presentación e intercambio.\n\n" +
        "La Esencia\n" +
        "Movimiento Fundacional de Renacer Argentino\n" +
        "laesenciarenacer.com.ar"
    ]
  ];

  hoja.clear();
  hoja.getRange(1, 1, 1, 2).setValues([["Modelo", "Mensaje"]]);
  hoja.getRange(2, 1, mensajes.length, 2).setValues(mensajes);
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#e7eef5");
  hoja.getRange(2, 2, mensajes.length, 1).setWrap(true);
  hoja.setColumnWidth(1, 280);
  hoja.setColumnWidth(2, 720);
}

function configurarSeguimientoOperativo_(hoja) {
  const encabezados = [
    "Fecha",
    "Contacto",
    "Tipo",
    "Provincia",
    "Área",
    "Estado",
    "Próximo paso",
    "Responsable",
    "Observaciones"
  ];

  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, encabezados.length).setFontWeight("bold").setBackground("#e7eef5");
  hoja.autoResizeColumns(1, encabezados.length);
}

function obtenerMapaColumnas_(hoja) {
  const ultimaColumna = Math.max(hoja.getLastColumn(), 1);
  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];

  return encabezados.reduce((mapa, encabezado, indice) => {
    const nombre = String(encabezado).trim();
    if (nombre) {
      mapa[nombre] = indice + 1;
    }
    return mapa;
  }, {});
}

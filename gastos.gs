// ============================================================
// GASTOS PERSONALES — Google Apps Script
// Tracker de gastos personales con categorías chilenas
// ============================================================

// Paleta de colores principal
var COLORES = {
  azul:          '#1e40af',
  azulClaro:     '#dbeafe',
  crema:         '#fef9ec',
  blanco:        '#ffffff',
  grisClaro:     '#f1f5f9',
  verde:         '#16a34a',
  rojo:          '#dc2626',
  rojoClaro:     '#fee2e2',
  verdeClaro:    '#dcfce7',
  grisTexto:     '#6b7280',
  grisFila:      '#f8fafc'
};

// Categorías de gastos chilenas
var CATEGORIAS = [
  'Supermercado',
  'Restaurantes y delivery',
  'Transporte',
  'Servicios básicos',
  'Salud',
  'Entretenimiento',
  'Suscripciones',
  'Vestuario',
  'Hogar',
  'Educación',
  'Transferencias',
  'Otros'
];

// ============================================================
// FUNCIÓN PRINCIPAL — punto de entrada del script
// ============================================================
function setupGastos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear las hojas en orden correcto
  var hojaTransacciones = crearHojaTransacciones(ss);
  var hojaDashboard     = crearHojaDashboard(ss);

  // Eliminar "Hoja 1" si existe (hoja por defecto de Google Sheets)
  eliminarHojaDefault(ss);

  // Forzar actualización de fórmulas
  SpreadsheetApp.flush();

  // Navegar al Dashboard al finalizar
  ss.setActiveSheet(hojaDashboard);

  SpreadsheetApp.getUi().alert(
    '✅ Configuración completada',
    'Las hojas "💳 Transacciones" y "📊 Dashboard" han sido creadas exitosamente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
// HOJA DE TRANSACCIONES
// ============================================================
function crearHojaTransacciones(ss) {
  var nombreHoja = '💳 Transacciones';

  // Eliminar si ya existe para recrear limpia
  var hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente) {
    ss.deleteSheet(hojaExistente);
  }

  // Insertar nueva hoja al principio
  var hoja = ss.insertSheet(nombreHoja, 0);

  // --- Anchos de columna ---
  hoja.setColumnWidth(1, 35);   // A: #
  hoja.setColumnWidth(2, 100);  // B: FECHA
  hoja.setColumnWidth(3, 290);  // C: DESCRIPCIÓN
  hoja.setColumnWidth(4, 140);  // D: CATEGORÍA
  hoja.setColumnWidth(5, 100);  // E: TIPO
  hoja.setColumnWidth(6, 110);  // F: MONTO
  hoja.setColumnWidth(7, 120);  // G: NOTAS

  // --- Fila 1: Título principal ---
  configurarTituloTransacciones(hoja);

  // --- Fila 2: Encabezados ---
  configurarEncabezadosTransacciones(hoja);

  // --- Filas 3-502: Datos (500 filas) ---
  configurarFilasDatos(hoja);

  // --- Filas 504-506: Totales ---
  configurarTotalesTransacciones(hoja);

  // --- Congelar fila 2 ---
  hoja.setFrozenRows(2);

  return hoja;
}

// Título principal de la hoja de transacciones
function configurarTituloTransacciones(hoja) {
  var rango = hoja.getRange('A1:G1');
  rango.merge();
  rango.setValue('💳 GASTOS PERSONALES — Registro de Transacciones');
  rango.setBackground(COLORES.azul);
  rango.setFontColor(COLORES.blanco);
  rango.setFontSize(14);
  rango.setFontWeight('bold');
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  hoja.setRowHeight(1, 45);
}

// Encabezados de columnas (fila 2)
function configurarEncabezadosTransacciones(hoja) {
  var encabezados = ['#', 'FECHA', 'DESCRIPCIÓN', 'CATEGORÍA', 'TIPO', 'MONTO ($)', 'NOTAS'];
  var rango = hoja.getRange('A2:G2');

  rango.setValues([encabezados]);
  rango.setBackground(COLORES.azulClaro);
  rango.setFontColor(COLORES.azul);
  rango.setFontSize(10);
  rango.setFontWeight('bold');
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  hoja.setRowHeight(2, 30);

  // Borde inferior del encabezado
  rango.setBorder(null, null, true, null, null, null, COLORES.azul, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Filas de datos: numeración, formatos, validaciones y formato condicional
function configurarFilasDatos(hoja) {
  var numFilas   = 500;
  var filaInicio = 3;
  var filaFin    = 502;

  // --- Numeración columna A (1-500) ---
  var numerosA = [];
  for (var i = 1; i <= numFilas; i++) {
    numerosA.push([i]);
  }
  var rangoNumeros = hoja.getRange(filaInicio, 1, numFilas, 1);
  rangoNumeros.setValues(numerosA);
  rangoNumeros.setHorizontalAlignment('center');
  rangoNumeros.setFontColor(COLORES.grisTexto);
  rangoNumeros.setFontSize(9);

  // --- Formato de fecha columna B ---
  var rangoFecha = hoja.getRange(filaInicio, 2, numFilas, 1);
  rangoFecha.setNumberFormat('dd/mm/yyyy');
  rangoFecha.setHorizontalAlignment('center');

  // --- Formato de descripción columna C ---
  var rangoDesc = hoja.getRange(filaInicio, 3, numFilas, 1);
  rangoDesc.setHorizontalAlignment('left');
  rangoDesc.setFontSize(10);

  // --- Validación de datos columna D: Categorías ---
  var reglaCategorias = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIAS, true)
    .setAllowInvalid(false)
    .build();
  var rangoCategorias = hoja.getRange(filaInicio, 4, numFilas, 1);
  rangoCategorias.setDataValidation(reglaCategorias);
  rangoCategorias.setHorizontalAlignment('center');

  // --- Validación de datos columna E: Tipo (Gasto/Ingreso) ---
  var reglaTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Gasto', 'Ingreso'], true)
    .setAllowInvalid(false)
    .build();
  var rangoTipo = hoja.getRange(filaInicio, 5, numFilas, 1);
  rangoTipo.setDataValidation(reglaTipo);
  rangoTipo.setHorizontalAlignment('center');

  // --- Formato numérico columna F: Monto ---
  var rangoMonto = hoja.getRange(filaInicio, 6, numFilas, 1);
  rangoMonto.setNumberFormat('#,##0');
  rangoMonto.setHorizontalAlignment('right');
  rangoMonto.setFontSize(10);

  // --- Formato columna G: Notas ---
  var rangoNotas = hoja.getRange(filaInicio, 7, numFilas, 1);
  rangoNotas.setFontSize(9);
  rangoNotas.setFontColor(COLORES.grisTexto);
  rangoNotas.setHorizontalAlignment('left');

  // --- Altura de filas de datos ---
  for (var f = filaInicio; f <= filaFin; f++) {
    hoja.setRowHeight(f, 22);
  }

  // --- Filas alternadas: blanco y gris claro ---
  for (var j = filaInicio; j <= filaFin; j++) {
    var colorFondo = (j % 2 === 0) ? COLORES.grisFila : COLORES.blanco;
    hoja.getRange(j, 1, 1, 7).setBackground(colorFondo);
  }

  // --- Formato condicional columna E: Gasto=rojo, Ingreso=verde ---
  var reglaGasto   = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Gasto')
    .setBackground(COLORES.rojoClaro)
    .setFontColor(COLORES.rojo)
    .setRanges([hoja.getRange(filaInicio, 5, numFilas, 1)])
    .build();

  var reglaIngreso = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Ingreso')
    .setBackground(COLORES.verdeClaro)
    .setFontColor(COLORES.verde)
    .setRanges([hoja.getRange(filaInicio, 5, numFilas, 1)])
    .build();

  var reglasActuales = hoja.getConditionalFormatRules();
  reglasActuales.push(reglaGasto);
  reglasActuales.push(reglaIngreso);
  hoja.setConditionalFormatRules(reglasActuales);
}

// Filas de totales (504-506)
function configurarTotalesTransacciones(hoja) {
  // --- Separador visual fila 503 ---
  hoja.getRange('A503:G503').setBackground(COLORES.azulClaro);
  hoja.setRowHeight(503, 8);

  // --- Fila 504: Total Gastos ---
  var rangoLabelGasto = hoja.getRange('A504:E504');
  rangoLabelGasto.merge();
  rangoLabelGasto.setValue('TOTAL GASTOS');
  rangoLabelGasto.setBackground(COLORES.rojoClaro);
  rangoLabelGasto.setFontColor(COLORES.rojo);
  rangoLabelGasto.setFontWeight('bold');
  rangoLabelGasto.setFontSize(11);
  rangoLabelGasto.setHorizontalAlignment('right');
  rangoLabelGasto.setVerticalAlignment('middle');

  var celdaMontoGasto = hoja.getRange('F504');
  celdaMontoGasto.setFormula('=SUMIF(E3:E502,"Gasto",F3:F502)');
  celdaMontoGasto.setNumberFormat('#,##0');
  celdaMontoGasto.setBackground(COLORES.rojoClaro);
  celdaMontoGasto.setFontColor(COLORES.rojo);
  celdaMontoGasto.setFontWeight('bold');
  celdaMontoGasto.setFontSize(11);
  celdaMontoGasto.setHorizontalAlignment('right');

  hoja.getRange('G504').setBackground(COLORES.rojoClaro);
  hoja.setRowHeight(504, 28);

  // --- Fila 505: Total Ingresos ---
  var rangoLabelIngreso = hoja.getRange('A505:E505');
  rangoLabelIngreso.merge();
  rangoLabelIngreso.setValue('TOTAL INGRESOS');
  rangoLabelIngreso.setBackground(COLORES.verdeClaro);
  rangoLabelIngreso.setFontColor(COLORES.verde);
  rangoLabelIngreso.setFontWeight('bold');
  rangoLabelIngreso.setFontSize(11);
  rangoLabelIngreso.setHorizontalAlignment('right');
  rangoLabelIngreso.setVerticalAlignment('middle');

  var celdaMontoIngreso = hoja.getRange('F505');
  celdaMontoIngreso.setFormula('=SUMIF(E3:E502,"Ingreso",F3:F502)');
  celdaMontoIngreso.setNumberFormat('#,##0');
  celdaMontoIngreso.setBackground(COLORES.verdeClaro);
  celdaMontoIngreso.setFontColor(COLORES.verde);
  celdaMontoIngreso.setFontWeight('bold');
  celdaMontoIngreso.setFontSize(11);
  celdaMontoIngreso.setHorizontalAlignment('right');

  hoja.getRange('G505').setBackground(COLORES.verdeClaro);
  hoja.setRowHeight(505, 28);

  // --- Fila 506: Balance ---
  var rangoLabelBalance = hoja.getRange('A506:E506');
  rangoLabelBalance.merge();
  rangoLabelBalance.setValue('BALANCE');
  rangoLabelBalance.setBackground(COLORES.azul);
  rangoLabelBalance.setFontColor(COLORES.blanco);
  rangoLabelBalance.setFontWeight('bold');
  rangoLabelBalance.setFontSize(12);
  rangoLabelBalance.setHorizontalAlignment('right');
  rangoLabelBalance.setVerticalAlignment('middle');

  var celdaBalance = hoja.getRange('F506');
  celdaBalance.setFormula('=F505-F504');
  celdaBalance.setNumberFormat('#,##0');
  celdaBalance.setBackground(COLORES.azul);
  celdaBalance.setFontColor(COLORES.blanco);
  celdaBalance.setFontWeight('bold');
  celdaBalance.setFontSize(12);
  celdaBalance.setHorizontalAlignment('right');

  hoja.getRange('G506').setBackground(COLORES.azul);
  hoja.setRowHeight(506, 32);
}

// ============================================================
// HOJA DASHBOARD
// ============================================================
function crearHojaDashboard(ss) {
  var nombreHoja = '📊 Dashboard';

  // Eliminar si ya existe para recrear limpia
  var hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente) {
    ss.deleteSheet(hojaExistente);
  }

  // Insertar después de Transacciones (posición 1)
  var hoja = ss.insertSheet(nombreHoja, 1);

  // --- Anchos de columna ---
  hoja.setColumnWidth(1, 35);   // A: #
  hoja.setColumnWidth(2, 200);  // B: Categoría
  hoja.setColumnWidth(3, 130);  // C: Total
  hoja.setColumnWidth(4, 130);  // D: Porcentaje
  hoja.setColumnWidth(5, 120);  // E: Visual

  // --- Fila 1: Título principal ---
  configurarTituloDashboard(hoja);

  // --- Filas 2-3: Tarjetas de resumen ---
  configurarTarjetasResumen(hoja);

  // --- Fila 4: Espaciado ---
  hoja.setRowHeight(4, 15);
  hoja.getRange('A4:E4').setBackground(COLORES.blanco);

  // --- Fila 5: Encabezado de sección categorías ---
  configurarEncabezadoCategorias(hoja);

  // --- Fila 6: Encabezados de tabla ---
  configurarEncabezadosTabla(hoja);

  // --- Filas 7-18: Datos por categoría ---
  configurarFilasCategorias(hoja);

  // --- Fila 19: Fila total ---
  configurarFilaTotalDashboard(hoja);

  // --- Congelar fila 1 ---
  hoja.setFrozenRows(1);

  return hoja;
}

// Título principal del Dashboard
function configurarTituloDashboard(hoja) {
  var rango = hoja.getRange('A1:E1');
  rango.merge();
  rango.setValue('📊 DASHBOARD — Resumen de Gastos');
  rango.setBackground(COLORES.azul);
  rango.setFontColor(COLORES.blanco);
  rango.setFontSize(14);
  rango.setFontWeight('bold');
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  hoja.setRowHeight(1, 45);
}

// Tarjetas de resumen (filas 2-3)
function configurarTarjetasResumen(hoja) {
  // --- Fila 2: Encabezados de tarjetas ---
  hoja.setRowHeight(2, 32);

  // Tarjeta Total Gastos (A2:B2)
  var rangoTG = hoja.getRange('A2:B2');
  rangoTG.merge();
  rangoTG.setValue('Total Gastos');
  rangoTG.setBackground(COLORES.rojo);
  rangoTG.setFontColor(COLORES.blanco);
  rangoTG.setFontSize(11);
  rangoTG.setFontWeight('bold');
  rangoTG.setHorizontalAlignment('center');
  rangoTG.setVerticalAlignment('middle');

  // Tarjeta Total Ingresos (C2:D2)
  var rangoTI = hoja.getRange('C2:D2');
  rangoTI.merge();
  rangoTI.setValue('Total Ingresos');
  rangoTI.setBackground(COLORES.verde);
  rangoTI.setFontColor(COLORES.blanco);
  rangoTI.setFontSize(11);
  rangoTI.setFontWeight('bold');
  rangoTI.setHorizontalAlignment('center');
  rangoTI.setVerticalAlignment('middle');

  // Tarjeta Balance (E2)
  var rangoB = hoja.getRange('E2');
  rangoB.setValue('Balance');
  rangoB.setBackground(COLORES.azul);
  rangoB.setFontColor(COLORES.blanco);
  rangoB.setFontSize(11);
  rangoB.setFontWeight('bold');
  rangoB.setHorizontalAlignment('center');
  rangoB.setVerticalAlignment('middle');

  // --- Fila 3: Valores de las tarjetas ---
  hoja.setRowHeight(3, 50);

  // Valor Total Gastos (A3:B3)
  var rangoVTG = hoja.getRange('A3:B3');
  rangoVTG.merge();
  rangoVTG.setFormula("=SUMIF('💳 Transacciones'!E3:E502,\"Gasto\",'💳 Transacciones'!F3:F502)");
  rangoVTG.setNumberFormat('#,##0 "CLP"');
  rangoVTG.setBackground(COLORES.rojoClaro);
  rangoVTG.setFontColor(COLORES.rojo);
  rangoVTG.setFontSize(16);
  rangoVTG.setFontWeight('bold');
  rangoVTG.setHorizontalAlignment('center');
  rangoVTG.setVerticalAlignment('middle');

  // Valor Total Ingresos (C3:D3)
  var rangoVTI = hoja.getRange('C3:D3');
  rangoVTI.merge();
  rangoVTI.setFormula("=SUMIF('💳 Transacciones'!E3:E502,\"Ingreso\",'💳 Transacciones'!F3:F502)");
  rangoVTI.setNumberFormat('#,##0 "CLP"');
  rangoVTI.setBackground(COLORES.verdeClaro);
  rangoVTI.setFontColor(COLORES.verde);
  rangoVTI.setFontSize(16);
  rangoVTI.setFontWeight('bold');
  rangoVTI.setHorizontalAlignment('center');
  rangoVTI.setVerticalAlignment('middle');

  // Valor Balance (E3)
  var rangoVB = hoja.getRange('E3');
  rangoVB.setFormula('=C3-A3');
  rangoVB.setNumberFormat('#,##0 "CLP"');
  rangoVB.setBackground(COLORES.azulClaro);
  rangoVB.setFontColor(COLORES.azul);
  rangoVB.setFontSize(16);
  rangoVB.setFontWeight('bold');
  rangoVB.setHorizontalAlignment('center');
  rangoVB.setVerticalAlignment('middle');
}

// Encabezado de sección "GASTOS POR CATEGORÍA" (fila 5)
function configurarEncabezadoCategorias(hoja) {
  var rango = hoja.getRange('A5:E5');
  rango.merge();
  rango.setValue('GASTOS POR CATEGORÍA');
  rango.setBackground(COLORES.azul);
  rango.setFontColor(COLORES.blanco);
  rango.setFontSize(12);
  rango.setFontWeight('bold');
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  hoja.setRowHeight(5, 32);
}

// Encabezados de la tabla de categorías (fila 6)
function configurarEncabezadosTabla(hoja) {
  var encabezados = ['#', 'Categoría', 'Total ($)', '% del Total', 'Visual'];
  var rango = hoja.getRange('A6:E6');
  rango.setValues([encabezados]);
  rango.setBackground(COLORES.azulClaro);
  rango.setFontColor(COLORES.azul);
  rango.setFontSize(10);
  rango.setFontWeight('bold');
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  hoja.setRowHeight(6, 28);
  rango.setBorder(null, null, true, null, null, null, COLORES.azul, SpreadsheetApp.BorderStyle.SOLID);
}

// Filas de datos por categoría (filas 7-18)
function configurarFilasCategorias(hoja) {
  for (var i = 0; i < CATEGORIAS.length; i++) {
    var numFila   = 7 + i;
    var numCateg  = i + 1;
    var categoria = CATEGORIAS[i];

    // Color de fondo alternado
    var colorFondo = (i % 2 === 0) ? COLORES.blanco : COLORES.grisClaro;

    hoja.setRowHeight(numFila, 24);

    // Columna A: número
    var celdaNum = hoja.getRange(numFila, 1);
    celdaNum.setValue(numCateg);
    celdaNum.setBackground(colorFondo);
    celdaNum.setFontColor(COLORES.grisTexto);
    celdaNum.setFontSize(9);
    celdaNum.setHorizontalAlignment('center');
    celdaNum.setVerticalAlignment('middle');

    // Columna B: nombre de categoría
    var celdaCateg = hoja.getRange(numFila, 2);
    celdaCateg.setValue(categoria);
    celdaCateg.setBackground(colorFondo);
    celdaCateg.setFontColor('#1f2937');
    celdaCateg.setFontSize(10);
    celdaCateg.setHorizontalAlignment('left');
    celdaCateg.setVerticalAlignment('middle');

    // Columna C: fórmula SUMIFS con criterios de categoría Y tipo=Gasto
    var celdaTotal = hoja.getRange(numFila, 3);
    var formulaSumifs = "=SUMIFS('💳 Transacciones'!F3:F502,'💳 Transacciones'!D3:D502,\"" + categoria + "\",'💳 Transacciones'!E3:E502,\"Gasto\")";
    celdaTotal.setFormula(formulaSumifs);
    celdaTotal.setNumberFormat('#,##0');
    celdaTotal.setBackground(colorFondo);
    celdaTotal.setFontColor('#1f2937');
    celdaTotal.setFontSize(10);
    celdaTotal.setHorizontalAlignment('right');
    celdaTotal.setVerticalAlignment('middle');

    // Columna D: porcentaje del total de gastos (con manejo de división por cero)
    var celdaPct = hoja.getRange(numFila, 4);
    var letraFilaRef = numFila; // número de fila actual en el dashboard
    var formulaPct  = '=IFERROR(C' + letraFilaRef + '/A3,0)';
    celdaPct.setFormula(formulaPct);
    celdaPct.setNumberFormat('0.0%');
    celdaPct.setBackground(colorFondo);
    celdaPct.setFontColor('#1f2937');
    celdaPct.setFontSize(10);
    celdaPct.setHorizontalAlignment('center');
    celdaPct.setVerticalAlignment('middle');

    // Columna E: barra visual con bloques █ (máximo 30 caracteres = 100%)
    var celdaVisual = hoja.getRange(numFila, 5);
    var formulaVisual = '=IFERROR(REPT("█",ROUND(D' + letraFilaRef + '*30,0)),"")';
    celdaVisual.setFormula(formulaVisual);
    celdaVisual.setBackground(colorFondo);
    celdaVisual.setFontColor(COLORES.azul);
    celdaVisual.setFontSize(9);
    celdaVisual.setHorizontalAlignment('left');
    celdaVisual.setVerticalAlignment('middle');
  }
}

// Fila 19: Total de todas las categorías
function configurarFilaTotalDashboard(hoja) {
  hoja.setRowHeight(19, 28);

  // Columna A y B: etiqueta "TOTAL"
  var rangoLabel = hoja.getRange('A19:B19');
  rangoLabel.merge();
  rangoLabel.setValue('TOTAL');
  rangoLabel.setBackground(COLORES.azulClaro);
  rangoLabel.setFontColor(COLORES.azul);
  rangoLabel.setFontWeight('bold');
  rangoLabel.setFontSize(11);
  rangoLabel.setHorizontalAlignment('right');
  rangoLabel.setVerticalAlignment('middle');

  // Columna C: suma de C7:C18
  var celdaTotalMonto = hoja.getRange('C19');
  celdaTotalMonto.setFormula('=SUM(C7:C18)');
  celdaTotalMonto.setNumberFormat('#,##0');
  celdaTotalMonto.setBackground(COLORES.azulClaro);
  celdaTotalMonto.setFontColor(COLORES.azul);
  celdaTotalMonto.setFontWeight('bold');
  celdaTotalMonto.setFontSize(11);
  celdaTotalMonto.setHorizontalAlignment('right');
  celdaTotalMonto.setVerticalAlignment('middle');

  // Columna D: 100%
  var celdaTotalPct = hoja.getRange('D19');
  celdaTotalPct.setValue('100%');
  celdaTotalPct.setBackground(COLORES.azulClaro);
  celdaTotalPct.setFontColor(COLORES.azul);
  celdaTotalPct.setFontWeight('bold');
  celdaTotalPct.setFontSize(11);
  celdaTotalPct.setHorizontalAlignment('center');
  celdaTotalPct.setVerticalAlignment('middle');

  // Columna E: vacía con mismo fondo
  hoja.getRange('E19').setBackground(COLORES.azulClaro);

  // Borde superior para separar del resto
  hoja.getRange('A19:E19').setBorder(
    true, null, null, null, null, null,
    COLORES.azul, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}

// ============================================================
// ELIMINAR HOJA POR DEFECTO
// ============================================================
function eliminarHojaDefault(ss) {
  // Intentar eliminar hoja con nombres comunes por defecto
  var nombresDefault = ['Hoja 1', 'Hoja1', 'Sheet1', 'Sheet 1'];

  nombresDefault.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (hoja) {
      // Solo eliminar si hay al menos 2 hojas (Google Sheets requiere mínimo 1)
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(hoja);
      }
    }
  });
}

// ============================================================
// GASTOS PERSONALES — Google Apps Script
// Crea las hojas "💳 Transacciones" y "📊 Dashboard"
// Función principal: setupGastos()
// ============================================================

// Paleta de colores
var COLORES = {
  azul:         '#1e40af',
  azulClaro:    '#dbeafe',
  crema:        '#fef9ec',
  blanco:       '#ffffff',
  grisClaro:    '#f1f5f9',
  verde:        '#16a34a',
  rojo:         '#dc2626',
  rojoClaro:    '#fee2e2',
  verdeClaro:   '#dcfce7',
  grisTexto:    '#6b7280',
  grisFilas:    '#f8fafc',
  amarilloCard: '#fef08a',
};

// Categorías chilenas
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
  'Otros',
];

// Colores de fondo para cada categoría en el dashboard
var COLORES_CATEGORIA = [
  '#eff6ff', // Supermercado
  '#fef3c7', // Restaurantes
  '#f0fdf4', // Transporte
  '#fdf4ff', // Servicios
  '#fff1f2', // Salud
  '#f0f9ff', // Entretenimiento
  '#fafaf9', // Suscripciones
  '#fdf2f8', // Vestuario
  '#f7fee7', // Hogar
  '#fff7ed', // Educación
  '#f0fdfa', // Transferencias
  '#f9fafb', // Otros
];

// ============================================================
// FUNCIÓN PRINCIPAL
// ============================================================
function setupGastos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('es_CL');

  // Crear o limpiar hojas
  var sheetTrans = obtenerOCrearHoja(ss, '💳 Transacciones');
  var sheetDash  = obtenerOCrearHoja(ss, '📊 Dashboard');

  // Eliminar "Hoja 1" si existe
  eliminarHojaDefault(ss);

  // Construir cada hoja
  construirTransacciones(sheetTrans);
  construirDashboard(sheetDash);

  // Mover el orden: Transacciones primero, Dashboard segundo
  ss.setActiveSheet(sheetTrans);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(sheetDash);
  ss.moveActiveSheet(2);

  // Dejar el dashboard activo al terminar
  ss.setActiveSheet(sheetDash);

  SpreadsheetApp.getUi().alert('✅ Hoja "Gastos" configurada correctamente.\n\nPuedes comenzar a registrar tus transacciones en la hoja 💳 Transacciones.');
}

// ============================================================
// HOJA TRANSACCIONES
// ============================================================
function construirTransacciones(sheet) {
  sheet.clear();
  sheet.clearConditionalFormatRules();

  // ── Anchos de columna ──
  sheet.setColumnWidth(1, 35);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 290);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 120);

  // ── Altura de filas especiales ──
  sheet.setRowHeight(1, 50);
  sheet.setRowHeight(2, 32);

  // ── FILA 1: Título ──
  var tituloRange = sheet.getRange('A1:G1');
  tituloRange.merge();
  tituloRange.setValue('💳 GASTOS PERSONALES — Registro de Transacciones');
  tituloRange.setBackground(COLORES.azul);
  tituloRange.setFontColor(COLORES.blanco);
  tituloRange.setFontSize(16);
  tituloRange.setFontWeight('bold');
  tituloRange.setHorizontalAlignment('center');
  tituloRange.setVerticalAlignment('middle');

  // ── FILA 2: Cabeceras ──
  var headers = ['#', 'FECHA', 'DESCRIPCIÓN', 'CATEGORÍA', 'TIPO', 'MONTO ($)', 'NOTAS'];
  var headerRange = sheet.getRange('A2:G2');
  headerRange.setValues([headers]);
  headerRange.setBackground(COLORES.azulClaro);
  headerRange.setFontColor(COLORES.azul);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, true, true, COLORES.azul, SpreadsheetApp.BorderStyle.SOLID);

  // ── FILAS DE DATOS 3–502 (500 filas) ──
  // Escribir números de fila (col A) en bloque
  var numeros = [];
  for (var i = 1; i <= 500; i++) {
    numeros.push([i]);
  }
  var colA = sheet.getRange(3, 1, 500, 1);
  colA.setValues(numeros);
  colA.setFontColor(COLORES.grisTexto);
  colA.setFontSize(9);
  colA.setHorizontalAlignment('center');
  colA.setVerticalAlignment('middle');

  // Formato de fecha en col B
  sheet.getRange(3, 2, 500, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(3, 2, 500, 1).setHorizontalAlignment('center');

  // Descripción col C
  sheet.getRange(3, 3, 500, 1).setHorizontalAlignment('left');

  // Monto col F
  var colF = sheet.getRange(3, 6, 500, 1);
  colF.setNumberFormat('#,##0');
  colF.setHorizontalAlignment('right');

  // Notas col G
  var colG = sheet.getRange(3, 7, 500, 1);
  colG.setFontSize(9);
  colG.setFontColor(COLORES.grisTexto);

  // Validación de datos: dropdown Categoría (col D)
  var reglaCat = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIAS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(3, 4, 500, 1).setDataValidation(reglaCat);

  // Validación de datos: dropdown Tipo (col E)
  var reglaTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Gasto', 'Ingreso'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(3, 5, 500, 1).setDataValidation(reglaTipo);
  sheet.getRange(3, 5, 500, 1).setHorizontalAlignment('center');

  // Altura de filas de datos
  for (var r = 3; r <= 502; r++) {
    sheet.setRowHeight(r, 22);
  }

  // Filas alternadas (blanco / gris muy claro)
  for (var f = 3; f <= 502; f++) {
    var fondo = (f % 2 === 1) ? COLORES.blanco : COLORES.grisFilas;
    sheet.getRange(f, 1, 1, 7).setBackground(fondo);
  }

  // ── FORMATO CONDICIONAL columna E (Tipo) ──
  var reglas = [];

  // Gasto → fondo rojo claro, texto rojo
  var reglaGasto = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Gasto')
    .setBackground(COLORES.rojoClaro)
    .setFontColor(COLORES.rojo)
    .setRanges([sheet.getRange('E3:E502')])
    .build();
  reglas.push(reglaGasto);

  // Ingreso → fondo verde claro, texto verde
  var reglaIngreso = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Ingreso')
    .setBackground(COLORES.verdeClaro)
    .setFontColor(COLORES.verde)
    .setRanges([sheet.getRange('E3:E502')])
    .build();
  reglas.push(reglaIngreso);

  sheet.setConditionalFormatRules(reglas);

  // ── FILA 503: separador visual ──
  sheet.setRowHeight(503, 10);
  sheet.getRange('A503:G503').setBackground(COLORES.azulClaro);

  // ── FILA 504: TOTAL GASTOS ──
  sheet.setRowHeight(504, 28);
  var totalGastosLabel = sheet.getRange('A504:E504');
  totalGastosLabel.merge();
  totalGastosLabel.setValue('TOTAL GASTOS');
  totalGastosLabel.setFontWeight('bold');
  totalGastosLabel.setFontSize(11);
  totalGastosLabel.setFontColor(COLORES.rojo);
  totalGastosLabel.setBackground(COLORES.rojoClaro);
  totalGastosLabel.setHorizontalAlignment('right');
  totalGastosLabel.setVerticalAlignment('middle');

  var totalGastosVal = sheet.getRange('F504');
  totalGastosVal.setFormula('=SUMIF(E3:E502,"Gasto",F3:F502)');
  totalGastosVal.setNumberFormat('#,##0');
  totalGastosVal.setFontWeight('bold');
  totalGastosVal.setFontSize(11);
  totalGastosVal.setFontColor(COLORES.rojo);
  totalGastosVal.setBackground(COLORES.rojoClaro);
  totalGastosVal.setHorizontalAlignment('right');

  sheet.getRange('G504').setBackground(COLORES.rojoClaro);

  // ── FILA 505: TOTAL INGRESOS ──
  sheet.setRowHeight(505, 28);
  var totalIngLabel = sheet.getRange('A505:E505');
  totalIngLabel.merge();
  totalIngLabel.setValue('TOTAL INGRESOS');
  totalIngLabel.setFontWeight('bold');
  totalIngLabel.setFontSize(11);
  totalIngLabel.setFontColor(COLORES.verde);
  totalIngLabel.setBackground(COLORES.verdeClaro);
  totalIngLabel.setHorizontalAlignment('right');
  totalIngLabel.setVerticalAlignment('middle');

  var totalIngVal = sheet.getRange('F505');
  totalIngVal.setFormula('=SUMIF(E3:E502,"Ingreso",F3:F502)');
  totalIngVal.setNumberFormat('#,##0');
  totalIngVal.setFontWeight('bold');
  totalIngVal.setFontSize(11);
  totalIngVal.setFontColor(COLORES.verde);
  totalIngVal.setBackground(COLORES.verdeClaro);
  totalIngVal.setHorizontalAlignment('right');

  sheet.getRange('G505').setBackground(COLORES.verdeClaro);

  // ── FILA 506: BALANCE ──
  sheet.setRowHeight(506, 28);
  var balLabel = sheet.getRange('A506:E506');
  balLabel.merge();
  balLabel.setValue('BALANCE');
  balLabel.setFontWeight('bold');
  balLabel.setFontSize(11);
  balLabel.setFontColor(COLORES.blanco);
  balLabel.setBackground(COLORES.azul);
  balLabel.setHorizontalAlignment('right');
  balLabel.setVerticalAlignment('middle');

  var balVal = sheet.getRange('F506');
  balVal.setFormula('=F505-F504');
  balVal.setNumberFormat('#,##0');
  balVal.setFontWeight('bold');
  balVal.setFontSize(11);
  balVal.setFontColor(COLORES.blanco);
  balVal.setBackground(COLORES.azul);
  balVal.setHorizontalAlignment('right');

  sheet.getRange('G506').setBackground(COLORES.azul);

  // ── CONGELAR fila de cabecera ──
  sheet.setFrozenRows(2);

  // Borde exterior de la tabla de datos
  sheet.getRange('A2:G506').setBorder(true, true, true, true, null, null, COLORES.azulClaro, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ============================================================
// HOJA DASHBOARD
// ============================================================
function construirDashboard(sheet) {
  sheet.clear();

  // ── Anchos de columna ──
  sheet.setColumnWidth(1, 35);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 120);

  // ── FILA 1: Título ──
  sheet.setRowHeight(1, 50);
  var titulo = sheet.getRange('A1:E1');
  titulo.merge();
  titulo.setValue('📊 DASHBOARD — Resumen de Gastos Personales');
  titulo.setBackground(COLORES.azul);
  titulo.setFontColor(COLORES.blanco);
  titulo.setFontSize(16);
  titulo.setFontWeight('bold');
  titulo.setHorizontalAlignment('center');
  titulo.setVerticalAlignment('middle');

  // ── FILA 2: Cabeceras de tarjetas ──
  sheet.setRowHeight(2, 30);

  // Total Gastos (A2:B2)
  var hGasto = sheet.getRange('A2:B2');
  hGasto.merge();
  hGasto.setValue('💸 Total Gastos');
  hGasto.setBackground(COLORES.rojo);
  hGasto.setFontColor(COLORES.blanco);
  hGasto.setFontWeight('bold');
  hGasto.setFontSize(11);
  hGasto.setHorizontalAlignment('center');
  hGasto.setVerticalAlignment('middle');

  // Total Ingresos (C2:D2)
  var hIngreso = sheet.getRange('C2:D2');
  hIngreso.merge();
  hIngreso.setValue('💰 Total Ingresos');
  hIngreso.setBackground(COLORES.verde);
  hIngreso.setFontColor(COLORES.blanco);
  hIngreso.setFontWeight('bold');
  hIngreso.setFontSize(11);
  hIngreso.setHorizontalAlignment('center');
  hIngreso.setVerticalAlignment('middle');

  // Balance (E2)
  var hBalance = sheet.getRange('E2');
  hBalance.setValue('⚖️ Balance');
  hBalance.setBackground(COLORES.azul);
  hBalance.setFontColor(COLORES.blanco);
  hBalance.setFontWeight('bold');
  hBalance.setFontSize(11);
  hBalance.setHorizontalAlignment('center');
  hBalance.setVerticalAlignment('middle');

  // ── FILA 3: Valores de tarjetas ──
  sheet.setRowHeight(3, 45);

  // Valor Total Gastos
  var vGasto = sheet.getRange('A3:B3');
  vGasto.merge();
  vGasto.setFormula('=SUMIF(\'💳 Transacciones\'!E3:E502,"Gasto",\'💳 Transacciones\'!F3:F502)');
  vGasto.setNumberFormat('#,##0');
  vGasto.setFontSize(16);
  vGasto.setFontWeight('bold');
  vGasto.setFontColor(COLORES.rojo);
  vGasto.setBackground(COLORES.rojoClaro);
  vGasto.setHorizontalAlignment('center');
  vGasto.setVerticalAlignment('middle');

  // Valor Total Ingresos
  var vIngreso = sheet.getRange('C3:D3');
  vIngreso.merge();
  vIngreso.setFormula('=SUMIF(\'💳 Transacciones\'!E3:E502,"Ingreso",\'💳 Transacciones\'!F3:F502)');
  vIngreso.setNumberFormat('#,##0');
  vIngreso.setFontSize(16);
  vIngreso.setFontWeight('bold');
  vIngreso.setFontColor(COLORES.verde);
  vIngreso.setBackground(COLORES.verdeClaro);
  vIngreso.setHorizontalAlignment('center');
  vIngreso.setVerticalAlignment('middle');

  // Valor Balance
  var vBalance = sheet.getRange('E3');
  vBalance.setFormula('=C3-A3');
  vBalance.setNumberFormat('#,##0');
  vBalance.setFontSize(16);
  vBalance.setFontWeight('bold');
  vBalance.setFontColor(COLORES.blanco);
  vBalance.setBackground(COLORES.azul);
  vBalance.setHorizontalAlignment('center');
  vBalance.setVerticalAlignment('middle');

  // ── FILA 4: Espacio ──
  sheet.setRowHeight(4, 18);
  sheet.getRange('A4:E4').setBackground(COLORES.crema);

  // ── FILA 5: Cabecera de sección Categorías ──
  sheet.setRowHeight(5, 32);
  var secCat = sheet.getRange('A5:E5');
  secCat.merge();
  secCat.setValue('📂 GASTOS POR CATEGORÍA');
  secCat.setBackground(COLORES.azulClaro);
  secCat.setFontColor(COLORES.azul);
  secCat.setFontWeight('bold');
  secCat.setFontSize(12);
  secCat.setHorizontalAlignment('center');
  secCat.setVerticalAlignment('middle');

  // ── FILA 6: Cabeceras de tabla de categorías ──
  sheet.setRowHeight(6, 28);
  var cabCat = [['#', 'Categoría', 'Total ($)', '% del Total', 'Visual']];
  var cabRange = sheet.getRange('A6:E6');
  cabRange.setValues(cabCat);
  cabRange.setBackground(COLORES.azul);
  cabRange.setFontColor(COLORES.blanco);
  cabRange.setFontWeight('bold');
  cabRange.setFontSize(10);
  cabRange.setHorizontalAlignment('center');
  cabRange.setVerticalAlignment('middle');

  // ── FILAS 7–18: Una fila por categoría ──
  for (var i = 0; i < CATEGORIAS.length; i++) {
    var fila    = 7 + i;
    var cat     = CATEGORIAS[i];
    var fondoCat = COLORES_CATEGORIA[i];

    sheet.setRowHeight(fila, 24);

    // Número
    var celdaNum = sheet.getRange(fila, 1);
    celdaNum.setValue(i + 1);
    celdaNum.setBackground(fondoCat);
    celdaNum.setFontColor(COLORES.grisTexto);
    celdaNum.setFontSize(9);
    celdaNum.setHorizontalAlignment('center');
    celdaNum.setVerticalAlignment('middle');

    // Nombre de categoría
    var celdaCat = sheet.getRange(fila, 2);
    celdaCat.setValue(cat);
    celdaCat.setBackground(fondoCat);
    celdaCat.setFontSize(10);
    celdaCat.setHorizontalAlignment('left');
    celdaCat.setVerticalAlignment('middle');

    // Total SUMIFS por categoría
    var celdaTotal = sheet.getRange(fila, 3);
    celdaTotal.setFormula(
      '=SUMIFS(\'💳 Transacciones\'!F3:F502,\'💳 Transacciones\'!D3:D502,"' + cat + '",\'💳 Transacciones\'!E3:E502,"Gasto")'
    );
    celdaTotal.setNumberFormat('#,##0');
    celdaTotal.setBackground(fondoCat);
    celdaTotal.setFontSize(10);
    celdaTotal.setHorizontalAlignment('right');
    celdaTotal.setVerticalAlignment('middle');

    // Porcentaje
    var celdaPct = sheet.getRange(fila, 4);
    celdaPct.setFormula('=IF(A3=0,0,C' + fila + '/A3)');
    celdaPct.setNumberFormat('0.0%');
    celdaPct.setBackground(fondoCat);
    celdaPct.setFontSize(10);
    celdaPct.setHorizontalAlignment('center');
    celdaPct.setVerticalAlignment('middle');

    // Barra visual con REPT
    var celdaBarra = sheet.getRange(fila, 5);
    celdaBarra.setFormula('=REPT("█",ROUND(D' + fila + '*30,0))');
    celdaBarra.setBackground(fondoCat);
    celdaBarra.setFontColor(COLORES.azul);
    celdaBarra.setFontSize(9);
    celdaBarra.setHorizontalAlignment('left');
    celdaBarra.setVerticalAlignment('middle');
  }

  // ── FILA 19: Totales de la tabla ──
  var filaTotales = 7 + CATEGORIAS.length; // fila 19
  sheet.setRowHeight(filaTotales, 28);

  var totLabel = sheet.getRange(filaTotales, 1, 1, 2);
  totLabel.merge();
  totLabel.setValue('TOTAL');
  totLabel.setBackground(COLORES.azulClaro);
  totLabel.setFontColor(COLORES.azul);
  totLabel.setFontWeight('bold');
  totLabel.setFontSize(11);
  totLabel.setHorizontalAlignment('center');
  totLabel.setVerticalAlignment('middle');

  var totVal = sheet.getRange(filaTotales, 3);
  totVal.setFormula('=SUM(C7:C18)');
  totVal.setNumberFormat('#,##0');
  totVal.setBackground(COLORES.azulClaro);
  totVal.setFontColor(COLORES.azul);
  totVal.setFontWeight('bold');
  totVal.setFontSize(11);
  totVal.setHorizontalAlignment('right');
  totVal.setVerticalAlignment('middle');

  var totPct = sheet.getRange(filaTotales, 4);
  totPct.setFormula('=SUM(D7:D18)');
  totPct.setNumberFormat('0.0%');
  totPct.setBackground(COLORES.azulClaro);
  totPct.setFontColor(COLORES.azul);
  totPct.setFontWeight('bold');
  totPct.setFontSize(11);
  totPct.setHorizontalAlignment('center');
  totPct.setVerticalAlignment('middle');

  sheet.getRange(filaTotales, 5).setBackground(COLORES.azulClaro);

  // ── Bordes de la tabla de categorías ──
  sheet.getRange(6, 1, CATEGORIAS.length + 2, 5)
    .setBorder(true, true, true, true, true, true, COLORES.azulClaro, SpreadsheetApp.BorderStyle.SOLID);

  // ── Separador final ──
  var filaFin = filaTotales + 1;
  sheet.setRowHeight(filaFin, 10);
  sheet.getRange(filaFin, 1, 1, 5).setBackground(COLORES.azul);

  // ── Nota al pie ──
  var filaNota = filaFin + 1;
  sheet.setRowHeight(filaNota, 22);
  var nota = sheet.getRange(filaNota, 1, 1, 5);
  nota.merge();
  nota.setValue('Última actualización: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '  •  Los datos se actualizan automáticamente al editar la hoja de Transacciones.');
  nota.setFontColor(COLORES.grisTexto);
  nota.setFontSize(8);
  nota.setHorizontalAlignment('center');
  nota.setBackground(COLORES.crema);

  // ── Fondo general de la hoja ──
  sheet.getRange('A1:E' + (filaNota + 2)).setBackground(COLORES.crema);

  // Re-aplicar fondos específicos encima del fondo general
  // (ya se hicieron célula a célula, así que no se necesita override)

  // ── CONGELAR fila de título ──
  sheet.setFrozenRows(1);
}

// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

// Obtiene una hoja existente o la crea nueva; si existe la limpia
function obtenerOCrearHoja(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    // Eliminar todas las validaciones existentes
    var ultimaFila = Math.max(sheet.getMaxRows(), 1);
    var ultimaCol  = Math.max(sheet.getMaxColumns(), 1);
    sheet.getRange(1, 1, ultimaFila, ultimaCol).clearDataValidations();
  } else {
    sheet = ss.insertSheet(nombre);
  }
  return sheet;
}

// Elimina "Hoja 1", "Sheet1" u hojas vacías por defecto si existen
function eliminarHojaDefault(ss) {
  var nombresDefault = ['Hoja 1', 'Hoja1', 'Sheet1', 'Sheet 1'];
  nombresDefault.forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    if (h) {
      // Solo eliminar si la hoja tiene pocas filas con datos (es la hoja default vacía)
      try {
        ss.deleteSheet(h);
      } catch (e) {
        // No se puede eliminar si es la única hoja; ignorar
      }
    }
  });
}

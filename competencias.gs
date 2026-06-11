// ============================================================
// COMPETENCIAS LAFOURCADE — Google Apps Script
// Instrucciones: Extensiones > Apps Script > pegar > Ejecutar setupCompetencias
// ============================================================

var C = {
  azul:     '#1e40af',
  azulL:    '#dbeafe',
  crema:    '#fef9ec',
  blanco:   '#ffffff',
  gris:     '#f1f5f9',
  grisT:    '#6b7280',
  verde:    '#16a34a',
  verdeL:   '#dcfce7',
  rojo:     '#dc2626',
  amarillo: '#d97706',
  amarilloL:'#fef3c7'
};

var GRUPOS_COLORES = ['#1e40af','#7c3aed','#db2777','#dc2626','#d97706','#16a34a','#0891b2','#059669'];

// ── Función principal ──────────────────────────────────────
function setupCompetencias() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHojaLiga(ss);
  crearHojaGrupos(ss);
  crearHojaEliminacion(ss);
  crearHojaMixto(ss);
  eliminarHojaDefault(ss);
  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName('📊 Liga'));
  SpreadsheetApp.getUi().alert(
    '✅ ¡Listo!',
    'Hojas creadas:\n📊 Liga  →  cargá participantes en col B, partidos en D-G\n👥 Grupos  →  asignales grupo en col B\n⚔️ Eliminación  →  ingresá resultados por ronda\n🔀 Mixto  →  se completa desde Grupos',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── Helpers generales ─────────────────────────────────────
function obtenerHoja(ss, nombre) {
  var h = ss.getSheetByName(nombre);
  if (h) ss.deleteSheet(h);
  return ss.insertSheet(nombre);
}

function titulo(sh, texto, numCols) {
  sh.getRange(1, 1, 1, numCols).merge()
    .setValue(texto)
    .setBackground(C.azul).setFontColor('#fff')
    .setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);
}

function seccion(sh, fila, colInicio, texto, colspan) {
  sh.getRange(fila, colInicio, 1, colspan).merge()
    .setValue(texto)
    .setBackground(C.crema).setFontColor('#1e293b')
    .setFontWeight('bold').setFontSize(10)
    .setBorder(false,false,true,false,false,false, '#ddd3c3', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(fila, 24);
}

function encabezados(sh, fila, colInicio, cabeceras, bgColor) {
  bgColor = bgColor || C.azul;
  var fgColor = (bgColor === C.azul) ? '#fff' : '#1e293b';
  cabeceras.forEach(function(h, i) {
    sh.getRange(fila, colInicio + i)
      .setValue(h)
      .setBackground(bgColor).setFontColor(fgColor)
      .setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.setRowHeight(fila, 28);
}

function eliminarHojaDefault(ss) {
  ['Hoja 1','Hoja1','Sheet1','Sheet 1'].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
}

// ============================================================
// HOJA LIGA
// ============================================================
function crearHojaLiga(ss) {
  var sh = obtenerHoja(ss, '📊 Liga');

  // anchos: A(#) B(nombre) C(gap) D(local) E(golL) F(golV) G(visit) H(gap) I(#) J(nombre) K(pts) L-R(stats)
  var anchos = [35, 175, 18, 175, 65, 65, 175, 18, 35, 175, 52, 42, 42, 42, 42, 45, 45, 52];
  anchos.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  titulo(sh, '🏆 LIGA — Todos contra todos', 18);

  // ── Participantes ──
  seccion(sh, 2, 1, '① PARTICIPANTES', 2);
  encabezados(sh, 3, 1, ['#', 'Nombre del participante']);
  sh.getRange(3, 2).setHorizontalAlignment('left');

  for (var i = 1; i <= 30; i++) {
    var r = i + 3;
    sh.getRange(r, 1).setValue(i).setFontColor(C.grisT).setHorizontalAlignment('center').setFontSize(10);
    sh.getRange(r, 2).setBackground(i % 2 ? C.blanco : C.gris).setFontSize(11);
    sh.setRowHeight(r, 22);
  }

  // Dropdown local/visitante desde lista de participantes
  var listaJugadores = sh.getRange('$B$4:$B$33');
  var reglaDrop = SpreadsheetApp.newDataValidation().requireValueInRange(listaJugadores, true).build();
  sh.getRange('D4:D300').setDataValidation(reglaDrop);
  sh.getRange('G4:G300').setDataValidation(reglaDrop);

  // ── Partidos ──
  seccion(sh, 2, 4, '② PARTIDOS — Ingresá los marcadores acá', 4);
  encabezados(sh, 3, 4, ['LOCAL', 'GOL L', 'GOL V', 'VISITANTE']);
  sh.getRange(3, 5).setHorizontalAlignment('center');
  sh.getRange(3, 6).setHorizontalAlignment('center');

  for (var j = 4; j <= 300; j++) {
    sh.getRange(j, 4, 1, 4).setBackground(j % 2 ? C.blanco : C.gris);
    sh.getRange(j, 5).setHorizontalAlignment('center');
    sh.getRange(j, 6).setHorizontalAlignment('center');
    sh.setRowHeight(j, 22);
  }

  // ── Tabla de Posiciones ──
  seccion(sh, 2, 9, '③ TABLA DE POSICIONES', 10);
  encabezados(sh, 3, 9, ['#', 'PARTICIPANTE', 'PTS', 'PJ', 'PG', 'PE', 'PP', 'GF', 'GC', 'DG']);
  sh.getRange(3, 10).setHorizontalAlignment('left');

  var D = '$D$4:$D$300', G = '$G$4:$G$300';
  var E = '$E$4:$E$300', F = '$F$4:$F$300';
  var jugado = '(' + E + '<>"")*(' + F + '<>"")';

  for (var r2 = 0; r2 < 30; r2++) {
    var row = r2 + 4;
    var J = 'J' + row;
    var bg = r2 % 2 ? C.blanco : C.gris;

    sh.getRange(row, 9).setValue(r2 + 1).setFontColor(C.grisT).setHorizontalAlignment('center');
    sh.getRange(row, 10).setFormula('=IFERROR(IF(B' + row + '="","",B' + row + '),"")');

    // PJ
    sh.getRange(row, 12).setFormula(
      '=IF(' + J + '="","",SUMPRODUCT((' + D + '=' + J + ')*' + jugado + ')' +
      '+SUMPRODUCT((' + G + '=' + J + ')*' + jugado + '))'
    );
    // PG
    sh.getRange(row, 13).setFormula(
      '=IF(' + J + '="","",SUMPRODUCT((' + D + '=' + J + ')*(' + E + '>' + F + ')*' + jugado + ')' +
      '+SUMPRODUCT((' + G + '=' + J + ')*(' + F + '>' + E + ')*' + jugado + '))'
    );
    // PE
    sh.getRange(row, 14).setFormula(
      '=IF(' + J + '="","",SUMPRODUCT((' + D + '=' + J + ')*(' + E + '=' + F + ')*' + jugado + ')' +
      '+SUMPRODUCT((' + G + '=' + J + ')*(' + F + '=' + E + ')*' + jugado + '))'
    );
    // PP
    sh.getRange(row, 15).setFormula('=IF(' + J + '="","",L' + row + '-M' + row + '-N' + row + ')');
    // GF
    sh.getRange(row, 16).setFormula(
      '=IF(' + J + '="","",SUMPRODUCT((' + D + '=' + J + ')*IFERROR(' + E + '*1,0))' +
      '+SUMPRODUCT((' + G + '=' + J + ')*IFERROR(' + F + '*1,0)))'
    );
    // GC
    sh.getRange(row, 17).setFormula(
      '=IF(' + J + '="","",SUMPRODUCT((' + D + '=' + J + ')*IFERROR(' + F + '*1,0))' +
      '+SUMPRODUCT((' + G + '=' + J + ')*IFERROR(' + E + '*1,0)))'
    );
    // DG
    sh.getRange(row, 18).setFormula('=IF(' + J + '="","",P' + row + '-Q' + row + ')');
    // PTS
    sh.getRange(row, 11).setFormula('=IF(' + J + '="","",M' + row + '*3+N' + row + ')');

    sh.getRange(row, 9, 1, 10).setBackground(bg).setFontSize(10);
    sh.getRange(row, 11).setFontColor(C.azul).setFontWeight('bold');
    sh.getRange(row, 11, 1, 8).setHorizontalAlignment('center');
    sh.getRange(row, 10).setHorizontalAlignment('left');
    sh.setRowHeight(row, 22);
  }

  // Tip de ordenamiento
  sh.getRange('A36').setValue('💡 Para ver la tabla ordenada: seleccioná I3:R33 → Datos → Ordenar rango por columna K (PTS), de mayor a menor.')
    .setFontColor(C.grisT).setFontSize(9).setFontStyle('italic');
  sh.getRange('A36:R36').merge();

  sh.setFrozenRows(3);
}

// ============================================================
// HOJA GRUPOS
// ============================================================
function crearHojaGrupos(ss) {
  var sh = obtenerHoja(ss, '👥 Grupos');

  var anchos = [35, 100, 170, 18, 170, 65, 65, 170, 80];
  anchos.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  titulo(sh, '👥 GRUPOS — Fase de grupos', 9);

  // Config
  sh.getRange('A2').setValue('N° de grupos:').setFontWeight('bold').setFontSize(10);
  sh.getRange('B2').setValue(4).setBackground(C.azulL).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setBorder(true,true,true,true,false,false);
  sh.getRange('C2').setValue('← editá acá (1-8)').setFontColor(C.grisT).setFontSize(9).setFontStyle('italic');
  sh.getRange('D2').setValue('Clasifican por grupo:').setFontWeight('bold').setFontSize(10);
  sh.getRange('E2').setValue(2).setBackground(C.azulL).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setBorder(true,true,true,true,false,false);
  sh.getRange('F2').setValue('← editá acá').setFontColor(C.grisT).setFontSize(9).setFontStyle('italic');
  sh.setRowHeight(2, 28);

  // ── Participantes ──
  seccion(sh, 3, 1, '① PARTICIPANTES — asignales un grupo', 3);
  encabezados(sh, 4, 1, ['#', 'GRUPO', 'NOMBRE']);
  sh.getRange(4, 2).setHorizontalAlignment('center');
  sh.getRange(4, 3).setHorizontalAlignment('left');

  var optsGrupo = ['Grupo A','Grupo B','Grupo C','Grupo D','Grupo E','Grupo F','Grupo G','Grupo H'];
  var reglaGrupo = SpreadsheetApp.newDataValidation().requireValueInList(optsGrupo, true).build();

  for (var i = 1; i <= 32; i++) {
    var r = i + 4;
    sh.getRange(r, 1).setValue(i).setFontColor(C.grisT).setHorizontalAlignment('center').setFontSize(10);
    sh.getRange(r, 2).setDataValidation(reglaGrupo).setValue('Grupo A').setHorizontalAlignment('center').setFontSize(10);
    sh.getRange(r, 3).setBackground(i % 2 ? C.blanco : C.gris).setFontSize(11);
    sh.setRowHeight(r, 22);
  }

  // ── Partidos ──
  seccion(sh, 3, 5, '② PARTIDOS — incluí el grupo en col I', 5);
  encabezados(sh, 4, 5, ['LOCAL', 'GOL L', 'GOL V', 'VISITANTE', 'GRUPO']);
  sh.getRange(4, 6).setHorizontalAlignment('center');
  sh.getRange(4, 7).setHorizontalAlignment('center');
  sh.getRange(4, 9).setHorizontalAlignment('center');

  var reglaGrupoPartido = SpreadsheetApp.newDataValidation().requireValueInList(optsGrupo, true).build();

  for (var j = 5; j <= 200; j++) {
    sh.getRange(j, 5, 1, 5).setBackground(j % 2 ? C.blanco : C.gris);
    sh.getRange(j, 6).setHorizontalAlignment('center');
    sh.getRange(j, 7).setHorizontalAlignment('center');
    sh.getRange(j, 9).setDataValidation(reglaGrupoPartido).setHorizontalAlignment('center');
    sh.setRowHeight(j, 22);
  }

  sh.setFrozenRows(4);

  // ── Mini-tablas por grupo (debajo de los datos) ──
  var startRow = 42;
  var nombresGrupos = ['Grupo A','Grupo B','Grupo C','Grupo D'];
  var D2 = '$E$5:$E$200', E2 = '$F$5:$F$200', F2 = '$G$5:$G$200', G2 = '$H$5:$H$200', I2 = '$I$5:$I$200';
  var jugado2 = '(' + E2 + '<>"")*(' + F2 + '<>"")';

  for (var g = 0; g < 4; g++) {
    var colBase = (g % 2 === 0) ? 1 : 5;
    var rowBase = (g < 2) ? startRow : startRow + 13;
    var gNombre = nombresGrupos[g];
    var gColor  = GRUPOS_COLORES[g];

    // Header del grupo
    sh.getRange(rowBase, colBase, 1, 4).merge()
      .setValue('Tabla ' + gNombre)
      .setBackground(gColor).setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    sh.setRowHeight(rowBase, 28);

    encabezados(sh, rowBase + 1, colBase, ['PARTICIPANTE', 'PTS', 'PJ', 'DG'], gColor + 'aa');
    sh.getRange(rowBase + 1, colBase).setHorizontalAlignment('left');

    for (var p = 0; p < 8; p++) {
      var pRow = rowBase + 2 + p;
      var bg2 = p % 2 ? C.blanco : C.gris;

      // Nombre del p-ésimo jugador en este grupo (usando FILTER)
      sh.getRange(pRow, colBase).setFormula(
        '=IFERROR(INDEX(FILTER($C$5:$C$36,$B$5:$B$36="' + gNombre + '"),' + (p + 1) + '),"")'
      ).setBackground(bg2).setFontSize(10);

      var nm = sh.getRange(pRow, colBase).getA1Notation();

      // PTS con filtro de grupo
      sh.getRange(pRow, colBase + 1).setFormula(
        '=IF(' + nm + '="","",(' +
        'SUMPRODUCT((' + D2 + '=' + nm + ')*(' + E2 + '>' + F2 + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '"))' +
        '+SUMPRODUCT((' + G2 + '=' + nm + ')*(' + F2 + '>' + E2 + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '")))*3+(' +
        'SUMPRODUCT((' + D2 + '=' + nm + ')*(' + E2 + '=' + F2 + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '"))' +
        '+SUMPRODUCT((' + G2 + '=' + nm + ')*(' + F2 + '=' + E2 + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '"))))'
      ).setBackground(bg2).setFontSize(10).setHorizontalAlignment('center')
        .setFontColor(C.azul).setFontWeight('bold');

      // PJ con filtro de grupo
      sh.getRange(pRow, colBase + 2).setFormula(
        '=IF(' + nm + '="","",SUMPRODUCT((' + D2 + '=' + nm + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '"))' +
        '+SUMPRODUCT((' + G2 + '=' + nm + ')*' + jugado2 + '*(' + I2 + '="' + gNombre + '")))'
      ).setBackground(bg2).setFontSize(10).setHorizontalAlignment('center');

      // DG con filtro de grupo
      sh.getRange(pRow, colBase + 3).setFormula(
        '=IF(' + nm + '="","",(' +
        'SUMPRODUCT((' + D2 + '=' + nm + ')*IFERROR(' + E2 + '*1,0)*(' + I2 + '="' + gNombre + '"))' +
        '+SUMPRODUCT((' + G2 + '=' + nm + ')*IFERROR(' + F2 + '*1,0)*(' + I2 + '="' + gNombre + '")))-(' +
        'SUMPRODUCT((' + D2 + '=' + nm + ')*IFERROR(' + F2 + '*1,0)*(' + I2 + '="' + gNombre + '"))' +
        '+SUMPRODUCT((' + G2 + '=' + nm + ')*IFERROR(' + E2 + '*1,0)*(' + I2 + '="' + gNombre + '"))))'
      ).setBackground(bg2).setFontSize(10).setHorizontalAlignment('center');

      sh.setRowHeight(pRow, 22);
    }
  }
}

// ============================================================
// HOJA ELIMINACIÓN
// ============================================================
function crearHojaEliminacion(ss) {
  var sh = obtenerHoja(ss, '⚔️ Eliminación');

  // Anchos: A(lista#) B(lista nombre) C(gap) D(equipo) E(score) F(vs) G(score) H(equipo) I(ganador)
  var anchos = [35, 165, 20, 165, 55, 35, 55, 165, 165];
  anchos.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  titulo(sh, '⚔️ ELIMINACIÓN DIRECTA', 9);

  // Instrucciones
  var instrucciones = [
    ['① Cargá los participantes en la tabla de la izquierda (col A-B)'],
    ['② Copiá los nombres en los slots de la primera ronda (col D y H)'],
    ['③ Ingresá los resultados (cols E y G). El ganador en col I se llena solo con fórmula'],
    ['④ En la siguiente ronda, en col D/H referenciá la celda del ganador anterior (ej: =I5)']
  ];
  instrucciones.forEach(function(inst, i) {
    sh.getRange(i + 2, 1, 1, 9).merge().setValue(inst[0])
      .setBackground(i % 2 ? C.azulL : C.crema)
      .setFontColor('#1e293b').setFontSize(9)
      .setFontStyle(i === 3 ? 'italic' : 'normal');
    sh.setRowHeight(i + 2, 22);
  });

  // Lista de participantes
  seccion(sh, 7, 1, 'PARTICIPANTES', 2);
  encabezados(sh, 8, 1, ['#', 'Nombre']);
  sh.getRange(8, 2).setHorizontalAlignment('left');
  for (var i = 1; i <= 16; i++) {
    sh.getRange(i + 8, 1).setValue(i).setFontColor(C.grisT).setHorizontalAlignment('center');
    sh.getRange(i + 8, 2).setBackground(i % 2 ? C.blanco : C.gris).setFontSize(11);
    sh.setRowHeight(i + 8, 22);
  }

  // Estructura de partidos (col D en adelante)
  var rondas = ['CUARTOS DE FINAL', 'SEMIFINALES', 'FINAL'];
  seccion(sh, 7, 4, 'BRACKET', 6);
  encabezados(sh, 8, 4, ['EQUIPO 1', 'GOL', 'vs', 'GOL', 'EQUIPO 2', 'GANADOR']);
  sh.getRange(8, 6).setValue('vs').setHorizontalAlignment('center');

  // Generar hasta 8 partidos (cuartos en una plantilla básica)
  var matchLabels = [
    'Cuartos — Partido 1', 'Cuartos — Partido 2', 'Cuartos — Partido 3', 'Cuartos — Partido 4',
    'Semifinal 1', 'Semifinal 2',
    'FINAL'
  ];
  var matchColors = [C.azulL, C.azulL, C.azulL, C.azulL, C.crema, C.crema, C.amarilloL];

  for (var m = 0; m < 7; m++) {
    var mRow = 10 + m * 3;
    var bg3 = matchColors[m];

    // Label del partido
    sh.getRange(mRow, 4, 1, 6).merge().setValue(matchLabels[m])
      .setBackground(m === 6 ? C.amarillo : C.azul)
      .setFontColor('#fff').setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center');
    sh.setRowHeight(mRow, 20);

    // Fila de datos del partido
    var dRow = mRow + 1;
    sh.getRange(dRow, 4).setBackground(bg3).setFontSize(11);           // Equipo 1
    sh.getRange(dRow, 5).setBackground(bg3).setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold'); // Gol 1
    sh.getRange(dRow, 6).setValue('-').setBackground(bg3).setHorizontalAlignment('center').setFontColor(C.grisT);  // vs
    sh.getRange(dRow, 7).setBackground(bg3).setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold'); // Gol 2
    sh.getRange(dRow, 8).setBackground(bg3).setFontSize(11);           // Equipo 2

    // Fórmula ganador: si Gol1 > Gol2 → Equipo1, si Gol2 > Gol1 → Equipo2, si vacío → ""
    var d = dRow;
    sh.getRange(d, 9).setFormula(
      '=IF(OR(E' + d + '="",G' + d + '=""),"",IF(E' + d + '>G' + d + ',D' + d + ',IF(G' + d + '>E' + d + ',H' + d + ',"Empate")))'
    ).setBackground(m === 6 ? C.amarilloL : C.verdeL)
      .setFontColor(m === 6 ? C.amarillo : C.verde)
      .setFontWeight('bold').setFontSize(m === 6 ? 12 : 10)
      .setHorizontalAlignment('center');

    sh.setRowHeight(dRow, 26);
    sh.setRowHeight(dRow + 1, 8); // spacer
  }

  // Campeón
  var champRow = 10 + 7 * 3;
  sh.getRange(champRow, 4, 1, 6).merge()
    .setValue('🏆 CAMPEÓN')
    .setBackground(C.amarillo).setFontColor('#fff')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.getRange(champRow + 1, 4, 1, 6).merge()
    .setFormula('=I' + (10 + 6 * 3 + 1))
    .setBackground(C.amarilloL).setFontColor(C.amarillo)
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sh.setRowHeight(champRow, 32);
  sh.setRowHeight(champRow + 1, 40);

  sh.setFrozenRows(8);
}

// ============================================================
// HOJA MIXTO
// ============================================================
function crearHojaMixto(ss) {
  var sh = obtenerHoja(ss, '🔀 Mixto');

  var anchos = [35, 200, 100, 35, 200, 100];
  anchos.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  titulo(sh, '🔀 MIXTO — Grupos + Eliminación Directa', 6);

  // Instrucciones paso a paso
  var pasos = [
    '① Completá todos los partidos en la hoja 👥 Grupos.',
    '② Los clasificados de cada grupo aparecen en las mini-tablas al final de esa hoja.',
    '③ Copiá los clasificados en la tabla de abajo (solo valores, sin fórmulas).',
    '④ Llevá esos nombres a la hoja ⚔️ Eliminación para armar la llave final.'
  ];
  pasos.forEach(function(p, i) {
    sh.getRange(i + 2, 1, 1, 6).merge().setValue(p)
      .setBackground(i % 2 ? C.azulL : C.crema)
      .setFontColor('#1e293b').setFontSize(10);
    sh.setRowHeight(i + 2, 26);
  });

  // Tabla de clasificados
  sh.getRange('A7:F7').merge()
    .setValue('CLASIFICADOS DE GRUPOS')
    .setBackground(C.azul).setFontColor('#fff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sh.setRowHeight(7, 30);

  encabezados(sh, 8, 1, ['Pos.', 'Nombre', 'Grupo', 'Pos.', 'Nombre', 'Grupo']);

  var gruposDisplay = [
    {nombre: 'Grupo A', col: C.azul},
    {nombre: 'Grupo B', col: '#7c3aed'},
    {nombre: 'Grupo C', col: '#db2777'},
    {nombre: 'Grupo D', col: '#dc2626'}
  ];

  // 2 clasificados por grupo (col 1-3 = grupos A y B, col 4-6 = grupos C y D)
  for (var g = 0; g < 4; g++) {
    var colBase2 = g < 2 ? 1 : 4;
    var subPos   = g % 2; // 0 = grupo A/C, 1 = grupo B/D → usa filas distintas
    var gColor2  = gruposDisplay[g].col;
    var gNombre2 = gruposDisplay[g].nombre;

    for (var pos = 1; pos <= 2; pos++) {
      var fRow = 9 + subPos * 2 + (pos - 1);
      if (g >= 2) fRow = 9 + (g - 2) * 2 + (pos - 1);
      // Ajustar filas para layout en 2 columnas
      var dRow2 = (g < 2) ? 9 + g * 2 + (pos - 1) : 9 + (g - 2) * 2 + (pos - 1);

      sh.getRange(dRow2, colBase2).setValue(pos).setFontColor(C.grisT).setHorizontalAlignment('center');
      sh.getRange(dRow2, colBase2 + 1).setBackground(gColor2 + '22').setFontSize(11);
      sh.getRange(dRow2, colBase2 + 2).setValue(gNombre2)
        .setBackground(gColor2 + '22').setFontColor(gColor2)
        .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
      sh.setRowHeight(dRow2, 24);
    }
  }

  sh.getRange('A17:F17').merge()
    .setValue('💡 Completá los nombres de los clasificados arriba, luego copiálos a ⚔️ Eliminación.')
    .setFontColor(C.grisT).setFontSize(9).setFontStyle('italic');
}

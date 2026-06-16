// ============================================================
// ACADEMIA DE PÁDEL — Google Apps Script  |  academia_padel.gs
//
// INSTRUCCIONES:
//   1. Abrí el Google Sheet en Google Drive (ej: "Academia Pádel")
//   2. Extensiones → Apps Script → borrar código existente → pegar este archivo
//   3. Guardá (Ctrl+S) y seleccioná la función: crearAcademiaPadel
//   4. Ejecutá ▶ y aceptá los permisos
//
// ARQUITECTURA (Fase 1 — Google Sheets):
//   Cada hoja representa una TABLA normalizada, relacionada por IDs.
//   La nomenclatura interna usa inglés para facilitar migración futura.
//   IDs únicos: J001-J008 = jugadores, P001-P008 = pagos,
//               E001-E008 = evaluaciones, T001-T007 = torneos
// ============================================================

// ── PALETA DE COLORES ─────────────────────────────────────────
var COL = {
  darkBlue:  '#1A3A5C',  // encabezados principales  → futuro: color de role "admin"
  medBlue:   '#2E75B6',  // encabezados secundarios
  lightBlue: '#D6E4F0',  // filas alternas
  cream:     '#FFF8F0',  // fondo general / filas base
  white:     '#FFFFFF',
  green:     '#70AD47',  // semáforo OK
  yellow:    '#FFD700',  // semáforo alerta
  red:       '#FF4444',  // semáforo peligro
  gray:      '#6B7280',
};

// ── TABLA PLAYERS — datos coherentes entre todas las hojas ────
// NOTA MIGRACIÓN: estos objetos son el esquema de la futura colección "players"
var PLAYERS = [
  { id:'J001', name:'Carlos García López',   dob:'15/03/1990', gender:'M', dni:'12345678A', phone:'612345678', email:'carlos@email.com', emergency:'Laura García 611111111',    position:'Drive', level:'3ª', ranking:250, joined:'10/01/2023', status:'Activo',         blood:'A+', medical:'Apto',    notes:'Jugador con mucha proyección' },
  { id:'J002', name:'María Martínez Ruiz',   dob:'22/07/1995', gender:'F', dni:'87654321B', phone:'623456789', email:'maria@email.com',  emergency:'José Martínez 622222222',  position:'Revés', level:'4ª', ranking:380, joined:'15/02/2023', status:'Activo',         blood:'O+', medical:'Apto',    notes:'Muy constante en entrenamientos' },
  { id:'J003', name:'Pablo Sánchez Torres',  dob:'08/11/1988', gender:'M', dni:'11223344C', phone:'634567890', email:'pablo@email.com',  emergency:'Rosa Sánchez 633333333',   position:'Ambas', level:'2ª', ranking:120, joined:'05/03/2022', status:'Activo',         blood:'B+', medical:'Apto',    notes:'Experiencia en torneos nacionales' },
  { id:'J004', name:'Ana López Fernández',   dob:'30/05/2005', gender:'F', dni:'44556677D', phone:'645678901', email:'ana@email.com',    emergency:'Pedro López 644444444',    position:'Drive', level:'5ª', ranking:510, joined:'20/09/2023', status:'Activo',         blood:'A-', medical:'Apto',    notes:'Juvenil con gran potencial' },
  { id:'J005', name:'Diego Romero Vega',     dob:'14/02/1992', gender:'M', dni:'55667788E', phone:'656789012', email:'diego@email.com',  emergency:'Carmen Romero 655555555', position:'Revés', level:'3ª', ranking:290, joined:'12/06/2022', status:'Baja temporal',  blood:'AB+',medical:'No apto', notes:'Lesión rodilla izquierda' },
  { id:'J006', name:'Lucía Torres Morales',  dob:'19/09/1998', gender:'F', dni:'66778899F', phone:'667890123', email:'lucia@email.com',  emergency:'Marta Torres 666666666',   position:'Drive', level:'4ª', ranking:400, joined:'08/01/2024', status:'Activo',         blood:'O-', medical:'Apto',    notes:'Buena pegada de bandeja' },
  { id:'J007', name:'Javier González Pérez', dob:'03/12/1985', gender:'M', dni:'77889900G', phone:'678901234', email:'javier@email.com', emergency:'Elena González 677777777', position:'Ambas', level:'2ª', ranking:95,  joined:'22/03/2021', status:'Activo',         blood:'B-', medical:'Apto',    notes:'Capitán del equipo A' },
  { id:'J008', name:'Sara Navarro Gil',      dob:'25/08/2001', gender:'F', dni:'88990011H', phone:'689012345', email:'sara@email.com',   emergency:'Luis Navarro 688888888',    position:'Revés', level:'5ª', ranking:550, joined:'14/11/2023', status:'Inactivo',       blood:'A+', medical:'Apto',    notes:'En pausa por estudios' },
];

// ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────
function crearAcademiaPadel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('es_ES');

  // Orden de creación — define el orden de las pestañas
  crearHojaInicio(ss);
  crearHojaPerfil(ss);
  crearHojaAsistenciaPagos(ss);
  crearHojaJugadores(ss);
  crearHojaPagos(ss);
  crearHojaProgreso(ss);
  crearHojaTorneos(ss);
  crearHojaMedico(ss);
  eliminarHojasDefault(ss);

  SpreadsheetApp.flush();
  ss.setActiveSheet(ss.getSheetByName('🏠 Inicio'));

  SpreadsheetApp.getUi().alert(
    '✅ Academia de Pádel creada',
    '8 hojas generadas:\n' +
    '🏠 Inicio · 👤 Perfil · 📅 Asistencia · 📋 Jugadores\n' +
    '💰 Pagos · 📈 Progreso · 🏆 Torneos · 🏥 Médico\n\n' +
    'Los datos de ejemplo son coherentes entre todas las hojas.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ══════════════════════════════════════════════════════════════
// HOJA 1 — INICIO
// ══════════════════════════════════════════════════════════════
function crearHojaInicio(ss) {
  var sh = obtenerHoja(ss, '🏠 Inicio');
  sh.setTabColor(COL.darkBlue);

  // Anchos: A(nombre), B(ver perfil), C(sep), D-I(panel)
  [180, 120, 15, 160, 90, 90, 130, 180, 100].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // Fila 1: Título principal
  fmtRange(sh, 'A1:I1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', valign: 'middle', merge: true });
  sh.getRange('A1').setValue('🎾 ACADEMIA DE PÁDEL');
  sh.setRowHeight(1, 52);

  // Fila 2 vacía
  sh.setRowHeight(2, 10);
  sh.getRange('A2:I2').setBackground(COL.cream);

  // ── Sección izquierda: Lista de Jugadores ──
  fmtRange(sh, 'A3:B3', { bg: COL.medBlue, fc: COL.white, bold: true, size: 13, halign: 'center', merge: true });
  sh.getRange('A3').setValue('👤 JUGADORES');
  sh.setRowHeight(3, 34);

  PLAYERS.forEach(function(p, i) {
    var row = i + 4;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(row, 23);

    sh.getRange(row, 1).setValue(p.name).setBackground(bg).setFontSize(11);

    var link = sh.getRange(row, 2);
    link.setValue('Ver perfil →')
      .setBackground(bg)
      .setFontColor(COL.darkBlue)
      .setFontSize(10)
      .setHorizontalAlignment('center');
  });

  // Borde tarjeta lista
  tarjeta(sh, 'A3:B11', COL.darkBlue);

  // Separador columna C
  sh.getRange('C1:C20').setBackground(COL.cream);

  // ── Sección derecha: Panel de Entrenamiento ──
  fmtRange(sh, 'D3:I3', { bg: COL.medBlue, fc: COL.white, bold: true, size: 13, halign: 'center', merge: true });
  sh.getRange('D3').setValue('📋 PANEL DE ENTRENAMIENTO');
  sh.setRowHeight(3, 34);

  var headers = ['Nombre', 'Nivel FEP', 'Posición', 'Teléfono', 'Email', 'Estado'];
  headers.forEach(function(h, i) {
    fmtCell(sh, 4, i + 4, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 10, halign: 'center' });
    sh.getRange(4, i + 4).setValue(h);
  });
  sh.setRowHeight(4, 28);

  PLAYERS.forEach(function(p, i) {
    var row = i + 5;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(row, 23);
    var vals = [p.name, p.level, p.position, p.phone, p.email, p.status];
    vals.forEach(function(v, j) {
      sh.getRange(row, j + 4).setValue(v).setBackground(bg).setFontSize(10);
    });
  });

  // Formato condicional: Estado (columna I = col 9)
  var estadoRange = sh.getRange('I5:I12');
  aplicarSemaforoEstado(sh, estadoRange);

  // Filtros en fila 4
  sh.getRange('D4:I12').createFilter();

  tarjeta(sh, 'D3:I12', COL.medBlue);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 2 — PERFIL JUGADOR
// Muestra datos del jugador J001 (Carlos García) como ejemplo
// ══════════════════════════════════════════════════════════════
function crearHojaPerfil(ss) {
  var sh = obtenerHoja(ss, '👤 Perfil Jugador');
  sh.setTabColor(COL.medBlue);
  [15, 180, 200, 15, 180, 200, 15, 180].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  // Fondo general crema
  sh.getRange('A1:H50').setBackground(COL.cream);

  // Fila 1: Título
  fmtRange(sh, 'A1:H1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', valign: 'middle', merge: true });
  sh.getRange('A1').setValue('👤 PERFIL DEL JUGADOR');
  sh.setRowHeight(1, 50);
  sh.setRowHeight(2, 12);

  var p = PLAYERS[0]; // J001 Carlos García por defecto

  // Helper para bloque de ficha: pares etiqueta-valor en 2 columnas
  function bloqueInfo(titulo, fila, pares) {
    // Header del bloque
    fmtRange(sh, 'A'+fila+':H'+fila, { bg: COL.medBlue, fc: COL.white, bold: true, size: 12, halign: 'center', merge: true });
    sh.getRange('A'+fila).setValue(titulo);
    sh.setRowHeight(fila, 30);
    // Pares col A-B y E-F (dos columnas de info)
    pares.forEach(function(par, i) {
      var r = fila + 1 + Math.floor(i / 2);
      var colLabel = (i % 2 === 0) ? 2 : 5;
      var colVal   = (i % 2 === 0) ? 3 : 6;
      sh.setRowHeight(r, 24);
      fmtCell(sh, r, colLabel, { bg: COL.lightBlue, bold: true, size: 10, halign: 'right' });
      sh.getRange(r, colLabel).setValue(par[0] + ':');
      fmtCell(sh, r, colVal, { bg: COL.cream, size: 11, halign: 'left' });
      sh.getRange(r, colVal).setValue(par[1]);
    });
  }

  // Cálculo de edad simple
  var hoy = new Date();
  var nac = p.dob.split('/');
  var edad = hoy.getFullYear() - parseInt(nac[2]);

  bloqueInfo('📋 DATOS PERSONALES', 3, [
    ['Nombre', p.name],          ['DNI / RUT', p.dni],
    ['Fecha nacimiento', p.dob], ['Teléfono', p.phone],
    ['Edad', edad + ' años'],    ['Email', p.email],
    ['Género', p.gender === 'M' ? 'Masculino' : 'Femenino'], ['Emergencia', p.emergency],
    ['Fecha ingreso', p.joined], ['Estado', p.status],
    ['Notas', p.notes],          ['ID Jugador', p.id],
  ]);

  sh.setRowHeight(15, 12);

  bloqueInfo('🎾 NIVEL DE JUEGO', 16, [
    ['Nivel FEP', p.level],          ['Ranking FEP', '#' + p.ranking],
    ['Posición', p.position],        ['Última evaluación', '15/01/2025'],
  ]);

  sh.setRowHeight(21, 12);

  bloqueInfo('💰 ESTADO DE PAGOS', 22, [
    ['Último pago', '10/01/2025'],  ['Estado del mes', '✓ Pagado'],
    ['Monto', '80 €'],              ['Método', 'Transferencia'],
  ]);

  sh.setRowHeight(27, 12);

  bloqueInfo('📅 ASISTENCIA', 28, [
    ['Total clases', '20'],          ['% Asistencia', '85%'],
    ['Presencias', '17'],            ['Ausencias', '3'],
  ]);

  sh.setRowHeight(33, 12);

  bloqueInfo('📈 PROGRESO TÉCNICO', 34, [
    ['Nota media', '7.4'],          ['Última evaluación', '15/01/2025'],
    ['Mejor golpe', 'Bandeja'],     ['A mejorar', 'Víbora'],
  ]);

  sh.setRowHeight(39, 12);

  bloqueInfo('🏆 TORNEOS', 40, [
    ['Total torneos', '5'],         ['Puntos FEP', '45'],
    ['Mejor resultado', 'Campeón'], ['Último torneo', 'Open Club Pádel Madrid 2025'],
  ]);

  sh.setRowHeight(45, 12);

  bloqueInfo('🏥 DATOS MÉDICOS', 46, [
    ['Grupo sanguíneo', p.blood],  ['Apto médico', p.medical],
    ['Lesiones actuales', 'Ninguna'], ['Restricciones', 'Ninguna'],
  ]);

  // Nota al pie
  sh.setRowHeight(51, 12);
  var nota = sh.getRange('A52:H52');
  nota.merge().setValue('ℹ️  Para editar datos ir a la hoja 📋 Jugadores')
    .setFontColor(COL.gray).setFontSize(9).setFontStyle('italic')
    .setHorizontalAlignment('center').setBackground(COL.cream);

  // Semáforo en celda de Estado (fila 13, col 3)
  var statusCell = sh.getRange('C13');
  aplicarSemaforoEstado(sh, statusCell);

  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 3 — ASISTENCIA Y PAGOS
// NOTA MIGRACIÓN: tabla ATTENDANCE — FK: player_id → players.id
// ══════════════════════════════════════════════════════════════
function crearHojaAsistenciaPagos(ss) {
  var sh = obtenerHoja(ss, '📅 Asistencia y Pagos');
  sh.setTabColor(COL.green);

  // Expandir hoja a 34 columnas (A hasta AH: 1 nombre + 31 días + % + pago)
  if (sh.getMaxColumns() < 34) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 34 - sh.getMaxColumns());
  }

  // Ancho columna A, B-AF (días, 30px), AG, AH
  sh.setColumnWidth(1, 180);
  for (var d = 2; d <= 32; d++) { sh.setColumnWidth(d, 30); }
  sh.setColumnWidth(33, 95);  // AG: % Asistencia
  sh.setColumnWidth(34, 110); // AH: Pago del mes

  // Fila 1: Título
  fmtRange(sh, 'A1:AH1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('📅 ASISTENCIA Y PAGOS');
  sh.setRowHeight(1, 46);

  // Fila 2: selector de mes
  sh.setRowHeight(2, 28);
  sh.getRange('A2').setValue('MES:').setFontWeight('bold').setFontSize(11).setBackground(COL.cream);
  var meses = ['Enero 2025','Febrero 2025','Marzo 2025','Abril 2025','Mayo 2025','Junio 2025',
               'Julio 2025','Agosto 2025','Septiembre 2025','Octubre 2025','Noviembre 2025','Diciembre 2025',
               'Enero 2026','Febrero 2026','Marzo 2026','Abril 2026','Mayo 2026','Junio 2026',
               'Julio 2026','Agosto 2026','Septiembre 2026','Octubre 2026','Noviembre 2026','Diciembre 2026'];
  sh.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(meses, true).build())
    .setValue('Enero 2025').setBackground(COL.lightBlue).setFontWeight('bold');

  sh.setRowHeight(3, 8);
  sh.getRange('A3:AH3').setBackground(COL.cream);

  // Fila 4: Encabezados
  sh.setRowHeight(4, 30);
  fmtCell(sh, 4, 1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 10, halign: 'center' });
  sh.getRange(4, 1).setValue('Nombre');
  for (var dia = 1; dia <= 31; dia++) {
    fmtCell(sh, 4, dia + 1, { bg: COL.medBlue, fc: COL.white, bold: true, size: 9, halign: 'center' });
    sh.getRange(4, dia + 1).setValue(dia);
  }
  fmtCell(sh, 4, 33, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 9, halign: 'center' });
  sh.getRange(4, 33).setValue('% Asist.');
  fmtCell(sh, 4, 34, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 9, halign: 'center' });
  sh.getRange(4, 34).setValue('Pago mes');

  // Opciones de asistencia y pago
  var optsAsist  = ['✓', '✗', 'J', '—'];
  var optsPago   = ['✓ Pagado', '⏳ Pendiente', '✗ Atrasado'];
  var reglAsist  = SpreadsheetApp.newDataValidation().requireValueInList(optsAsist, true).build();
  var reglaPago  = SpreadsheetApp.newDataValidation().requireValueInList(optsPago, true).build();

  // Ejemplo de asistencia para cada jugador (20 días de clase de 31 posibles)
  var sampleAsist = [
    ['✓','✓','✓','—','✓','✓','—','✓','✓','✓','—','✓','✓','—','✓','✓','✓','—','✓','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','✓','—','✓','✓','✓','—','✓','✓','—','✓','✓','—','✓','✓','✓','—','✓','✓','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','✓','✓','✓','—','✓','✓','✓','—','✓','✓','✓','✓','—','✓','✓','✓','✓','—','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','—','✓','✓','✓','—','✓','✓','✓','—','✓','✓','✓','✓','—','✓','✓','✓','—','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','✓','✗','✗','✗','—','✓','✗','—','✗','✗','—','✓','✓','—','✓','✗','—','✗','✗','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','✓','✓','—','✓','✓','✓','—','✓','✓','—','✓','✓','✓','—','✓','✓','—','✓','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✓','✓','✓','✓','✓','—','✓','✓','✓','✓','—','✓','✓','✓','✓','—','✓','✓','✓','✓','—','—','—','—','—','—','—','—','—','—','—'],
    ['✗','—','✗','✗','—','✗','✗','—','✗','✗','—','✗','✗','—','✗','✗','—','✗','✗','—','—','—','—','—','—','—','—','—','—','—','—'],
  ];
  var samplePago = ['✓ Pagado','✓ Pagado','✓ Pagado','✓ Pagado','⏳ Pendiente','✓ Pagado','✓ Pagado','✗ Atrasado'];

  for (var i = 0; i < PLAYERS.length; i++) {
    var row  = i + 5;
    var bg   = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(row, 22);
    sh.getRange(row, 1).setValue(PLAYERS[i].name).setBackground(bg).setFontSize(10);

    for (var d2 = 0; d2 < 31; d2++) {
      var cell = sh.getRange(row, d2 + 2);
      cell.setValue(sampleAsist[i][d2]).setBackground(bg).setFontSize(9)
        .setHorizontalAlignment('center').setDataValidation(reglAsist);
    }

    // % Asistencia (fórmula: ✓ / (no vacías - "—"))
    var colLetter = 'B';  // días B:AF
    var pctFormula = '=IFERROR(COUNTIF(B'+row+':AF'+row+',"✓")/(COUNTA(B'+row+':AF'+row+')-COUNTIF(B'+row+':AF'+row+',"—")),0)';
    sh.getRange(row, 33).setFormula(pctFormula).setNumberFormat('0%')
      .setBackground(bg).setHorizontalAlignment('center').setFontWeight('bold');

    sh.getRange(row, 34).setValue(samplePago[i]).setBackground(bg)
      .setHorizontalAlignment('center').setDataValidation(reglaPago).setFontSize(10);
  }

  // Formato condicional: celdas de asistencia
  var asistRng = sh.getRange('B5:AF12');
  var reglas = sh.getConditionalFormatRules();
  [['✓', COL.green, COL.white], ['✗', COL.red, COL.white], ['J', COL.yellow, '#1A1A1A'], ['—', COL.lightBlue, COL.gray]].forEach(function(r) {
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setFontColor(r[2]).setRanges([asistRng]).build());
  });
  // Semáforo % asistencia: col AG
  var pctRng = sh.getRange('AG5:AG12');
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8).setBackground(COL.green).setFontColor(COL.white).setRanges([pctRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.6, 0.799).setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([pctRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.6).setBackground(COL.red).setFontColor(COL.white).setRanges([pctRng]).build());
  // Semáforo pago: col AH
  var pagoRng = sh.getRange('AH5:AH12');
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Pagado').setBackground(COL.green).setFontColor(COL.white).setRanges([pagoRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Pendiente').setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([pagoRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Atrasado').setBackground(COL.red).setFontColor(COL.white).setRanges([pagoRng]).build());
  sh.setConditionalFormatRules(reglas);

  sh.setFrozenRows(4);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 4 — JUGADORES (tabla principal PLAYERS)
// NOTA MIGRACIÓN: esta hoja → colección/tabla "players" en BD
//   player_id  = ID (PK)
//   status     = enum: 'active' | 'suspended' | 'inactive'
// ══════════════════════════════════════════════════════════════
function crearHojaJugadores(ss) {
  var sh = obtenerHoja(ss, '📋 Jugadores');
  sh.setTabColor(COL.darkBlue);

  var cols = ['ID','Nombre completo','Fecha nac.','Edad','Categoría','Género','DNI / RUT','Teléfono','Email','Contacto emergencia','Posición','Nivel FEP','Ranking FEP','Fecha ingreso','Estado','Notas'];
  var widths = [55, 185, 90, 45, 90, 60, 100, 100, 160, 180, 70, 75, 85, 90, 100, 200];

  widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  fmtRange(sh, 'A1:P1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('📋 BASE DE DATOS — JUGADORES');
  sh.setRowHeight(1, 44);

  sh.setRowHeight(2, 8);

  cols.forEach(function(h, i) {
    fmtCell(sh, 3, i+1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 10, halign: 'center' });
    sh.getRange(3, i+1).setValue(h);
  });
  sh.setRowHeight(3, 28);

  PLAYERS.forEach(function(p, i) {
    var row = i + 4;
    var bg  = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(row, 23);

    // Fórmula de edad: DATEDIF desde fecha de nacimiento
    // NOTA MIGRACIÓN: en backend, calcular edad en el servidor, no en cliente
    var edadFormula   = '=IFERROR(DATEDIF(C'+row+',TODAY(),"Y"),"")';
    var categoriaFormula = '=IFERROR(IF(D'+row+'<=10,"Benjamín",IF(D'+row+'<=12,"Alevín",IF(D'+row+'<=14,"Infantil",IF(D'+row+'<=16,"Cadete",IF(D'+row+'<=18,"Juvenil","Adulto")))),"")';

    var vals = [p.id, p.name, p.dob, edadFormula, categoriaFormula, p.gender, p.dni, p.phone, p.email, p.emergency, p.position, p.level, p.ranking, p.joined, p.status, p.notes];
    vals.forEach(function(v, j) {
      var cell = sh.getRange(row, j+1);
      if (j === 3 || j === 4) {
        cell.setFormula(v);
      } else {
        cell.setValue(v);
      }
      cell.setBackground(bg).setFontSize(10);
    });

    sh.getRange(row, 1).setFontColor(COL.medBlue).setFontWeight('bold'); // ID destacado
  });

  aplicarSemaforoEstado(sh, sh.getRange('O4:O11'));
  tarjeta(sh, 'A3:P11', COL.darkBlue);
  sh.getRange('A3:P11').createFilter();
  sh.setFrozenRows(3);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 5 — PAGOS (tabla PAYMENTS)
// NOTA MIGRACIÓN: payment_id PK, player_id FK → players.player_id
//   status = enum: 'paid' | 'pending' | 'overdue'
//   type   = enum: 'monthly' | 'single' | 'enrollment' | 'tournament'
// ══════════════════════════════════════════════════════════════
function crearHojaPagos(ss) {
  var sh = obtenerHoja(ss, '💰 Pagos');
  sh.setTabColor(COL.green);

  var cols   = ['ID Pago','ID Jugador','Nombre jugador','Tipo','Mes / Período','Monto (€)','Fecha pago','Método','Estado','Notas'];
  var widths = [70, 70, 185, 120, 110, 80, 90, 110, 100, 200];
  widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  fmtRange(sh, 'A1:J1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('💰 HISTORIAL DE PAGOS');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 8);

  cols.forEach(function(h,i) {
    fmtCell(sh, 3, i+1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 10, halign: 'center' });
    sh.getRange(3, i+1).setValue(h);
  });
  sh.setRowHeight(3, 28);

  // Tipos y métodos como enum — facilita migración a BD
  var tipoOpts   = ['Mensualidad','Clase suelta','Inscripción','Torneo'];
  var metodOpts  = ['Efectivo','Transferencia','Tarjeta','Bizum'];
  var estadoOpts = ['✓ Pagado','⏳ Pendiente','✗ Atrasado'];

  var pagosData = [
    ['P001','J001','Carlos García López','Mensualidad','Enero 2025',80,'10/01/2025','Transferencia','✓ Pagado',''],
    ['P002','J002','María Martínez Ruiz','Mensualidad','Enero 2025',80,'08/01/2025','Bizum','✓ Pagado',''],
    ['P003','J003','Pablo Sánchez Torres','Mensualidad','Enero 2025',80,'05/01/2025','Transferencia','✓ Pagado','Paga siempre puntual'],
    ['P004','J004','Ana López Fernández','Mensualidad','Enero 2025',60,'12/01/2025','Efectivo','✓ Pagado','Tarifa juvenil'],
    ['P005','J005','Diego Romero Vega','Mensualidad','Enero 2025',80,'','','⏳ Pendiente','En baja temporal por lesión'],
    ['P006','J006','Lucía Torres Morales','Mensualidad','Enero 2025',80,'09/01/2025','Tarjeta','✓ Pagado',''],
    ['P007','J007','Javier González Pérez','Mensualidad','Enero 2025',80,'03/01/2025','Transferencia','✓ Pagado',''],
    ['P008','J008','Sara Navarro Gil','Mensualidad','Diciembre 2024',80,'','','✗ Atrasado','No contacta. Pendiente desde diciembre'],
  ];

  pagosData.forEach(function(row, i) {
    var r  = i + 4;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(r, 23);
    row.forEach(function(v, j) {
      sh.getRange(r, j+1).setValue(v).setBackground(bg).setFontSize(10);
    });
    sh.getRange(r, 1).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 2).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 6).setNumberFormat('#,##0 €').setHorizontalAlignment('right');
    sh.getRange(r, 4).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tipoOpts,true).build());
    sh.getRange(r, 8).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(metodOpts,true).build());
    sh.getRange(r, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(estadoOpts,true).build());
  });

  aplicarSemaforoPago(sh, sh.getRange('I4:I11'));
  tarjeta(sh, 'A3:J11', COL.darkBlue);
  sh.getRange('A3:J11').createFilter();

  // ── Bloque resumen ──
  sh.setRowHeight(13, 8);
  var resumenRange = sh.getRange('A14:J16');
  resumenRange.setBackground(COL.lightBlue);
  tarjeta(sh, 'A14:J16', COL.medBlue);

  sh.setRowHeight(14, 30); sh.setRowHeight(15, 28); sh.setRowHeight(16, 28);

  fmtRange(sh, 'A14:J14', { bg: COL.medBlue, fc: COL.white, bold: true, size: 12, halign: 'center', merge: true });
  sh.getRange('A14').setValue('📊 RESUMEN DEL MES');

  var labels = [['Total recaudado','=SUMIF(I4:I11,"✓ Pagado",F4:F11)'], ['Total pendiente','=SUMIF(I4:I11,"⏳ Pendiente",F4:F11)+SUMIF(I4:I11,"✗ Atrasado",F4:F11)'], ['Alumnos al día','=COUNTIF(I4:I11,"✓ Pagado")&" / "&COUNTA(I4:I11)']];
  labels.forEach(function(l, i) {
    var col = i * 3 + 1;
    sh.getRange(15, col).setValue(l[0]).setFontWeight('bold').setFontSize(10).setBackground(COL.lightBlue);
    sh.getRange(16, col).setFormula(l[1]).setFontSize(12).setFontWeight('bold').setFontColor(COL.darkBlue)
      .setHorizontalAlignment('center').setBackground(COL.cream);
    if (i < 2) sh.getRange(16, col).setNumberFormat('#,##0 €');
  });

  sh.setFrozenRows(3);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 6 — PROGRESO TÉCNICO (tabla EVALUATIONS)
// NOTA MIGRACIÓN: evaluation_id PK, player_id FK
//   Cada golpe (1-10) = campo numérico en BD
// ══════════════════════════════════════════════════════════════
function crearHojaProgreso(ss) {
  var sh = obtenerHoja(ss, '📈 Progreso Técnico');
  sh.setTabColor('#7c3aed');

  var cols = ['ID Eval','ID Jugador','Nombre jugador','Fecha','Nivel FEP','Saque','Volea','Globo','Bajada','Bandeja','Víbora','Remate','Táctica','Nota media','Objetivos','Observaciones'];
  var wids = [65, 70, 185, 85, 70, 50,50,50,60,60,55,60,60, 80, 200, 200];
  wids.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  fmtRange(sh, 'A1:P1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('📈 PROGRESO TÉCNICO');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 8);

  cols.forEach(function(h,i) {
    fmtCell(sh, 3, i+1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 9, halign: 'center' });
    sh.getRange(3, i+1).setValue(h);
  });
  sh.setRowHeight(3, 28);

  var evalData = [
    ['E001','J001','Carlos García López','15/01/2025','3ª',7,8,6,7,9,5,7,8,'','Mejorar víbora','Buen nivel general'],
    ['E002','J002','María Martínez Ruiz','16/01/2025','4ª',6,7,7,6,7,6,5,7,'','Potenciar el remate','Progresa bien'],
    ['E003','J003','Pablo Sánchez Torres','17/01/2025','2ª',9,8,7,8,8,8,9,9,'','Mantener nivel competitivo','Excelente evolución'],
    ['E004','J004','Ana López Fernández','18/01/2025','5ª',5,6,6,5,6,4,5,6,'','Trabajar golpe de fondo','Mucho potencial'],
    ['E005','J005','Diego Romero Vega','20/01/2025','3ª',7,6,5,7,7,6,6,7,'','Recuperación y volea','Parado por lesión'],
    ['E006','J006','Lucía Torres Morales','21/01/2025','4ª',6,7,7,6,9,5,6,7,'','Víbora y saque','Bandeja excelente'],
    ['E007','J007','Javier González Pérez','22/01/2025','2ª',9,9,8,9,9,8,9,9,'','Planificación táctica','Nivel elite local'],
    ['E008','J008','Sara Navarro Gil','23/01/2025','5ª',4,5,5,4,5,4,4,5,'','Todos los golpes básicos','Falta continuidad'],
  ];

  evalData.forEach(function(row, i) {
    var r  = i + 4;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(r, 23);
    var notaFormula = '=ROUND(AVERAGE(F'+r+':M'+r+'),1)';
    row.forEach(function(v, j) {
      sh.getRange(r, j+1).setValue(v).setBackground(bg).setFontSize(10);
    });
    // Nota media calculada (col 14)
    sh.getRange(r, 14).setFormula(notaFormula).setFontWeight('bold').setBackground(bg).setFontSize(11);
    sh.getRange(r, 1).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 2).setFontColor(COL.medBlue).setFontWeight('bold');
    // Scores centrados
    for (var s = 6; s <= 13; s++) { sh.getRange(r, s).setHorizontalAlignment('center'); }
  });

  // Semáforo nota media
  var notaRng = sh.getRange('N4:N11');
  var reglas = sh.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground(COL.green).setFontColor(COL.white).setRanges([notaRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5, 6.99).setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([notaRng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground(COL.red).setFontColor(COL.white).setRanges([notaRng]).build());
  sh.setConditionalFormatRules(reglas);

  tarjeta(sh, 'A3:P11', COL.darkBlue);
  sh.getRange('A3:P11').createFilter();
  sh.setFrozenRows(3);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 7 — TORNEOS (tabla TOURNAMENTS)
// NOTA MIGRACIÓN: tournament_id PK, player_id FK
//   phase = enum: 'champion'|'finalist'|'semi'|'quarter'|'round16'|'round1'
// ══════════════════════════════════════════════════════════════
function crearHojaTorneos(ss) {
  var sh = obtenerHoja(ss, '🏆 Torneos');
  sh.setTabColor(COL.yellow);

  var cols = ['ID','ID Jugador','Nombre jugador','Nombre torneo','Fecha','Club / Sede','Categoría','Pareja','Fase alcanzada','Posición','Pts FEP','Notas'];
  var wids = [55,70,185,200,85,160,70,185,120,70,65,200];
  wids.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  fmtRange(sh, 'A1:L1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('🏆 TORNEOS');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 8);

  cols.forEach(function(h,i) {
    fmtCell(sh, 3, i+1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 10, halign: 'center' });
    sh.getRange(3, i+1).setValue(h);
  });
  sh.setRowHeight(3, 28);

  var faseOpts = ['Campeón','Finalista','Semifinales','Cuartos de final','Octavos de final','Primera ronda'];
  var catOpts  = ['1ª','2ª','3ª','4ª','5ª','6ª','7ª','Open'];
  var reglaFase = SpreadsheetApp.newDataValidation().requireValueInList(faseOpts,true).build();
  var reglaCat  = SpreadsheetApp.newDataValidation().requireValueInList(catOpts,true).build();

  var torneosData = [
    ['T001','J001','Carlos García López','Open Club Pádel Madrid 2025','15/01/2025','Club Pádel Madrid','3ª','Pablo Sánchez Torres','Campeón','1º',25,'Gran actuación en final'],
    ['T002','J003','Pablo Sánchez Torres','Open Club Pádel Madrid 2025','15/01/2025','Club Pádel Madrid','3ª','Carlos García López','Campeón','1º',25,'Pareja con J001'],
    ['T003','J007','Javier González Pérez','Torneo Regional Pádel Sur','08/01/2025','Club Sur','2ª','Jugador externo','Finalista','2º',15,'Perdió final por 7-6'],
    ['T004','J002','María Martínez Ruiz','Copa Femenina Enero 2025','20/01/2025','Centro Deportivo Norte','4ª','Lucía Torres Morales','Semifinales','3º',8,'Buen torneo para el nivel'],
    ['T005','J006','Lucía Torres Morales','Copa Femenina Enero 2025','20/01/2025','Centro Deportivo Norte','4ª','María Martínez Ruiz','Semifinales','3º',8,'Pareja con J002'],
    ['T006','J004','Ana López Fernández','Circuito Juvenil Enero 2025','22/01/2025','Club Juvenil','5ª','Compañera externo','Cuartos de final','5º',5,'Primer torneo oficial'],
    ['T007','J005','Diego Romero Vega','Liga Interna Academia','10/12/2024','Academia Pádel','3ª','Carlos García López','Octavos de final','8º',2,'Antes de la lesión'],
  ];

  torneosData.forEach(function(row, i) {
    var r  = i + 4;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(r, 23);
    row.forEach(function(v, j) {
      sh.getRange(r, j+1).setValue(v).setBackground(bg).setFontSize(10);
    });
    sh.getRange(r, 1).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 2).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 9).setDataValidation(reglaFase);
    sh.getRange(r, 7).setDataValidation(reglaCat);
  });

  // Semáforo fase alcanzada
  var faseRng = sh.getRange('I4:I10');
  var reglas2 = sh.getConditionalFormatRules();
  [['Campeón',COL.green,COL.white],['Finalista',COL.green,COL.white],['Semifinales',COL.yellow,'#1A1A1A'],['Cuartos de final',COL.yellow,'#1A1A1A'],['Octavos de final',COL.red,COL.white],['Primera ronda',COL.red,COL.white]].forEach(function(r) {
    reglas2.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setFontColor(r[2]).setRanges([faseRng]).build());
  });
  sh.setConditionalFormatRules(reglas2);

  tarjeta(sh, 'A3:L10', COL.darkBlue);
  sh.getRange('A3:L10').createFilter();
  sh.setFrozenRows(3);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// HOJA 8 — MÉDICO (tabla MEDICAL)
// NOTA MIGRACIÓN: tabla "medical_records", player_id FK (1:1 con players)
//   apto_medico = enum: 'fit' | 'not_fit' | 'conditional'
//   IMPORTANTE: datos sensibles — en BD aplicar cifrado en reposo
// ══════════════════════════════════════════════════════════════
function crearHojaMedico(ss) {
  var sh = obtenerHoja(ss, '🏥 Médico');
  sh.setTabColor(COL.red);

  var cols = ['ID Jugador','Nombre jugador','Grupo sang.','Alergias','Lesiones previas','Lesiones actuales','Medicación','Médico cabecera','Seguro médico','Apto médico','Última revisión','Restricciones físicas','Notas'];
  var wids = [70,185,70,120,180,160,120,160,120,90,100,180,200];
  wids.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });

  fmtRange(sh, 'A1:M1', { bg: COL.darkBlue, fc: COL.white, bold: true, size: 16, halign: 'center', merge: true });
  sh.getRange('A1').setValue('🏥 DATOS MÉDICOS');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 8);

  cols.forEach(function(h,i) {
    fmtCell(sh, 3, i+1, { bg: COL.darkBlue, fc: COL.white, bold: true, size: 9, halign: 'center' });
    sh.getRange(3, i+1).setValue(h);
  });
  sh.setRowHeight(3, 28);

  var aptoOpts = ['Apto','No apto','Condicional'];
  var medData = [
    ['J001','Carlos García López','A+','Ninguna','Esguince tobillo derecho (2021)','Ninguna','Ninguna','Dr. Pérez García','AXA Salud','Apto','15/01/2025','Ninguna',''],
    ['J002','María Martínez Ruiz','O+','Polen','Tendinitis hombro (2022)','Ninguna','Ninguna','Dra. López Ruiz','Sanitas','Apto','10/01/2025','Ninguna','Alérgica en primavera'],
    ['J003','Pablo Sánchez Torres','B+','Ninguna','Rotura fibrilar gemelo (2019)','Ninguna','Ibuprofeno ocasional','Dr. García Sánchez','Adeslas','Apto','12/01/2025','Ninguna',''],
    ['J004','Ana López Fernández','A-','Látex','Ninguna','Ninguna','Ninguna','Dr. Fernández López','Asisa','Apto','20/01/2025','Ninguna','Menor de edad — autorización parental'],
    ['J005','Diego Romero Vega','AB+','Ninguna','Menisco (operado 2018)','Lesión LCA rodilla izquierda','Antiinflamatorios','Dr. Romero Vega','Mutua Madrileña','No apto','08/01/2025','Sin actividad de alta intensidad','Seguimiento semanal'],
    ['J006','Lucía Torres Morales','O-','Penicilina','Ninguna','Ninguna','Ninguna','Dra. Torres García','AXA Salud','Apto','18/01/2025','Ninguna','Alérgica a penicilina — IMPORTANTE'],
    ['J007','Javier González Pérez','B-','Ninguna','Codo de tenista (2017), Rotura parcial manguito (2020)','Ninguna','Ninguna','Dr. González Pérez','Sanitas','Apto','05/01/2025','Calentamiento extendido','Historial de lesiones en hombro'],
    ['J008','Sara Navarro Gil','A+','Ninguna','Ninguna','Ninguna','Ninguna','Dra. Navarro Gil','Sin seguro privado','Apto','14/11/2023','Ninguna','Pendiente de renovar revisión'],
  ];

  medData.forEach(function(row, i) {
    var r  = i + 4;
    var bg = i % 2 === 0 ? COL.lightBlue : COL.cream;
    sh.setRowHeight(r, 23);
    row.forEach(function(v, j) {
      sh.getRange(r, j+1).setValue(v).setBackground(bg).setFontSize(10);
    });
    sh.getRange(r, 1).setFontColor(COL.medBlue).setFontWeight('bold');
    sh.getRange(r, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(aptoOpts,true).build())
      .setHorizontalAlignment('center').setFontWeight('bold');
  });

  // Semáforo Apto médico
  var aptoRng = sh.getRange('J4:J11');
  var reglas3 = sh.getConditionalFormatRules();
  reglas3.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Apto').setBackground(COL.green).setFontColor(COL.white).setRanges([aptoRng]).build());
  reglas3.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Condicional').setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([aptoRng]).build());
  reglas3.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No apto').setBackground(COL.red).setFontColor(COL.white).setRanges([aptoRng]).build());
  sh.setConditionalFormatRules(reglas3);

  tarjeta(sh, 'A3:M11', COL.darkBlue);
  sh.getRange('A3:M11').createFilter();
  sh.setFrozenRows(3);
  sh.setHiddenGridlines(true);
}

// ══════════════════════════════════════════════════════════════
// FUNCIONES AUXILIARES
// ══════════════════════════════════════════════════════════════

// Obtiene hoja (borrando si existe) y oculta líneas de grilla
function obtenerHoja(ss, nombre) {
  var h = ss.getSheetByName(nombre);
  if (h) ss.deleteSheet(h);
  var sh = ss.insertSheet(nombre);
  sh.setHiddenGridlines(true);
  return sh;
}

// Formatea un rango con opciones simplificadas
// opts: { bg, fc, bold, size, halign, valign, merge }
function fmtRange(sh, a1, opts) {
  var r = sh.getRange(a1);
  if (opts.merge)  r.merge();
  if (opts.bg)     r.setBackground(opts.bg);
  if (opts.fc)     r.setFontColor(opts.fc);
  if (opts.bold)   r.setFontWeight(opts.bold ? 'bold' : 'normal');
  if (opts.size)   r.setFontSize(opts.size);
  if (opts.halign) r.setHorizontalAlignment(opts.halign);
  if (opts.valign) r.setVerticalAlignment(opts.valign);
  r.setFontFamily('Arial');
  return r;
}

// Formatea una celda por fila/columna (sin merge)
function fmtCell(sh, row, col, opts) {
  var r = sh.getRange(row, col);
  if (opts.bg)     r.setBackground(opts.bg);
  if (opts.fc)     r.setFontColor(opts.fc);
  if (opts.bold)   r.setFontWeight(opts.bold ? 'bold' : 'normal');
  if (opts.size)   r.setFontSize(opts.size);
  if (opts.halign) r.setHorizontalAlignment(opts.halign);
  r.setFontFamily('Arial');
  return r;
}

// Borde tipo tarjeta alrededor de un rango
function tarjeta(sh, a1, color) {
  sh.getRange(a1).setBorder(true, true, true, true, false, false, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Semáforo de estado de jugador: Activo / Baja temporal / Inactivo
function aplicarSemaforoEstado(sh, rng) {
  var reglas = sh.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Activo').setBackground(COL.green).setFontColor(COL.white).setRanges([rng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Baja temporal').setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([rng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Inactivo').setBackground(COL.red).setFontColor(COL.white).setRanges([rng]).build());
  sh.setConditionalFormatRules(reglas);
}

// Semáforo de estado de pago
function aplicarSemaforoPago(sh, rng) {
  var reglas = sh.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Pagado').setBackground(COL.green).setFontColor(COL.white).setRanges([rng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Pendiente').setBackground(COL.yellow).setFontColor('#1A1A1A').setRanges([rng]).build());
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Atrasado').setBackground(COL.red).setFontColor(COL.white).setRanges([rng]).build());
  sh.setConditionalFormatRules(reglas);
}

// Borra las hojas por defecto que Google Sheets crea automáticamente
function eliminarHojasDefault(ss) {
  ['Hoja 1','Hoja1','Sheet1','Sheet 1'].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
}

// ══════════════════════════════════════════════════════════════
// FASE 2 — ROADMAP DE MIGRACIÓN A APP
//
// Cuando este sistema se convierta en aplicación móvil/web,
// cada hoja de Google Sheets se mapea a una colección/tabla:
//
//  Hoja "Jugadores"         → tabla  players
//    columns: player_id (PK), name, dob, gender, dni, phone,
//             email, emergency_contact, position, level_fep,
//             ranking_fep, joined_date, status, notes
//
//  Hoja "Asistencia y Pagos" → tabla  attendance
//    columns: attendance_id (PK), player_id (FK), date,
//             status ENUM('present','absent','justified','na'),
//             month_payment_status ENUM('paid','pending','overdue')
//
//  Hoja "Pagos"             → tabla  payments
//    columns: payment_id (PK), player_id (FK), type ENUM(...),
//             period, amount, payment_date, method, status, notes
//
//  Hoja "Progreso Técnico"  → tabla  evaluations
//    columns: evaluation_id (PK), player_id (FK), eval_date,
//             level_fep, serve, volley, lob, wall_drop, bandeja,
//             vibora, smash, tactics, avg_score, objectives, notes
//
//  Hoja "Torneos"           → tabla  tournaments
//    columns: tournament_id (PK), player_id (FK), name, date,
//             venue, category, partner, phase ENUM(...),
//             position, fep_points, notes
//
//  Hoja "Médico"            → tabla  medical_records
//    columns: record_id (PK), player_id (FK, UNIQUE), blood_type,
//             allergies, past_injuries, current_injuries,
//             medication, doctor, insurance, fit_status ENUM(...),
//             last_checkup, restrictions, notes
//    ⚠️ CIFRADO EN REPOSO obligatorio (datos sensibles)
//
// RECOMENDACIONES TÉCNICAS:
//   - Backend sugerido: Supabase (PostgreSQL) o Firebase Firestore
//   - Auth: Firebase Auth con roles coach / player / admin
//   - API: REST o GraphQL (hasura sobre Supabase)
//   - Mobile: React Native / Flutter
//   - El campo "ID" de cada tabla (J001, P001…) debe reemplazarse
//     por UUID autogenerado en la BD real
// ══════════════════════════════════════════════════════════════

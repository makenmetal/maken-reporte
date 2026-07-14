// ============================================================
// MAKEN METAL - Apps Script PRODUCCION v11.31
// Este script vive en el Google Sheet de Reporte de Producción
// v11.31 (Jul 2026):
//   - procesarReportesEficiencia lee columna "Tipo Manto" del sheet BD
//     (Correctivo/Preventivo/Predictivo con posible subtipo separado por " - ").
//   - actsMantoSetup incluye tipoManto en cada reporte de manto.
//   - El frontend filtra solo reportes de manto en el modal (sin montajes)
//     y muestra la clasificación como badge de color por tipo.
// v11.30 (previo): Fix rango exclusivo (>=)
//   - tiemposCicloPorParteOp filtra últimos 180 días (6 meses)
//   - Retorna mediana, desviación estándar muestral (n-1),
//     coeficiente de variación (%) e interpretación de estabilidad
//     (Muy estable/Estable/Variable/Inestable) cuando n≥3
// v11.14 (previo): Módulo de Ingeniería — Tiempo de Ciclo
//   - getEmpleados incluye "esIngenieria" desde columna "Ingenieria"
//   - Nueva pestaña TiemposCiclo (autocreada) con 9 columnas
//   - doPost soporta tipo="TiempoCiclo" para guardar mediciones
//   - Endpoint "tiemposCicloPorParteOp": histórico + promedio + variación
//     vs catálogo para una parte + operación
//   - Endpoint "buscarPartes": búsqueda parcial por número de parte
// v11.13 (previo): Botón "Ayer" + rangoLabel
//   - Reemplaza "detallePiezaOperador" con "detalleParteOperacion"
//     que analiza una parte + operación de forma GLOBAL:
//     no filtra por máquina/turno/operador, solo respeta rango de fechas
//     (con periodo propio, independiente del dashboard)
//   - Retorna: operaciones disponibles del catálogo, tendencia diaria
//     agregada, y tabla consolidada operador+máquina
// v11.9 (previo): Performance con CacheService de catálogos
// v11.8 (previo): Desglose por turno en modal
//   - Regla 1: Máquina con manto ≥ 90% del total → eff = null (N/A)
//     Rationale: no hubo tiempo suficiente para operar
//   - Regla 2: Máquinas seleccionadas sin actividad → eff = 0
//     Ahora SÍ penalizan el promedio global (antes eran ignoradas)
//   - Regla 3: Máquinas con eff = null (N/A) → EXCLUIDAS del promedio global
//   - Cambio de cálculo global: era sum(hProd)/sum(hDisp), ahora es
//     promedio simple de eficiencias individuales excluyendo N/A
//   - Response incluye numMaquinasEnManto y numMaquinasContadas
// v11.5 (previo):
//   - FIX crítico: parseFechaDashboard detecta Date objects (columna con
//     formato fecha en el sheet) y normaliza a 6 AM. "Hoy" ya no sale vacío.
// v11.4 (previo):
//   - Fórmula: eficiencia excluye manto del denominador
//   - Tarjetas vacías para máquinas seleccionadas sin actividad
// v11.3 (previo):
//   - FIX: Se reconocen "Manto" y "Mantenimiento"
// v11.2 (previo):
//   - Filtro por selección múltiple de máquinas
// v11.1 (previo):
//   - dashboardEficiencia con hManto
//   - Filtro por turno
// v11 (previo):
//   - Nuevo endpoint "dashboardEficiencia"
// v10 (previo):
//   - "registrosOperador", "maquinas" endpoint dinámico
// ============================================================

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  if (action === "miturno") {
    result = getMiTurno(ss, e.parameter.fecha, e.parameter.turno, e.parameter.empleado);
  } else if (action === "actualizar") {
    result = actualizarReporte(ss, e.parameter);
  } else if (action === "borrar") {
    result = borrarReporte(ss, e.parameter);
  } else if (action === "reportes") {
    result = getReportes(ss, e.parameter.fecha, e.parameter.turno);
  } else if (action === "empleados") {
    result = getEmpleados(ss);
  } else if (action === "operaciones") {
    result = getOperaciones(ss);
  } else if (action === "partes") {
    result = getPartes(ss);
  } else if (action === "tiposmanto") {
    result = getTiposManto(ss);
  } else if (action === "maquinas") {
    result = getMaquinas(ss);
  } else if (action === "registrosOperador") {
    result = registrosOperador(ss, e.parameter.fecha, e.parameter.turno, e.parameter.empleado);
  } else if (action === "dashboardProd") {
    result = dashboardProduccion(ss, e.parameter.periodo);
  } else if (action === "dashboardEficiencia") {
    result = dashboardEficiencia(ss, e.parameter.periodo, e.parameter.area, e.parameter.desde, e.parameter.hasta, e.parameter.turno, e.parameter.maquinas);
  } else if (action === "dashboardEfMaquina") {
    result = dashboardEfMaquina(ss, e.parameter.maquina, e.parameter.periodo, e.parameter.desde, e.parameter.hasta, e.parameter.turno);
  } else if (action === "detalleParteOperacion") {
    result = detalleParteOperacion(ss, e.parameter.noParte, e.parameter.operacion, e.parameter.periodo, e.parameter.desde, e.parameter.hasta);
  } else if (action === "tiemposCicloPorParteOp") {
    result = tiemposCicloPorParteOp(ss, e.parameter.noParte, e.parameter.operacion);
  } else if (action === "buscarPartes") {
    result = buscarPartes(ss, e.parameter.query);
  } else if (action === "limpiarCache") {
    // v11.20: Limpia todos los cachés del sheet
    try {
      var cache = CacheService.getScriptCache();
      cache.removeAll(['maq_v1', 'partes_v1', 'emp_v1', 'tc_medidos_v1']);
      result = { ok: true, mensaje: "Caché limpiado — refresca el dashboard" };
    } catch (e) {
      result = { ok: false, error: String(e) };
    }
  } else if (action === "diagnostico") {
    // v11.20: Diagnóstico completo de lectura de reportes
    result = diagnosticoLectura(ss, e.parameter.periodo, e.parameter.desde, e.parameter.hasta);
  } else if (action === "reporteMediciones") {
    // v11.21: Reporte diario de mediciones TC vs producción
    result = reporteMediciones(ss, e.parameter.periodo, e.parameter.desde, e.parameter.hasta);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    // v11.14: Nuevo tipo de captura — TiempoCiclo (medición de ingeniería)
    if (data.tipo === "TiempoCiclo") {
      var sheet = ensureTiemposCicloSheet(ss);
      sheet.appendRow([
        data.fecha, data.turno, data.ingeniero, data.maquina,
        data.noParte, data.operacion,
        data.tiempoCicloReal,     // en minutos con decimales
        data.observaciones || "",
        new Date()
      ]);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // Comportamiento por default: reportes normales
    var sheet = ss.getSheetByName("Reportes") || ss.getActiveSheet();
    sheet.appendRow([
      data.fecha, data.turno, data.tipo, data.maquina, data.responsable,
      data.noParte || "", data.ordenTrabajo || "", data.tipoOperacion || "", data.operacion || "",
      data.piezas || "", data.horasSetup || "", data.horasManto || "",
      data.tipoManto || "", data.descripcion || "",
      new Date()
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// v11.14: Asegura que exista la pestaña TiemposCiclo con encabezados correctos
function ensureTiemposCicloSheet(ss) {
  var sh = ss.getSheetByName("TiemposCiclo");
  if (!sh) {
    sh = ss.insertSheet("TiemposCiclo");
    sh.appendRow([
      "Fecha", "Turno", "Ingeniero", "Máquina",
      "NoParte", "Operacion", "TiempoCicloReal", "Observaciones", "Timestamp"
    ]);
    // Formato de encabezados
    var head = sh.getRange(1, 1, 1, 9);
    head.setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ============================================================
// HELPERS
// ============================================================

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) row[headers[j]] = data[i][j];
    row._rowIndex = i + 1;
    rows.push(row);
  }
  return { headers: headers, rows: rows };
}

function colMap(headers) {
  var m = {};
  for (var k = 0; k < headers.length; k++) m[headers[k]] = k + 1;
  return m;
}

// ============================================================
// REPORTES (ya existían, sin cambios)
// ============================================================

function getMiTurno(ss, fecha, turno, empleado) {
  var sheet = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var r = sheetToObjects(sheet);
  var rows = [];
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (String(row["Fecha"]).trim() !== fecha) continue;
    if (String(row["Turno"]).trim() !== turno) continue;
    if (String(row["Responsable"]).indexOf(empleado) < 0) continue;
    rows.push({
      timestamp:    String(row["Timestamp"] || ""),
      tipo:         row["Tipo"],
      maquina:      row["Máquina"],
      noParte:      row["No. Parte"],
      ordenTrabajo: row["Orden Trabajo"],
      tipoOperacion:row["Tipo Operación"],
      operacion:    row["Operación"],
      piezas:       row["Piezas"],
      horasSetup:   row["Horas Setup"],
      horasManto:   row["Horas Manto"],
      tipoManto:    row["Tipo Manto"],
      descripcion:  row["Descripción"]
    });
  }
  return rows;
}

// NUEVO: dueño ve registros de cualquier operador
function registrosOperador(ss, fecha, turno, empleado) {
  return getMiTurno(ss, fecha, turno, empleado);
}

function actualizarReporte(ss, p) {
  var sheet = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var r = sheetToObjects(sheet);
  var cm = colMap(r.headers);
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (String(row["Timestamp"]).trim() !== p.timestamp) continue;
    var esDueno = p.empleado === "9473" || p.dueno === "1";
    if (!esDueno && String(row["Responsable"]).indexOf(p.empleado) < 0) continue;

    var updates = {
      "Máquina":       p.maquina       || "",
      "No. Parte":     p.noParte       || "",
      "Orden Trabajo": p.ordenTrabajo  || "",
      "Tipo Operación":p.tipoOperacion || "",
      "Operación":     p.operacion     || "",
      "Piezas":        p.piezas        || "",
      "Horas Setup":   p.horasSetup    || "",
      "Horas Manto":   p.horasManto    || "",
      "Tipo Manto":    p.tipoManto     || "",
      "Descripción":   p.descripcion   || ""
    };
    for (var col in updates) {
      if (cm[col]) sheet.getRange(row._rowIndex, cm[col]).setValue(updates[col]);
    }
    if (cm["Modificado"]) sheet.getRange(row._rowIndex, cm["Modificado"]).setValue(new Date());
    return { ok: true };
  }
  return { ok: false, error: "Registro no encontrado o no autorizado" };
}

function borrarReporte(ss, p) {
  var sheet = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var r = sheetToObjects(sheet);
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (String(row["Timestamp"]).trim() !== p.timestamp) continue;
    var esDueno = p.empleado === "9473" || p.dueno === "1";
    if (!esDueno && String(row["Responsable"]).indexOf(p.empleado) < 0) continue;

    var histSheet = ss.getSheetByName("Historial_Borrados");
    if (!histSheet) {
      histSheet = ss.insertSheet("Historial_Borrados");
      var headersHist = r.headers.slice();
      headersHist.push("Fecha Borrado");
      headersHist.push("Borrado Por");
      histSheet.appendRow(headersHist);
      histSheet.getRange(1, 1, 1, headersHist.length).setFontWeight("bold");
    }
    var filaDatos = sheet.getRange(row._rowIndex, 1, 1, r.headers.length).getDisplayValues()[0];
    var filaCopia = filaDatos.slice();
    filaCopia.push(new Date());
    filaCopia.push(p.empleado);
    histSheet.appendRow(filaCopia);
    sheet.deleteRow(row._rowIndex);
    return { ok: true };
  }
  return { ok: false, error: "Registro no encontrado o no autorizado" };
}

function getReportes(ss, fecha, turno) {
  var sheet = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var r = sheetToObjects(sheet);
  var rows = [];
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (String(row["Fecha"]).trim() !== fecha) continue;
    if (String(row["Turno"]).trim() !== turno) continue;
    rows.push({
      timestamp:    String(row["Timestamp"] || ""),
      tipo:         row["Tipo"],
      maquina:      row["Máquina"],
      responsable:  row["Responsable"],
      noParte:      row["No. Parte"],
      ordenTrabajo: row["Orden Trabajo"],
      tipoOperacion:row["Tipo Operación"],
      operacion:    row["Operación"],
      piezas:       row["Piezas"],
      horasSetup:   row["Horas Setup"],
      horasManto:   row["Horas Manto"],
      tipoManto:    row["Tipo Manto"],
      descripcion:  row["Descripción"]
    });
  }
  return rows;
}

// ============================================================
// CATALOGOS
// ============================================================

function getEmpleados(ss) {
  var data = ss.getSheetByName("Empleados").getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  for (var k = 0; k < headers.length; k++) idx[headers[k]] = k;
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      numero: data[i][0],
      nombre: data[i][1],
      puedeSolicitar: idx["Puede_Solicitar"] >= 0 ? String(data[i][idx["Puede_Solicitar"]]).toLowerCase().indexOf("s") === 0 : false,
      puedeAceptar:   idx["Puede_Aceptar"]   >= 0 ? String(data[i][idx["Puede_Aceptar"]]).toLowerCase().indexOf("s") === 0 : false,
      esManto:        idx["Es_Manto"]        >= 0 ? String(data[i][idx["Es_Manto"]]).toLowerCase().indexOf("s") === 0 : false,
      puedeReportarManto: idx["Puede_Reportar_Manto"] >= 0 ? String(data[i][idx["Puede_Reportar_Manto"]]).toLowerCase().indexOf("s") === 0 : false,
      esIngenieria:   idx["Ingenieria"]      >= 0 ? String(data[i][idx["Ingenieria"]]).toLowerCase().indexOf("s") === 0 : false
    });
  }
  return rows;
}

function getOperaciones(ss) {
  var data = ss.getSheetByName("Operaciones").getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) rows.push({ codigo: String(data[i][0]), nombre: data[i][1] });
  }
  return rows;
}

function getPartes(ss) {
  var data = ss.getSheetByName("Partes").getDataRange().getValues();
  var obj = {};
  for (var i = 1; i < data.length; i++) {
    var parte = data[i][0], op = data[i][1], tc = data[i][2];
    if (!parte || !op) continue;
    if (!obj[parte]) obj[parte] = {};
    obj[parte][op] = tc;
  }
  return obj;
}

function getTiposManto(ss) {
  var sh = ss.getSheetByName("TiposManto");
  if (!sh) return { principales: ["Preventivo","Correctivo","Predictivo","Emergencia"], sub: {} };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { principales: [], sub: {} };
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var hasCategoria = headers.indexOf("Categoria") >= 0;
  var idxCat = headers.indexOf("Categoria");
  var principales = [];
  var sub = {};
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][0] || "").trim();
    if (!nombre) continue;
    var cat = hasCategoria ? String(data[i][idxCat] || "").trim() : "";
    if (!cat) {
      principales.push(nombre);
    } else {
      if (!sub[cat]) sub[cat] = [];
      sub[cat].push(nombre);
    }
  }
  return { principales: principales, sub: sub };
}

// NUEVO: catálogo Maquinas dinámico con fallback
function getMaquinas(ss) {
  var sh = ss.getSheetByName("Maquinas");
  if (!sh) {
    // Fallback: catálogo hardcoded (retrocompatibilidad)
    return [
      { maquina: "Puma 1", area: "Torno", tipo: "Principal" },
      { maquina: "Puma 2", area: "Torno", tipo: "Principal" },
      { maquina: "Haas 1", area: "Torno", tipo: "Principal" },
      { maquina: "Haas 2", area: "Torno", tipo: "Principal" },
      { maquina: "Doosan 2", area: "Centro de Maquinado", tipo: "Principal" },
      { maquina: "Doosan 3", area: "Centro de Maquinado", tipo: "Principal" },
      { maquina: "Amada", area: "Corte", tipo: "Principal" },
      { maquina: "BSA", area: "Rectificado Exterior", tipo: "Principal" },
      { maquina: "Kellen", area: "Rectificado Exterior", tipo: "Principal" },
      { maquina: "Okamoto", area: "Rectificado Interior", tipo: "Principal" },
      { maquina: "Ensamble 1", area: "Ensamble", tipo: "Principal" },
      { maquina: "Soldadura 1", area: "Soldadura", tipo: "Principal" },
      { maquina: "Soldadura 2", area: "Soldadura", tipo: "Principal" },
      { maquina: "Detallado ZP", area: "Detallado", tipo: "Principal" },
      { maquina: "Detallado Final", area: "Detallado", tipo: "Principal" },
      { maquina: "Revolver", area: "Avellanado", tipo: "Principal" },
      { maquina: "Machueleadora", area: "Roscado", tipo: "Principal" }
    ];
  }
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  for (var k = 0; k < headers.length; k++) idx[headers[k]] = k;
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][idx["Máquina"]] || "").trim();
    if (!nombre) continue;
    var activa = idx["Activa"] >= 0 ? String(data[i][idx["Activa"]]).toLowerCase().indexOf("s") === 0 : true;
    if (!activa) continue;
    rows.push({
      maquina: nombre,
      area:    String(data[i][idx["Área"]] || data[i][idx["Area"]] || "").trim(),
      tipo:    String(data[i][idx["Tipo"]] || "Principal").trim()
    });
  }
  return rows;
}

// ============================================================
// DASHBOARD PRODUCCION
// ============================================================

// Devuelve {desde:Date, hasta:Date} aplicando regla de inicio de día a las 6 AM
function getRangoDashboard(periodo) {
  var now = new Date();
  var hasta = now;
  var desde;
  // v11.12: El día laboral inicia a las 7 AM. Si son menos de las 7 AM,
  // el "día laboral" es el día anterior (asume 3er turno todavía activo).
  var refDate = new Date(now);
  if (refDate.getHours() < 7) {
    refDate.setDate(refDate.getDate() - 1);
  }

  if (periodo === "hoy") {
    desde = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate(), 7, 0, 0);
  } else if (periodo === "ayer") {
    // v11.13: Ayer = día laboral anterior completo (7 AM día-1 → 7 AM día actual)
    var ayerRef = new Date(refDate);
    ayerRef.setDate(ayerRef.getDate() - 1);
    desde = new Date(ayerRef.getFullYear(), ayerRef.getMonth(), ayerRef.getDate(), 7, 0, 0);
    hasta = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate(), 7, 0, 0);
  } else if (periodo === "semana") {
    var diaSemana = refDate.getDay(); // 0=Dom, 1=Lun, ..., 6=Sab
    var diasAtras = diaSemana === 0 ? 6 : (diaSemana - 1); // Semana inicia Lunes
    desde = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate() - diasAtras, 7, 0, 0);
  } else if (periodo === "mes") {
    desde = new Date(refDate.getFullYear(), refDate.getMonth(), 1, 7, 0, 0);
  } else {
    desde = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate(), 7, 0, 0);
  }
  return { desde: desde, hasta: hasta };
}

// v11.13: Genera un texto legible del rango para mostrar en UI
function formatRangoLabel(rango, periodo) {
  if (!rango || !rango.desde || !rango.hasta) return "";
  var meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
  var dias = ["dom","lun","mar","mié","jue","vie","sáb"];
  var d = rango.desde;
  var h = rango.hasta;

  var fmtDia = function(dt) {
    return dias[dt.getDay()] + " " + dt.getDate() + " " + meses[dt.getMonth()];
  };
  var fmtHora = function(dt) {
    var hh = dt.getHours();
    var mm = String(dt.getMinutes()).length < 2 ? "0" + dt.getMinutes() : String(dt.getMinutes());
    return hh + ":" + mm;
  };

  if (periodo === "hoy") {
    return fmtDia(d) + " · desde " + fmtHora(d) + " hasta " + fmtHora(h);
  }
  if (periodo === "ayer") {
    return fmtDia(d) + " · turno laboral completo (7 AM a 7 AM)";
  }
  if (periodo === "semana") {
    // Rango desde el lunes hasta ahora
    return "Del " + fmtDia(d) + " al " + fmtDia(h);
  }
  if (periodo === "mes") {
    return "Mes en curso · desde el " + fmtDia(d);
  }
  // custom
  return "Del " + fmtDia(d) + " al " + fmtDia(h);
}

// Parsea fecha del reporte a Date a las 7 AM del día indicado (día laboral).
// Acepta: Date object (de getValues sobre celda fecha), "DD/MM/YYYY", "YYYY-MM-DD"
// v11.12: Normalizado a 7 AM para coincidir con el inicio del día laboral
function parseFechaDashboard(fechaStr) {
  if (fechaStr === null || fechaStr === undefined || fechaStr === "") return null;
  // FIX v11.5: Si viene como Date object (Google Sheets con formato de fecha),
  // extraer componentes y normalizar a las 7 AM local. Evita comparar 00:00 vs 07:00
  // que hacía que "hoy" saliera vacío.
  if (fechaStr instanceof Date && !isNaN(fechaStr.getTime())) {
    return new Date(fechaStr.getFullYear(), fechaStr.getMonth(), fechaStr.getDate(), 7, 0, 0);
  }
  var s = String(fechaStr).trim();
  if (!s) return null;
  // Formato DD/MM/YYYY
  if (s.indexOf("/") >= 0) {
    var parts = s.split("/");
    if (parts.length === 3) {
      var day = parseInt(parts[0]);
      var month = parseInt(parts[1]);
      var year = parseInt(parts[2]);
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        return new Date(year, month - 1, day, 7, 0, 0);
      }
    }
  }
  // Formato YYYY-MM-DD
  if (s.indexOf("-") >= 0) {
    var parts2 = s.split("-");
    if (parts2.length === 3) {
      var y = parseInt(parts2[0]);
      var m = parseInt(parts2[1]);
      var d0 = parseInt(parts2[2]);
      if (!isNaN(y) && !isNaN(m) && !isNaN(d0)) {
        return new Date(y, m - 1, d0, 7, 0, 0);
      }
    }
  }
  // Último recurso: intentar parsear como Date genérico y normalizar a 7 AM
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 7, 0, 0);
  }
  return null;
}

function dashboardProduccion(ss, periodo) {
  var rango = getRangoDashboard(periodo || "hoy");
  var sh = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  if (!sh) return { ok: false, error: "Sheet Reportes no encontrado" };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, periodo: periodo, totalPiezas: 0, topMaquinas: [], topOperadores: [], porTurno: [], tendenciaDiaria: [], totalHorasManto: 0 };

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  var totalPiezas = 0;
  var totalHorasManto = 0;
  var porMaquina = {};
  var porOperador = {};
  var porTurno = { "1ro": 0, "2do": 0, "3ro": 0 };
  var porDia = {};

  for (var r = 1; r < data.length; r++) {
    var fechaStr = data[r][idx["Fecha"]];
    var fecha = parseFechaDashboard(fechaStr);
    if (!fecha) continue;
    if (fecha < rango.desde || fecha >= rango.hasta) continue;

    var piezas = parseFloat(data[r][idx["Piezas"]]) || 0;
    var horasManto = parseFloat(data[r][idx["Horas Manto"]]) || 0;
    var maquina = String(data[r][idx["Máquina"]] || data[r][idx["Maquina"]] || "").trim();
    var responsable = String(data[r][idx["Responsable"]] || "").trim();
    var turno = String(data[r][idx["Turno"]] || "").trim();

    totalPiezas += piezas;
    totalHorasManto += horasManto;

    if (maquina && piezas > 0) {
      porMaquina[maquina] = (porMaquina[maquina] || 0) + piezas;
    }
    if (responsable && piezas > 0) {
      porOperador[responsable] = (porOperador[responsable] || 0) + piezas;
    }
    if (porTurno.hasOwnProperty(turno)) {
      porTurno[turno] += piezas;
    }

    // Tendencia diaria
    var diaKey = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM");
    porDia[diaKey] = (porDia[diaKey] || 0) + piezas;
  }

  // Top 5 máquinas
  var topMaquinas = Object.keys(porMaquina).map(function(k) { return { maquina: k, piezas: porMaquina[k] }; });
  topMaquinas.sort(function(a, b) { return b.piezas - a.piezas; });
  topMaquinas = topMaquinas.slice(0, 5);

  // Top 5 operadores
  var topOperadores = Object.keys(porOperador).map(function(k) {
    var nombre = k.split(" - ");
    return { operador: nombre.length > 1 ? nombre[1] : k, piezas: porOperador[k] };
  });
  topOperadores.sort(function(a, b) { return b.piezas - a.piezas; });
  topOperadores = topOperadores.slice(0, 5);

  // Por turno
  var arrTurnos = [
    { turno: "1ro", piezas: porTurno["1ro"] },
    { turno: "2do", piezas: porTurno["2do"] },
    { turno: "3ro", piezas: porTurno["3ro"] }
  ];

  // Tendencia diaria ordenada
  var tendencia = Object.keys(porDia).map(function(k) { return { dia: k, piezas: porDia[k] }; });
  tendencia.sort(function(a, b) {
    var pa = a.dia.split("/"); var pb = b.dia.split("/");
    if (pa[1] !== pb[1]) return parseInt(pa[1]) - parseInt(pb[1]);
    return parseInt(pa[0]) - parseInt(pb[0]);
  });

  return {
    ok: true,
    periodo: periodo,
    totalPiezas: totalPiezas,
    totalHorasManto: Math.round(totalHorasManto * 10) / 10,
    topMaquinas: topMaquinas,
    topOperadores: topOperadores,
    porTurno: arrTurnos,
    tendenciaDiaria: tendencia
  };
}

// ============================================================
// CACHE HELPERS (v11.9)
// Los catálogos (Maquinas, Partes, Empleados) casi no cambian durante
// una sesión de uso del dashboard. Cachearlos 5 minutos reduce lecturas
// del sheet y acelera cada request 3x.
// ============================================================

var _CACHE_TTL_SEG = 300; // 5 minutos

function _cacheGetOrCompute(key, computeFn) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(key);
    if (cached) return JSON.parse(cached);
    var value = computeFn();
    // Cache put puede fallar si el string es demasiado grande (>100KB)
    try {
      cache.put(key, JSON.stringify(value), _CACHE_TTL_SEG);
    } catch (e) { /* silent: si excede tamaño, seguimos sin cache */ }
    return value;
  } catch (e) {
    // Si CacheService falla por alguna razón, cae al cómputo directo
    return computeFn();
  }
}

function getMaquinasCached(ss) {
  return _cacheGetOrCompute('maq_v1', function() { return getMaquinas(ss); });
}
function getPartesCached(ss) {
  return _cacheGetOrCompute('partes_v1', function() { return getPartes(ss); });
}
function getEmpleadosCached(ss) {
  return _cacheGetOrCompute('emp_v1', function() { return getEmpleados(ss); });
}

// v11.17: Retorna un mapa parte|op → TC_real promedio (últimos 6 meses)
// Usado para calcular la eficiencia máxima potencial en el modal por máquina
function getTiemposCicloMap(ss) {
  var sh = ss.getSheetByName("TiemposCiclo");
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  // Ventana de últimos 6 meses
  var ahora = new Date();
  var fechaCorte = new Date(ahora.getFullYear(), ahora.getMonth() - 6, ahora.getDate());

  var acumulador = {};  // key → { suma, n }
  for (var r = 1; r < data.length; r++) {
    var fecha = parseFechaDashboard(data[r][idx["Fecha"]]);
    if (fecha && fecha < fechaCorte) continue;
    var np = String(data[r][idx["NoParte"]] || "").trim();
    var op = String(data[r][idx["Operacion"]] || "").trim();
    if (!np || !op) continue;
    var tcReal = parseFloat(data[r][idx["TiempoCicloReal"]]);
    if (isNaN(tcReal) || tcReal <= 0) continue;
    var key = np + "|" + op;
    if (!acumulador[key]) acumulador[key] = { suma: 0, n: 0 };
    acumulador[key].suma += tcReal;
    acumulador[key].n += 1;
  }

  var promedios = {};
  for (var k in acumulador) {
    promedios[k] = acumulador[k].suma / acumulador[k].n;
  }
  return promedios;
}

// ============================================================
// DASHBOARD EFICIENCIA POR MÁQUINA (v11)
// ============================================================

// Constantes de negocio
var HRS_TURNO_ESTANDAR = 9.5;
var SETUP_MAX_OPERATIVO = 1.5;

// Determina rango con soporte a periodo custom
function getRangoEficiencia(periodo, desdeStr, hastaStr) {
  if (periodo === "custom" && desdeStr && hastaStr) {
    var d = parseFechaDashboard(desdeStr);
    var h = parseFechaDashboard(hastaStr);
    if (d && h) {
      // v11.12: "hasta" incluye el día laboral completo hasta las 7 AM del siguiente
      var hEnd = new Date(h.getFullYear(), h.getMonth(), h.getDate() + 1, 7, 0, 0);
      return { desde: d, hasta: hEnd };
    }
  }
  return getRangoDashboard(periodo || "hoy");
}

// v11.21: Reporte diario de mediciones TC cruzado con producción
// Para cada día del rango, agrupa por máquina y muestra qué partes+op se corrieron
// y si el ingeniero capturó TC para esa combinación (o null si no midió).
function reporteMediciones(ss, periodo, desdeStr, hastaStr) {
  try {
    var rango = getRangoEficiencia(periodo || "hoy", desdeStr, hastaStr);
    var partes = getPartesCached(ss);
    var empleados = getEmpleadosCached(ss);
    var empMap = {};
    for (var i = 0; i < empleados.length; i++) empMap[String(empleados[i].numero)] = empleados[i].nombre;

    var sh = ss.getSheetByName("Reportes") || ss.getSheetByName("BD");
    if (!sh) return { ok: false, error: "No existe la hoja 'Reportes' ni 'BD'" };

    // v11.28: Set de máquinas Principales (excluir Secundarias)
    var maquinasCat = getMaquinasCached(ss);
    var maqPrincipales = {};
    for (var mi = 0; mi < maquinasCat.length; mi++) {
      var tipoM = String(maquinasCat[mi].tipo || "Principal").trim().toLowerCase();
      if (tipoM !== "secundaria") maqPrincipales[maquinasCat[mi].maquina] = 1;
    }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, rangoLabel: formatRangoLabel(rango, periodo), dias: [], totales: { dias: 0, combinaciones: 0, conMedicion: 0, sinMedicion: 0, cobertura: 0 } };
    var headers = data[0];
    var idx = {};
    for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

    // v11.23: Helper tolerante a variantes de nombres de columna
    var getIdx = function(names) {
      for (var i = 0; i < names.length; i++) {
        if (idx[names[i]] !== undefined) return idx[names[i]];
      }
      return -1;
    };
    var idxTipo = getIdx(["Tipo"]);
    var idxFecha = getIdx(["Fecha"]);
    var idxMaq = getIdx(["Máquina", "Maquina"]);
    var idxNp = getIdx(["No. Parte", "No.Parte", "NoParte", "No parte", "Numero de Parte"]);
    var idxOp = getIdx(["Operación", "Operacion"]);
    var idxPzas = getIdx(["Piezas"]);

    if (idxTipo < 0 || idxFecha < 0 || idxMaq < 0 || idxNp < 0 || idxOp < 0 || idxPzas < 0) {
      return { ok: false, error: "Columnas no encontradas en 'Reportes'. Encontradas: " + headers.join(", ") };
    }

    // Estructura: { yyyy-MM-dd: { maquina: { np|op: {noParte, op, piezas, tcCat, medicion} } } }
    var dias = {};

    // 1. Recolectar reportes de PRODUCCIÓN del rango
    for (var r = 1; r < data.length; r++) {
      var tipo = String(data[r][idxTipo] || "").trim();
      if (tipo !== "Producción" && tipo !== "Produccion") continue;
      var fecha = parseFechaDashboard(data[r][idxFecha]);
      if (!fecha) continue;
      if (fecha < rango.desde || fecha >= rango.hasta) continue;
      var maq = String(data[r][idxMaq] || "").trim();
      var np = String(data[r][idxNp] || "").trim();
      var op = String(data[r][idxOp] || "").trim();
      var piezas = parseFloat(data[r][idxPzas]) || 0;
      if (!maq || !np || !op) continue;
      // v11.28: Excluir Secundarias
      if (!maqPrincipales[maq]) continue;

      var fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (!dias[fechaStr]) dias[fechaStr] = {};
      if (!dias[fechaStr][maq]) dias[fechaStr][maq] = {};
      var key = np + "|" + op;
      if (!dias[fechaStr][maq][key]) {
        var tcCat = (partes[np] && partes[np][op] != null) ? parseFloat(partes[np][op]) : null;
        dias[fechaStr][maq][key] = {
          noParte: np, operacion: op,
          piezas: 0, tcCatalogo: tcCat, medicion: null
        };
      }
      dias[fechaStr][maq][key].piezas += piezas;
    }

    // 2. Recolectar mediciones TC del rango y cruzar
    var shTC = ss.getSheetByName("TiemposCiclo");
    if (shTC) {
      var dataTC = shTC.getDataRange().getValues();
      if (dataTC.length >= 2) {
        var hTC = dataTC[0];
        var iTC = {};
        for (var i = 0; i < hTC.length; i++) iTC[String(hTC[i]).trim()] = i;

        var getIdxTC = function(names) {
          for (var i = 0; i < names.length; i++) {
            if (iTC[names[i]] !== undefined) return iTC[names[i]];
          }
          return -1;
        };
        var iFecha = getIdxTC(["Fecha"]);
        var iMaq = getIdxTC(["Máquina", "Maquina"]);
        var iNp = getIdxTC(["NoParte", "No. Parte", "No.Parte"]);
        var iOp = getIdxTC(["Operacion", "Operación"]);
        var iTc = getIdxTC(["TiempoCicloReal", "Tiempo Ciclo Real", "TC Real"]);
        var iIng = getIdxTC(["Ingeniero"]);
        var iTurno = getIdxTC(["Turno"]);
        var iObs = getIdxTC(["Observaciones"]);

        if (iFecha >= 0 && iMaq >= 0 && iNp >= 0 && iOp >= 0 && iTc >= 0) {
          for (var r = 1; r < dataTC.length; r++) {
            var fechaTC = parseFechaDashboard(dataTC[r][iFecha]);
            if (!fechaTC) continue;
            if (fechaTC < rango.desde || fechaTC >= rango.hasta) continue;
            var maqTC = String(dataTC[r][iMaq] || "").trim();
            var npTC = String(dataTC[r][iNp] || "").trim();
            var opTC = String(dataTC[r][iOp] || "").trim();
            var tc = parseFloat(dataTC[r][iTc]);
            if (isNaN(tc) || !maqTC || !npTC || !opTC) continue;
            // v11.28: Excluir Secundarias
            if (!maqPrincipales[maqTC]) continue;

            var ingRaw = iIng >= 0 ? String(dataTC[r][iIng] || "").trim() : "";
            var ingNum = ingRaw.split(" - ")[0].trim();
            var ingNombre = empMap[ingNum] || ingRaw;

            var fechaStrTC = Utilities.formatDate(fechaTC, Session.getScriptTimeZone(), "yyyy-MM-dd");
            if (!dias[fechaStrTC]) dias[fechaStrTC] = {};
            if (!dias[fechaStrTC][maqTC]) dias[fechaStrTC][maqTC] = {};
            var keyTC = npTC + "|" + opTC;
            if (!dias[fechaStrTC][maqTC][keyTC]) {
              var tcCatX = (partes[npTC] && partes[npTC][opTC] != null) ? parseFloat(partes[npTC][opTC]) : null;
              dias[fechaStrTC][maqTC][keyTC] = {
                noParte: npTC, operacion: opTC,
                piezas: 0, tcCatalogo: tcCatX, medicion: null
              };
            }
            var med = dias[fechaStrTC][maqTC][keyTC].medicion;
            var tcR = Math.round(tc * 100) / 100;
            if (!med) {
              dias[fechaStrTC][maqTC][keyTC].medicion = {
                tc: tcR,
                ingeniero: ingNombre,
                turno: iTurno >= 0 ? String(dataTC[r][iTurno] || "").trim() : "",
                observaciones: iObs >= 0 ? String(dataTC[r][iObs] || "").trim() : "",
                numMediciones: 1,
                tcs: [tcR]
              };
            } else {
              med.tcs.push(tcR);
              med.numMediciones = med.tcs.length;
              var suma = 0;
              for (var mm = 0; mm < med.tcs.length; mm++) suma += med.tcs[mm];
              med.tc = Math.round((suma / med.numMediciones) * 100) / 100;
            }
          }
        }
      }
    }

    // 3. Convertir a array ordenado
    var diasArr = [];
    var totalCombinaciones = 0, totalConMed = 0;
    var diasNombres = ["dom","lun","mar","mié","jue","vie","sáb"];
    var meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];

    for (var fs in dias) {
      var maquinasArr = [];
      for (var maq in dias[fs]) {
        var partesArr = [];
        for (var k in dias[fs][maq]) {
          var item = dias[fs][maq][k];
          partesArr.push(item);
          totalCombinaciones++;
          if (item.medicion) totalConMed++;
        }
        partesArr.sort(function(a,b) {
          if (a.noParte !== b.noParte) return a.noParte < b.noParte ? -1 : 1;
          return a.operacion < b.operacion ? -1 : 1;
        });
        maquinasArr.push({ maquina: maq, partes: partesArr });
      }
      maquinasArr.sort(function(a,b) { return a.maquina < b.maquina ? -1 : 1; });

      var partsFecha = fs.split("-");
      var y = parseInt(partsFecha[0]), mo = parseInt(partsFecha[1]), d0 = parseInt(partsFecha[2]);
      var dt = new Date(y, mo-1, d0);
      var fechaLegible = diasNombres[dt.getDay()] + " " + d0 + " " + meses[mo-1];

      diasArr.push({
        fecha: (d0<10?"0":"")+d0+"/"+(mo<10?"0":"")+mo+"/"+y,
        fechaSort: fs,
        fechaLegible: fechaLegible,
        maquinas: maquinasArr
      });
    }
    // Orden descendente por fecha (más reciente primero)
    diasArr.sort(function(a,b) { return a.fechaSort < b.fechaSort ? 1 : -1; });

    return {
      ok: true,
      rangoLabel: formatRangoLabel(rango, periodo),
      dias: diasArr,
      totales: {
        dias: diasArr.length,
        combinaciones: totalCombinaciones,
        conMedicion: totalConMed,
        sinMedicion: totalCombinaciones - totalConMed,
        cobertura: totalCombinaciones > 0 ? Math.round(totalConMed / totalCombinaciones * 100) : 0
      }
    };
  } catch(err) {
    return {
      ok: false,
      error: String(err && err.message ? err.message : err),
      stack: String(err && err.stack ? err.stack : "").substring(0, 500)
    };
  }
}

// v11.20: Diagnóstico: retorna estadísticas de por qué se descartan reportes
function diagnosticoLectura(ss, periodo, desdeStr, hastaStr) {
  var rango = getRangoEficiencia(periodo || "hoy", desdeStr, hastaStr);
  var proc = procesarReportesEficiencia(ss, rango, null, null, null, null);
  var maquinas = getMaquinasCached(ss);
  var maquinasCatalogo = maquinas.map(function(m) { return m.maquina; });

  // Lista máquinas descartadas y comparar con catálogo (buscar posibles matches por trim/case)
  var problemas = [];
  var descartadas = proc.diagnostico.maquinasDescartadasEjemplos || {};
  for (var nombre in descartadas) {
    var lowerNombre = nombre.toLowerCase().replace(/\s+/g, ' ').trim();
    var matches = [];
    for (var i = 0; i < maquinasCatalogo.length; i++) {
      var cat = maquinasCatalogo[i];
      var lowerCat = String(cat).toLowerCase().replace(/\s+/g, ' ').trim();
      if (lowerCat === lowerNombre && cat !== nombre) {
        matches.push(cat);
      }
    }
    problemas.push({
      nombreEnReporte: nombre,
      apariciones: descartadas[nombre],
      posibleMatch: matches.length > 0 ? matches[0] : null
    });
  }

  return {
    ok: true,
    rango: {
      desde: Utilities.formatDate(rango.desde, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      hasta: Utilities.formatDate(rango.hasta, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    },
    catalogo: {
      maquinasCount: maquinasCatalogo.length,
      maquinasLista: maquinasCatalogo
    },
    diagnostico: proc.diagnostico,
    problemasNombres: problemas,
    maquinasConReportes: Object.keys(proc.byMachine).sort()
  };
}

// Procesa reportes y agrupa por máquina (helper compartido con dashboardEfMaquina)
// turnoFilter: null / "all" / "1ro" / "3ro" — filtra reportes por turno
// maquinasArr: array de nombres de máquinas seleccionadas (null/vacío = todas del área)
function procesarReportesEficiencia(ss, rango, areaFilter, maquinaFiltro, turnoFilter, maquinasArr) {
  // v11.9: Usar catálogos cacheados (5 min TTL) — reduce 3 lecturas de sheet
  var maquinas = getMaquinasCached(ss);
  var maqMap = {};
  var hayMaquinasSeleccionadas = maquinasArr && maquinasArr.length > 0;
  for (var i = 0; i < maquinas.length; i++) {
    // v11.28: Excluir máquinas Secundarias del dashboard (soporte, no productivas)
    var tipoMaq = String(maquinas[i].tipo || "Principal").trim().toLowerCase();
    if (tipoMaq === "secundaria") continue;
    // Filtro por máquina específica (uso interno del modal)
    if (maquinaFiltro && maquinas[i].maquina !== maquinaFiltro) continue;
    // Filtro por selección múltiple: si hay lista, solo esas máquinas
    if (hayMaquinasSeleccionadas) {
      if (maquinasArr.indexOf(maquinas[i].maquina) < 0) continue;
    } else {
      // Sin selección específica: aplica filtro por área
      if (areaFilter && areaFilter !== "all" && maquinas[i].area !== areaFilter) continue;
    }
    maqMap[maquinas[i].maquina] = maquinas[i].area;
  }
  var partes = getPartesCached(ss);
  var empleados = getEmpleadosCached(ss);
  var empMap = {};
  for (var i = 0; i < empleados.length; i++) empMap[String(empleados[i].numero)] = empleados[i].nombre;

  var sh = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { byMachine: {}, empMap: empMap, partes: partes, maqMap: maqMap, diagnostico: {} };

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  // v11.20: Contadores para diagnóstico de reportes descartados
  var diag = {
    totalFilas: data.length - 1,
    fechaNoParsea: 0,
    fechaFueraRango: 0,
    maquinaNoEnCatalogo: 0,
    turnoNoMatch: 0,
    procesados: 0,
    maquinasDescartadasEjemplos: {},  // { nombre_como_está_escrito: contador } máx 10
    fechasFueraEjemplos: []  // ejemplos de fechas fuera del rango, máx 5
  };

  var byMachine = {};
  for (var r = 1; r < data.length; r++) {
    var fechaRaw = data[r][idx["Fecha"]];
    var fecha = parseFechaDashboard(fechaRaw);
    if (!fecha) {
      diag.fechaNoParsea++;
      continue;
    }
    if (fecha < rango.desde || fecha >= rango.hasta) {
      diag.fechaFueraRango++;
      // Solo guardar ejemplos si están cerca del rango (últimos 30 días para no llenar de basura antigua)
      var haceMes = new Date(rango.desde.getTime() - 30 * 24 * 60 * 60 * 1000);
      if (fecha >= haceMes && diag.fechasFueraEjemplos.length < 5) {
        diag.fechasFueraEjemplos.push(Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
      }
      continue;
    }

    var maq = String(data[r][idx["Máquina"]] || data[r][idx["Maquina"]] || "").trim();
    if (!maq || !maqMap[maq]) {
      diag.maquinaNoEnCatalogo++;
      // Guardar ejemplos de nombres mal escritos o no en catálogo (máx 10 nombres diferentes)
      if (maq && Object.keys(diag.maquinasDescartadasEjemplos).length < 10) {
        diag.maquinasDescartadasEjemplos[maq] = (diag.maquinasDescartadasEjemplos[maq] || 0) + 1;
      }
      continue;
    }

    var turnoStr = String(data[r][idx["Turno"]] || "").trim();
    // Filtro por turno específico
    if (turnoFilter && turnoFilter !== "all" && turnoStr !== turnoFilter) {
      diag.turnoNoMatch++;
      continue;
    }

    diag.procesados++;
    var tipo = String(data[r][idx["Tipo"]] || "").trim();
    var fKey = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM");
    var turnoKey = fKey + "|" + turnoStr;

    if (!byMachine[maq]) {
      byMachine[maq] = { area: maqMap[maq], turnos: {}, turnosPresentes: {}, reportes: [] };
    }
    byMachine[maq].turnos[turnoKey] = 1;
    if (turnoStr) byMachine[maq].turnosPresentes[turnoStr] = 1;
    byMachine[maq].reportes.push({
      fecha: fKey,
      turno: turnoStr,
      tipo: tipo,
      noParte: String(data[r][idx["No. Parte"]] || "").trim(),
      operacion: String(data[r][idx["Operación"]] || "").trim(),
      piezas: parseFloat(data[r][idx["Piezas"]]) || 0,
      horasSetup: parseFloat(data[r][idx["Horas Setup"]]) || 0,
      horasManto: parseFloat(data[r][idx["Horas Manto"]]) || 0,
      // v11.31: Tipo de manto (Correctivo/Preventivo/Predictivo, opcional con subtipo)
      tipoManto: idx["Tipo Manto"] != null ? String(data[r][idx["Tipo Manto"]] || "").trim() : "",
      responsable: String(data[r][idx["Responsable"]] || "").trim()
    });
  }
  return { byMachine: byMachine, empMap: empMap, partes: partes, maqMap: maqMap, diagnostico: diag };
}

// Calcula desglose para 1 conjunto de reportes (Prod, Setup, Sobre, Manto, Ocioso, actividad)
// Retorna: { hProd, hSet, hOver, hManto, hIdle, piezas, sinStd, acts }
// NOTA: Los reportes sin TiempoCiclo en catálogo SÍ se incluyen en 'acts' con flag sinStd,
//       para que el usuario pueda identificarlos. Se cuentan sus piezas pero hrs=0.
// v11.17: Retorna mapa {parte|op: promedio_min} con TC medidos por ingeniería
// en los últimos 6 meses. Usado para recalibrar la eficiencia de la máquina.
function getPromediosTCMedidos(ss) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('tc_medidos_v1');
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  var sh = ss.getSheetByName("TiemposCiclo");
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  var ahora = new Date();
  var fechaCorte = new Date(ahora.getFullYear(), ahora.getMonth() - 6, ahora.getDate());

  var acc = {};
  for (var r = 1; r < data.length; r++) {
    var tcReal = parseFloat(data[r][idx["TiempoCicloReal"]]);
    if (isNaN(tcReal)) continue;
    var fechaRaw = data[r][idx["Fecha"]];
    var fechaObj = parseFechaDashboard(fechaRaw);
    if (fechaObj && fechaObj < fechaCorte) continue;
    var np = String(data[r][idx["NoParte"]] || "").trim();
    var op = String(data[r][idx["Operacion"]] || "").trim();
    if (!np || !op) continue;
    var k = np + "|" + op;
    if (!acc[k]) acc[k] = { suma: 0, count: 0 };
    acc[k].suma += tcReal;
    acc[k].count += 1;
  }

  var result = {};
  for (var k in acc) result[k] = acc[k].suma / acc[k].count;

  try {
    CacheService.getScriptCache().put('tc_medidos_v1', JSON.stringify(result), 120);
  } catch(e) {}
  return result;
}

function calcularDesglose(reportes, N, empMap, partes, tcOverride) {
  var totalHrs = N * HRS_TURNO_ESTANDAR;
  var hSet = 0, hOver = 0, hProd = 0, hManto = 0, sinStd = 0, piezas = 0;
  var hOtro = 0;
  var acts = {};
  var partesConMedicion = {};
  var partesSinMedicion = {};
  // v11.29: Lista de reportes de manto y setup para diagnóstico
  var actsMantoSetup = [];

  for (var i = 0; i < reportes.length; i++) {
    var rep = reportes[i];
    // v11.25: Detectar reporte "Otro" — la parte empieza con "Otro"
    var esOtro = String(rep.noParte || "").indexOf("Otro") === 0;

    if (rep.tipo === "Setup") {
      hSet += Math.min(rep.horasSetup, SETUP_MAX_OPERATIVO);
      hOver += Math.max(0, rep.horasSetup - SETUP_MAX_OPERATIVO);
      // v11.29: Registrar para diagnóstico
      var respPS = rep.responsable.split(" - ");
      actsMantoSetup.push({
        kind: "setup",
        fecha: rep.fecha, turno: rep.turno,
        resp: respPS[0].trim(),
        respName: respPS.length > 1 ? respPS.slice(1).join(" - ").trim() : (empMap[respPS[0].trim()] || rep.responsable),
        hrs: rep.horasSetup,
        detalle: rep.noParte + " · op " + rep.operacion
      });
    }
    // Manto: el tab de la app puede reportar "Manto" o "Mantenimiento" según versión
    if (rep.tipo === "Manto" || rep.tipo === "Mantenimiento") {
      hManto += rep.horasManto || 0;
      // v11.29: Registrar para diagnóstico
      var respPM = rep.responsable.split(" - ");
      actsMantoSetup.push({
        kind: "manto",
        fecha: rep.fecha, turno: rep.turno,
        resp: respPM[0].trim(),
        respName: respPM.length > 1 ? respPM.slice(1).join(" - ").trim() : (empMap[respPM[0].trim()] || rep.responsable),
        hrs: rep.horasManto || 0,
        // v11.31: Clasificación del manto (Correctivo/Preventivo/Predictivo)
        tipoManto: rep.tipoManto || "",
        detalle: rep.noParte || "-"
      });
    }
    if (rep.tipo === "Producción" || rep.tipo === "Produccion") {
      piezas += rep.piezas;
      var h = 0;
      var isSinStd = false;

      if (esOtro) {
        // v11.25: "Otro" no tiene TC en catálogo; el operador reporta las horas
        // manualmente en el campo Horas Setup. Se contabilizan como operativas.
        h = parseFloat(rep.horasSetup) || 0;
        hOtro += h;
        // No cuenta como sinStd (no es un error del catálogo, es intencional)
      } else {
        var tc = null;
        var poKey = rep.noParte + "|" + rep.operacion;
        // v11.17: Prioridad al TC medido si viene tcOverride
        if (tcOverride && tcOverride[poKey] !== undefined && tcOverride[poKey] !== null) {
          tc = tcOverride[poKey];
          partesConMedicion[poKey] = 1;
        } else if (partes[rep.noParte] && partes[rep.noParte][rep.operacion] != null) {
          tc = parseFloat(partes[rep.noParte][rep.operacion]);
          if (tcOverride) partesSinMedicion[poKey] = 1;
        }
        isSinStd = (tc == null || isNaN(tc) || tc <= 0);
        if (isSinStd) sinStd += 1;
        h = isSinStd ? 0 : (rep.piezas * tc) / 60;
        if (!isSinStd) hProd += h;
      }

      var respParts = rep.responsable.split(" - ");
      var respNum = respParts[0].trim();
      var respName = respParts.length > 1 ? respParts.slice(1).join(" - ").trim() : (empMap[respNum] || rep.responsable);
      var aKey = respNum + "|" + rep.noParte + "|" + rep.operacion;
      if (!acts[aKey]) {
        acts[aKey] = { resp: respNum, respName: respName, np: rep.noParte, op: rep.operacion,
                       pzas: 0, hrs: 0, sinStd: false, esOtro: esOtro };
      }
      acts[aKey].pzas += rep.piezas;
      acts[aKey].hrs += h;
      // Si alguno de los reportes agrupados no tiene estándar (excepto Otro), marcar la línea
      if (isSinStd) acts[aKey].sinStd = true;
    }
  }
  var hIdle = Math.max(0, totalHrs - hProd - hSet - hOver - hManto - hOtro);
  var hDisponible = Math.max(0, totalHrs - hManto);
  var UMBRAL_MANTO_NA = 0.90;
  var pctManto = totalHrs > 0 ? (hManto / totalHrs) : 0;
  // v11.25: Horas operadas ahora suma: producción con catálogo + montaje ≤1.5h + "Otro"
  var horasOperadas = hProd + hSet + hOtro;
  var eff;
  if (pctManto >= UMBRAL_MANTO_NA) {
    eff = null;
  } else if (hDisponible > 0) {
    eff = Math.round((horasOperadas / hDisponible) * 100);
  } else {
    eff = 0;
  }
  var actsArr = [];
  for (var k in acts) actsArr.push({
    resp: acts[k].resp, respName: acts[k].respName, np: acts[k].np, op: acts[k].op,
    pzas: acts[k].pzas, hrs: Math.round(acts[k].hrs * 100) / 100,
    sinStd: acts[k].sinStd
  });
  // Ordenar: primero por horas desc, pero los sinStd al final
  actsArr.sort(function(a, b) {
    if (a.sinStd !== b.sinStd) return a.sinStd ? 1 : -1;
    return b.hrs - a.hrs;
  });

  return {
    totalHrs: totalHrs,
    hDisponible: Math.round(hDisponible * 100) / 100,
    hProd: Math.round(hProd * 100) / 100,
    hSet: Math.round(hSet * 100) / 100,
    hOver: Math.round(hOver * 100) / 100,
    hManto: Math.round(hManto * 100) / 100,
    hIdle: Math.round(hIdle * 100) / 100,
    // v11.22: horas operativas totales (prod + setup normal + Otro), usado en fórmula de eficiencia
    hOperadas: Math.round(horasOperadas * 100) / 100,
    // v11.25: horas de "Otro" (partes ad-hoc reportadas manualmente)
    hOtro: Math.round(hOtro * 100) / 100,
    piezas: piezas, sinStd: sinStd, eff: eff, acts: actsArr,
    // v11.29: Reportes de manto y setup para diagnóstico visual
    actsMantoSetup: actsMantoSetup,
    // v11.17: solo se llenan cuando se pasó tcOverride
    partesConMedicion: Object.keys(partesConMedicion).length,
    partesSinMedicion: Object.keys(partesSinMedicion).length
  };
}

// Endpoint principal: dashboard de eficiencia por máquina
// maquinasStr: string separado por comas de nombres de máquinas seleccionadas
//              (vacío/null = todas del área)
function dashboardEficiencia(ss, periodo, area, desdeStr, hastaStr, turno, maquinasStr) {
  var rango = getRangoEficiencia(periodo, desdeStr, hastaStr);
  var maquinasArr = null;
  if (maquinasStr) {
    maquinasArr = String(maquinasStr).split(",").map(function(s){return s.trim();}).filter(function(s){return s.length > 0;});
    if (maquinasArr.length === 0) maquinasArr = null;
  }
  var proc = procesarReportesEficiencia(ss, rango, area, null, turno, maquinasArr);

  var cards = [];
  var totalPiezas = 0, totalHrsProd = 0, totalHrsDisponible = 0;

  // v11.18: Cargar TC medidos una sola vez para todas las máquinas
  var tcMedidos = getPromediosTCMedidos(ss);

  for (var maq in proc.byMachine) {
    var m = proc.byMachine[maq];
    var N = Object.keys(m.turnos).length;
    var d = calcularDesglose(m.reportes, N, proc.empMap, proc.partes);
    // v11.18: Segundo cálculo con TC medidos (para eficiencia recalibrada)
    var dRec = calcularDesglose(m.reportes, N, proc.empMap, proc.partes, tcMedidos);

    // Turnos presentes ordenados: 1ro, 3ro
    var presentes = Object.keys(m.turnosPresentes).sort();

    cards.push({
      maquina: maq, area: m.area, turnos: N,
      turnosPresentes: presentes,
      eff: d.eff, piezas: d.piezas,
      hProd: d.hProd, hSet: d.hSet, hOver: d.hOver, hManto: d.hManto, hIdle: d.hIdle,
      hOtro: d.hOtro,  // v11.27
      totalHrs: d.totalHrs, hDisponible: d.hDisponible,
      sinStd: d.sinStd,
      acts: d.acts,
      sinActividad: false,
      // v11.18: Eficiencia recalibrada con TC medidos
      effRecalibrada: dRec.eff,
      hProdRecalibrada: dRec.hProd,
      partesConMedicion: dRec.partesConMedicion,
      partesSinMedicion: dRec.partesSinMedicion
    });
    totalPiezas += d.piezas;
    totalHrsProd += d.hProd;
    totalHrsDisponible += d.hDisponible;
  }

  // v11.6: Máquinas seleccionadas sin actividad se incluyen con eff=0 (Regla 2:
  // penalizan el promedio global). El flag "sinActividad: true" se mantiene
  // para el frontend, que muestra visual apagada pero eff=0 cuenta en KPIs.
  if (maquinasArr && maquinasArr.length > 0) {
    for (var i = 0; i < maquinasArr.length; i++) {
      var nom = maquinasArr[i];
      if (!proc.byMachine[nom] && proc.maqMap[nom]) {
        cards.push({
          maquina: nom, area: proc.maqMap[nom], turnos: 0,
          turnosPresentes: [],
          eff: 0,  // <-- v11.6: cuenta como 0% en promedio global
          piezas: 0,
          hProd: 0, hSet: 0, hOver: 0, hManto: 0, hIdle: 0,
          totalHrs: 0, hDisponible: 0,
          sinStd: 0, acts: [],
          sinActividad: true
        });
      }
    }
  }

  // v11.6: Cálculo del promedio global.
  // - Máquinas con eff = null (todo manto) → EXCLUIDAS del promedio (Regla 3)
  // - Máquinas con eff = 0 (sin actividad o cero producción) → CUENTAN como 0
  // - Máquinas con eff = N → CUENTAN como N
  var effsValidas = [];
  var numMaquinasEnManto = 0;
  for (var j = 0; j < cards.length; j++) {
    if (cards[j].eff === null) {
      numMaquinasEnManto++;
    } else {
      effsValidas.push(cards[j].eff);
    }
  }
  var effGlobal = 0;
  if (effsValidas.length > 0) {
    var sumaEff = 0;
    for (var k = 0; k < effsValidas.length; k++) sumaEff += effsValidas[k];
    effGlobal = Math.round(sumaEff / effsValidas.length);
  }

  return {
    ok: true,
    periodo: periodo,
    rangoLabel: formatRangoLabel(rango, periodo),
    area: area || "all",
    turno: turno || "all",
    maquinasFiltradas: maquinasArr,
    effGlobal: effGlobal,
    numMaquinasEnManto: numMaquinasEnManto,   // Excluidas del promedio (N/A)
    numMaquinasContadas: effsValidas.length,   // Contribuyen al promedio
    totalPiezas: totalPiezas,
    totalHrsDisp: Math.round(totalHrsDisponible * 10) / 10,
    totalHrsProd: Math.round(totalHrsProd * 10) / 10,
    numMaquinas: cards.length,
    cards: cards,
    // v11.20: Diagnóstico de lectura de reportes (para detectar problemas)
    diagnostico: proc.diagnostico || null
  };
}

// Endpoint secundario: tendencia diaria de eficiencia de una máquina (para modal)
function dashboardEfMaquina(ss, maquina, periodo, desdeStr, hastaStr, turno) {
  var rango = getRangoEficiencia(periodo, desdeStr, hastaStr);
  // Al pedir una máquina específica, no se aplica filtro de área ni de selección múltiple
  var proc = procesarReportesEficiencia(ss, rango, "all", maquina, turno, null);

  var m = proc.byMachine[maquina];
  if (!m) return { ok: true, maquina: maquina, tendencia: [], area: "", resumen: null };

  // Agrupar reportes por día (todos los turnos de un día se suman)
  var byDia = {};
  for (var i = 0; i < m.reportes.length; i++) {
    var r = m.reportes[i];
    if (!byDia[r.fecha]) byDia[r.fecha] = { turnos: {}, reportes: [] };
    byDia[r.fecha].turnos[r.turno] = 1;
    byDia[r.fecha].reportes.push(r);
  }

  var tendencia = [];
  for (var dia in byDia) {
    var d = byDia[dia];
    var N = Object.keys(d.turnos).length;
    var desg = calcularDesglose(d.reportes, N, proc.empMap, proc.partes);
    tendencia.push({
      dia: dia, turnos: N,
      eff: desg.eff, piezas: desg.piezas,
      hProd: desg.hProd, hSet: desg.hSet, hOver: desg.hOver, hManto: desg.hManto, hIdle: desg.hIdle,
      hDisponible: desg.hDisponible,
      sinStd: desg.sinStd
    });
  }
  // Ordenar por fecha (dd/MM)
  tendencia.sort(function(a, b) {
    var pa = a.dia.split("/"); var pb = b.dia.split("/");
    if (pa[1] !== pb[1]) return parseInt(pa[1]) - parseInt(pb[1]);
    return parseInt(pa[0]) - parseInt(pb[0]);
  });

  // Resumen agregado del periodo
  var totalDesg = calcularDesglose(m.reportes, Object.keys(m.turnos).length, proc.empMap, proc.partes);

  // v11.8: Desglose por turno (1ro / 3ro)
  var reportesPorTurno = { "1ro": [], "3ro": [] };
  var turnosUnicosPorTurno = { "1ro": {}, "3ro": {} };
  for (var k = 0; k < m.reportes.length; k++) {
    var rr = m.reportes[k];
    if (rr.turno === "1ro" || rr.turno === "3ro") {
      reportesPorTurno[rr.turno].push(rr);
      turnosUnicosPorTurno[rr.turno][rr.fecha] = 1;
    }
  }
  var porTurno = {};
  ["1ro", "3ro"].forEach(function(t) {
    var Nt = Object.keys(turnosUnicosPorTurno[t]).length;
    if (Nt > 0) {
      var desgT = calcularDesglose(reportesPorTurno[t], Nt, proc.empMap, proc.partes);
      porTurno[t] = {
        turnos: Nt,
        eff: desgT.eff, piezas: desgT.piezas,
        hProd: desgT.hProd, hSet: desgT.hSet,
        hOver: desgT.hOver, hManto: desgT.hManto, hIdle: desgT.hIdle,
        hOtro: desgT.hOtro,  // v11.27
        totalHrs: desgT.totalHrs, hDisponible: desgT.hDisponible,
        sinStd: desgT.sinStd
      };
    }
  });

  // v11.17: Cálculo de EFICIENCIA RECALIBRADA con TC medidos por ingeniería
  // Se recalcula hProd usando los promedios de TC medidos de últimos 6 meses.
  // Para partes+op sin mediciones, se usa el TC del catálogo como fallback.
  var tcMedidos = getPromediosTCMedidos(ss);
  var desgloseRecal = calcularDesglose(m.reportes, Object.keys(m.turnos).length, proc.empMap, proc.partes, tcMedidos);

  return {
    ok: true,
    maquina: maquina,
    area: m.area,
    turno: turno || "all",
    turnosPresentes: Object.keys(m.turnosPresentes).sort(),
    resumen: {
      turnos: Object.keys(m.turnos).length,
      eff: totalDesg.eff,
      piezas: totalDesg.piezas,
      hProd: totalDesg.hProd, hSet: totalDesg.hSet,
      hOver: totalDesg.hOver, hManto: totalDesg.hManto, hIdle: totalDesg.hIdle,
      hOtro: totalDesg.hOtro,  // v11.27
      totalHrs: totalDesg.totalHrs,
      hDisponible: totalDesg.hDisponible,
      sinStd: totalDesg.sinStd,
      acts: totalDesg.acts,
      actsMantoSetup: totalDesg.actsMantoSetup,  // v11.29
      porTurno: porTurno,
      // v11.17: Eficiencia recalibrada con TC medidos por ingeniería
      effRecalibrada: desgloseRecal.eff,
      hProdRecalibrada: desgloseRecal.hProd,
      partesConMedicion: desgloseRecal.partesConMedicion,
      partesSinMedicion: desgloseRecal.partesSinMedicion
    },
    tendencia: tendencia
  };
}


// ============================================================
// DETALLE POR PARTE + OPERACIÓN (v11.11)
// Alcance global: información completa de una parte + operación
// SIN filtrar por máquina, turno u operador.
// Sí filtra por rango de fechas (independiente del dashboard).
// Retorna:
//   - Info general: TiempoCiclo, piezas/hora estándar, totales
//   - Operaciones disponibles del catálogo para esa parte (para selector)
//   - Tendencia diaria: día → piezas, hrs, # máquinas activas
//   - Tabla consolidada operador+máquina
// ============================================================

function detalleParteOperacion(ss, noParte, operacion, periodo, desdeStr, hastaStr) {
  var rango = getRangoEficiencia(periodo, desdeStr, hastaStr);

  var partes = getPartesCached(ss);
  var empleados = getEmpleadosCached(ss);
  var empMap = {};
  for (var i = 0; i < empleados.length; i++) empMap[String(empleados[i].numero)] = empleados[i].nombre;

  // Operaciones disponibles para esta parte (del catálogo)
  var opsDisponibles = [];
  if (partes[noParte]) {
    for (var op in partes[noParte]) opsDisponibles.push(op);
    opsDisponibles.sort();
  }

  // TiempoCiclo estándar del catálogo (puede ser null si no está)
  var tc = null;
  if (partes[noParte] && partes[noParte][operacion] != null) {
    tc = parseFloat(partes[noParte][operacion]);
    if (isNaN(tc)) tc = null;
  }

  var sh = ss.getSheetByName("Reportes") || ss.getActiveSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return {
      ok: true, noParte: noParte, operacion: operacion,
      operacionesDisponibles: opsDisponibles,
      tiempoCiclo: tc, pzsPorHoraEstandar: (tc && tc > 0) ? Math.round((60/tc)*100)/100 : null,
      totalPiezas: 0, totalHorasEstandar: 0, numReportes: 0,
      tendencia: [], porOperadorMaquina: [], maquinasActivas: []
    };
  }

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  var totalPiezas = 0;
  var totalHoras = 0;
  var numReportes = 0;
  var byDia = {};
  var byOperMaq = {};
  var maquinasSet = {};

  for (var r = 1; r < data.length; r++) {
    var fecha = parseFechaDashboard(data[r][idx["Fecha"]]);
    if (!fecha) continue;
    if (fecha < rango.desde || fecha >= rango.hasta) continue;

    var tipo = String(data[r][idx["Tipo"]] || "").trim();
    if (tipo !== "Producción" && tipo !== "Produccion") continue;

    var np = String(data[r][idx["No. Parte"]] || "").trim();
    if (np !== noParte) continue;

    var op = String(data[r][idx["Operación"]] || "").trim();
    if (op !== operacion) continue;

    var maquina = String(data[r][idx["Máquina"]] || data[r][idx["Maquina"]] || "").trim();
    var responsableRep = String(data[r][idx["Responsable"]] || "").trim();
    var respNum = responsableRep.split(" - ")[0].trim();
    var turnoStr = String(data[r][idx["Turno"]] || "").trim();

    var piezas = parseFloat(data[r][idx["Piezas"]]) || 0;
    var fKey = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM");
    var h = (tc !== null) ? (piezas * tc) / 60 : 0;

    totalPiezas += piezas;
    totalHoras += h;
    numReportes++;
    if (maquina) maquinasSet[maquina] = 1;

    // Tendencia diaria
    if (!byDia[fKey]) byDia[fKey] = { piezas: 0, horas: 0, maquinas: {}, reportes: 0 };
    byDia[fKey].piezas += piezas;
    byDia[fKey].horas += h;
    byDia[fKey].reportes += 1;
    if (maquina) byDia[fKey].maquinas[maquina] = 1;

    // Consolidado por operador+máquina
    var omKey = respNum + "|" + maquina;
    if (!byOperMaq[omKey]) {
      byOperMaq[omKey] = {
        operador: respNum,
        operadorNombre: empMap[respNum] || (responsableRep.split(" - ")[1] || respNum),
        maquina: maquina,
        piezas: 0,
        horas: 0,
        numReportes: 0,
        turnos: {},
        primerDia: fKey,
        ultimoDia: fKey
      };
    }
    byOperMaq[omKey].piezas += piezas;
    byOperMaq[omKey].horas += h;
    byOperMaq[omKey].numReportes += 1;
    byOperMaq[omKey].turnos[turnoStr] = 1;
    // Actualizar ultimoDia si es más reciente (comparación básica dd/MM)
    if (fKey > byOperMaq[omKey].ultimoDia) byOperMaq[omKey].ultimoDia = fKey;
    if (fKey < byOperMaq[omKey].primerDia) byOperMaq[omKey].primerDia = fKey;
  }

  // Tendencia ordenada por fecha (dd/MM)
  var tendencia = [];
  for (var dia in byDia) {
    tendencia.push({
      dia: dia,
      piezas: byDia[dia].piezas,
      horas: Math.round(byDia[dia].horas * 100) / 100,
      maquinas: Object.keys(byDia[dia].maquinas).length,
      reportes: byDia[dia].reportes
    });
  }
  tendencia.sort(function(a, b) {
    var pa = a.dia.split("/"); var pb = b.dia.split("/");
    if (pa[1] !== pb[1]) return parseInt(pa[1]) - parseInt(pb[1]);
    return parseInt(pa[0]) - parseInt(pb[0]);
  });

  // Tabla operador-máquina ordenada por piezas descendente
  var porOperadorMaquina = [];
  for (var k in byOperMaq) {
    var om = byOperMaq[k];
    porOperadorMaquina.push({
      operador: om.operador,
      operadorNombre: om.operadorNombre,
      maquina: om.maquina,
      piezas: om.piezas,
      horas: Math.round(om.horas * 100) / 100,
      numReportes: om.numReportes,
      turnos: Object.keys(om.turnos).sort()
    });
  }
  porOperadorMaquina.sort(function(a, b) { return b.piezas - a.piezas; });

  return {
    ok: true,
    noParte: noParte,
    operacion: operacion,
    rangoLabel: formatRangoLabel(rango, periodo),
    operacionesDisponibles: opsDisponibles,
    tiempoCiclo: tc,
    pzsPorHoraEstandar: (tc !== null && tc > 0) ? Math.round((60 / tc) * 100) / 100 : null,
    totalPiezas: totalPiezas,
    totalHorasEstandar: Math.round(totalHoras * 100) / 100,
    numReportes: numReportes,
    maquinasActivas: Object.keys(maquinasSet).sort(),
    tendencia: tendencia,
    porOperadorMaquina: porOperadorMaquina
  };
}


// ============================================================
// TIEMPOS DE CICLO — Módulo de Ingeniería (v11.14)
// ============================================================

// Retorna histórico de mediciones de tiempo de ciclo para una parte + operación
// v11.15: Filtra últimos 180 días + calcula estadísticas (mediana, desv est, CV)
function tiemposCicloPorParteOp(ss, noParte, operacion) {
  var sh = ss.getSheetByName("TiemposCiclo");
  if (!sh) return { ok: true, mediciones: [], promedio: null, minimo: null, maximo: null, catalogo: null };

  // v11.15: Ventana de últimos 6 meses (180 días)
  var ahora = new Date();
  var fechaCorte = new Date(ahora.getFullYear(), ahora.getMonth() - 6, ahora.getDate());

  var data = sh.getDataRange().getValues();
  var partes = getPartesCached(ss);
  var tcCatalogo = (partes[noParte] && partes[noParte][operacion] != null) ? parseFloat(partes[noParte][operacion]) : null;

  if (data.length < 2) {
    return {
      ok: true, mediciones: [], promedio: null, minimo: null, maximo: null,
      catalogo: tcCatalogo, ventanaMeses: 6,
      mediana: null, desviacionEstandar: null, coeficienteVariacion: null,
      estabilidad: null, numMediciones: 0
    };
  }

  var headers = data[0];
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[String(headers[i]).trim()] = i;

  var empleados = getEmpleadosCached(ss);
  var empMap = {};
  for (var i = 0; i < empleados.length; i++) empMap[String(empleados[i].numero)] = empleados[i].nombre;

  var mediciones = [];
  var valores = [];  // solo los TC reales, para stats
  var suma = 0, minVal = null, maxVal = null;

  for (var r = 1; r < data.length; r++) {
    var np = String(data[r][idx["NoParte"]] || "").trim();
    if (np !== noParte) continue;
    var op = String(data[r][idx["Operacion"]] || "").trim();
    if (op !== operacion) continue;

    var tcReal = parseFloat(data[r][idx["TiempoCicloReal"]]);
    if (isNaN(tcReal)) continue;

    // v11.15: Filtrar por ventana de 6 meses (usando fecha del reporte)
    var fechaRaw = data[r][idx["Fecha"]];
    var fechaObj = parseFechaDashboard(fechaRaw);
    if (fechaObj && fechaObj < fechaCorte) continue;

    var fechaStr = (fechaRaw instanceof Date)
      ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), "dd/MM/yyyy")
      : String(fechaRaw || "").trim();

    var ingRaw = String(data[r][idx["Ingeniero"]] || "").trim();
    var ingNum = ingRaw.split(" - ")[0].trim();
    var ingNombre = empMap[ingNum] || ingRaw;

    mediciones.push({
      fecha: fechaStr,
      turno: String(data[r][idx["Turno"]] || "").trim(),
      ingeniero: ingNum,
      ingenieroNombre: ingNombre,
      maquina: String(data[r][idx["Máquina"]] || "").trim(),
      tiempoCicloReal: Math.round(tcReal * 100) / 100,
      observaciones: String(data[r][idx["Observaciones"]] || "").trim(),
      timestamp: String(data[r][idx["Timestamp"]] || "").trim()
    });
    valores.push(tcReal);
    suma += tcReal;
    if (minVal === null || tcReal < minVal) minVal = tcReal;
    if (maxVal === null || tcReal > maxVal) maxVal = tcReal;
  }

  // Ordenar mediciones por timestamp descendente (más reciente primero)
  mediciones.sort(function(a, b) {
    if (a.timestamp > b.timestamp) return -1;
    if (a.timestamp < b.timestamp) return 1;
    return 0;
  });

  var n = valores.length;
  var promedio = n > 0 ? suma / n : null;

  // v11.15: Cálculo de mediana
  var mediana = null;
  if (n > 0) {
    var ordenados = valores.slice().sort(function(a, b) { return a - b; });
    if (n % 2 === 1) {
      mediana = ordenados[Math.floor(n / 2)];
    } else {
      mediana = (ordenados[n / 2 - 1] + ordenados[n / 2]) / 2;
    }
  }

  // v11.15: Desviación estándar muestral (n-1) — más apropiada para muestras pequeñas
  var desvEst = null;
  var cv = null;
  var estabilidad = null;
  if (n >= 3) {
    var sumaCuadradosDif = 0;
    for (var v = 0; v < n; v++) {
      sumaCuadradosDif += Math.pow(valores[v] - promedio, 2);
    }
    desvEst = Math.sqrt(sumaCuadradosDif / (n - 1));
    if (promedio > 0) {
      cv = (desvEst / promedio) * 100;
      // Interpretación de estabilidad basada en CV
      if (cv < 5) estabilidad = "Muy estable";
      else if (cv < 15) estabilidad = "Estable";
      else if (cv < 25) estabilidad = "Variable";
      else estabilidad = "Inestable";
    }
  }

  return {
    ok: true,
    noParte: noParte,
    operacion: operacion,
    catalogo: tcCatalogo,
    ventanaMeses: 6,
    numMediciones: n,
    promedio: n > 0 ? Math.round(promedio * 100) / 100 : null,
    mediana: mediana !== null ? Math.round(mediana * 100) / 100 : null,
    desviacionEstandar: desvEst !== null ? Math.round(desvEst * 100) / 100 : null,
    coeficienteVariacion: cv !== null ? Math.round(cv * 10) / 10 : null,
    estabilidad: estabilidad,
    minimo: minVal !== null ? Math.round(minVal * 100) / 100 : null,
    maximo: maxVal !== null ? Math.round(maxVal * 100) / 100 : null,
    variacionVsCatalogo: (tcCatalogo !== null && promedio !== null)
      ? Math.round(((promedio - tcCatalogo) / tcCatalogo * 100) * 10) / 10
      : null,
    // v11.16: Eficiencia = TC catálogo / TC real × 100
    // >100% = más rápido que el estándar (posible sobrestimación del catálogo)
    // ~100% = a ritmo del estándar
    // <100% = más lento que el estándar
    eficiencia: (tcCatalogo !== null && promedio !== null && promedio > 0)
      ? Math.round((tcCatalogo / promedio) * 100 * 10) / 10
      : null,
    mediciones: mediciones
  };
}

// Busca partes en el catálogo por coincidencia parcial del número
// Retorna hasta 20 partes con las operaciones que tiene cada una
function buscarPartes(ss, query) {
  var partes = getPartesCached(ss);
  var q = String(query || "").trim().toLowerCase();
  if (q.length < 2) return { ok: true, resultados: [] };

  var resultados = [];
  for (var np in partes) {
    if (String(np).toLowerCase().indexOf(q) < 0) continue;
    var ops = [];
    for (var op in partes[np]) ops.push(op);
    ops.sort();
    resultados.push({
      noParte: np,
      operaciones: ops
    });
    if (resultados.length >= 20) break;
  }
  // Ordenar alfabéticamente por número de parte
  resultados.sort(function(a, b) {
    if (a.noParte < b.noParte) return -1;
    if (a.noParte > b.noParte) return 1;
    return 0;
  });
  return { ok: true, resultados: resultados };
}

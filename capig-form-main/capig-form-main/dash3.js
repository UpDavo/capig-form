/**
 * Tablas para el dashboard de porcentaje de gerentes por genero y tamano.
 * Ejecutar refreshDashboardGenero() para regenerar los pre-aggregados.
 */
(function (global) {
  const SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    OUT_DETAIL: "DASH_GENERO_GERENTES",
    OUT_PIVOT: "PIVOT_GERENTES_GENERO_TAMANO",
    OUT_WIDE: "PIVOT_GERENTES_GENERO_TAMANO_WIDE",
  };

  const TAMANO_BY_CODE = { 1: "MICRO", 2: "PEQUENA", 3: "MEDIANA", 4: "GRANDE" };
  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, GLOBAL: 0 };
  const GENDER_ALIASES = ["GENERO", "GÉNERO", "G�NERO", "GÊNERO"];
  const TAMANO_ALIASES = ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMA�'O", "TAMANO_EMP", "TAMAÑO", "TAMA�O"];
  const FECHA_ALIASES = ["FECHA_AFILIACION", "FECHA AFILIACION", "FECHA_INGRESO", "FECHA DE INGRESO"];
  const CARGO_ALIASES = ["CARGO", "PUESTO", "OCUPACION"];

  // -------------- Utils ---------------- //
  function normalizeLabel(label) {
    let txt = (label || "").toString().trim().toUpperCase();
    txt = txt.replace(/���/g, "N");
    txt = txt.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    txt = txt.replace(/[^A-Z0-9]+/g, "_");
    txt = txt.replace(/^_+|_+$/g, "");
    return txt;
  }
  function normalizeName(val) {
    return (val || "").toString().trim().toUpperCase().replace(/\s+/g, " ");
  }
  function cleanRuc(val) {
    return (val || "").toString().replace(/[^0-9]/g, "");
  }
  function padRuc13(ruc) {
    const r = cleanRuc(ruc);
    if (!r) return "";
    if (r.length >= 13) return r;
    return r.padStart(13, "0");
  }
  function normalizeTamano(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return "";
    if (t.indexOf("MICRO") !== -1) return "MICRO";
    if (t.indexOf("PEQU") !== -1) return "PEQUENA";
    if (t.indexOf("MEDI") !== -1) return "MEDIANA";
    if (t.indexOf("GRAN") !== -1) return "GRANDE";
    return t;
  }
  function sizeFromCode(val) {
    const code = (val || "").toString().trim();
    if (TAMANO_BY_CODE[code]) return TAMANO_BY_CODE[code];
    return normalizeTamano(val);
  }
  function normalizeGender(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return "";

    // Detectar palabras completas primero para evitar ambigüedad
    // MUJER/FEMENINO/FEMENINA → FEMENINO
    if (t.includes("MUJER") || t.includes("FEMEN")) return "FEMENINO";

    // HOMBRE/MASCULINO/MASCULINA → MASCULINO
    if (t.includes("HOMBRE") || t.includes("MASC")) return "MASCULINO";

    // Fallback para abreviaciones (F, M, etc.)
    if (t.startsWith("F")) return "FEMENINO";
    if (t.startsWith("M")) return "MASCULINO";

    return "";
  }
  function parseDateFlexible(val) {
    if (!val) return null;
    let d = new Date(val);
    if (!isNaN(d.getTime())) return d;
    const str = val.toString().trim();
    const m = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }
  function getYear(val) {
    const d = parseDateFlexible(val);
    return d ? d.getFullYear().toString() : "";
  }
  function getSheetOrCreate(name) {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    return sh;
  }
  function readTableFlexibleByCandidates(names = []) {
    const ss = SpreadsheetApp.getActive();
    for (const name of names) {
      if (!name) continue;
      const sh = ss.getSheetByName(name);
      if (sh) {
        const data = readTableFlexible(name);
        return { ...data, name };
      }
    }
    return { headerIndex: {}, rows: [], name: null };
  }
  function readTableFlexible(sheetName) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { headerIndex: {}, rows: [] };

    const values = sh.getDataRange().getDisplayValues();
    let headerRowIdx = -1;
    for (let i = 0; i < values.length; i++) {
      const norm = values[i].map(normalizeLabel);
      const hasRuc = norm.includes("RUC");
      const hasRazon = norm.includes("RAZON_SOCIAL");
      const nonEmpty = norm.filter((c) => c && c !== "NO").length;
      if ((hasRuc || hasRazon) && nonEmpty >= 2) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) headerRowIdx = 0;

    const headersNorm = values[headerRowIdx].map(normalizeLabel);
    const headerIndex = {};
    headersNorm.forEach((h, idx) => {
      if (h && headerIndex[h] === undefined) headerIndex[h] = idx;
    });
    const rows = values.slice(headerRowIdx + 1);
    return { headerIndex, rows };
  }
  function getVal(row, headerIndex, aliasList) {
    for (const alias of aliasList) {
      const key = normalizeLabel(alias);
      const idx = headerIndex[key];
      if (idx !== undefined && idx < row.length) {
        const val = row[idx];
        if (val !== null && val !== undefined && val !== "") return val;
      }
    }
    return "";
  }

  // -------------- Builder ---------------- //
  /**
   * Detecta automáticamente todas las columnas de tipo T20XX (ej: T2023, T2022, T2021...)
   * y las retorna ordenadas descendentemente por año.
   */
  function detectYearColumns(headerIndex) {
    const yearColumns = [];
    for (const colName in headerIndex) {
      // Buscar columnas que coincidan con T + 4 dígitos (T2023, T2022, T1999, etc.)
      const match = colName.match(/^T(\d{4})$/);
      if (match) {
        const year = parseInt(match[1], 10);
        const colIndex = headerIndex[colName];
        yearColumns.push({ year, colName, colIndex });
      }
    }
    // Ordenar descendentemente por año (más reciente primero)
    yearColumns.sort((a, b) => b.year - a.year);
    return yearColumns;
  }

  function buildGeneroGerentes() {
    const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
    const detail = [];
    const yearTotals = new Map(); // key anio|tam -> {F:count, M:count}
    const seenRecords = new Set(); // Deduplicación RUC+AÑO+TAMAÑO+GÉNERO

    // Detectar columnas T202x dinámicamente
    const yearColumns = detectYearColumns(base.headerIndex);

    base.rows.forEach((row) => {
      // FILTRO #1: Solo contar filas con cargo "GERENTE"
      const cargo = normalizeLabel(getVal(row, base.headerIndex, CARGO_ALIASES));
      if (!cargo || !cargo.includes("GERENTE")) return;

      const genero = normalizeGender(getVal(row, base.headerIndex, GENDER_ALIASES));
      if (!genero) return;

      let tam = "";
      let anio = "";

      // Cascada dinámica: recorrer columnas T202x en orden descendente (más reciente primero)
      for (const { year, colName, colIndex } of yearColumns) {
        if (colIndex !== undefined && row[colIndex]) {
          tam = sizeFromCode(row[colIndex]);
          anio = year.toString();
          break; // Tomar el año más reciente que tenga datos
        }
      }

      // Fallback: si no se encontró en ninguna columna T202x
      if (!tam || !anio) {
        tam = normalizeTamano(getVal(row, base.headerIndex, TAMANO_ALIASES));
        anio = getYear(getVal(row, base.headerIndex, FECHA_ALIASES));
      }

      // FILTRO #2: Descartar filas sin año válido
      if (!anio || anio === "DESCONOCIDO") return;

      // Si no hay tamaño pero sí hay año válido, marcar como SIN_TAMANO
      if (!tam) {
        tam = "SIN_TAMANO";
      }

      const ruc = padRuc13(getVal(row, base.headerIndex, ["RUC"]));
      const razon = normalizeName(getVal(row, base.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]));

      // FILTRO #3: Deduplicación - evitar contar múltiples veces mismo RUC+AÑO+TAMAÑO+GÉNERO
      const dedupKey = `${ruc}|${anio}|${tam}|${genero}`;
      if (seenRecords.has(dedupKey)) return;
      seenRecords.add(dedupKey);

      detail.push([anio, tam, genero, ruc, razon, SHEETS.BASE]);

      const k = `${anio}||${tam}`;
      const current = yearTotals.get(k) || { F: 0, M: 0 };
      if (genero === "FEMENINO") current.F += 1;
      if (genero === "MASCULINO") current.M += 1;
      yearTotals.set(k, current);
    });

    // Global por año
    const globalTotals = new Map();
    yearTotals.forEach((val, key) => {
      const [anio] = key.split("||");
      const gKey = `${anio}||GLOBAL`;
      const current = globalTotals.get(gKey) || { F: 0, M: 0 };
      current.F += val.F;
      current.M += val.M;
      globalTotals.set(gKey, current);
    });
    globalTotals.forEach((val, key) => yearTotals.set(key, val));

    // Hoja detalle
    const shDetail = getSheetOrCreate(SHEETS.OUT_DETAIL);
    shDetail.clear();
    const headersDetail = ["ANIO", "TAMANO", "GENERO", "RUC", "RAZON_SOCIAL", "FUENTE"];
    const outDetail = detail.length ? [headersDetail, ...detail] : [headersDetail];
    shDetail.getRange(1, 1, outDetail.length, headersDetail.length).setValues(outDetail);
    shDetail.autoResizeColumns(1, headersDetail.length);

    // Pivot largo
    const pivotRows = [];
    yearTotals.forEach((val, key) => {
      const [anio, tam] = key.split("||");
      const total = val.F + val.M;
      if (total === 0) return;
      const pctF = total ? val.F / total : 0;
      const pctM = total ? val.M / total : 0;
      pivotRows.push([anio, tam, "FEMENINO", val.F, total, pctF, TAMANO_ORDER[tam] || 99]);
      pivotRows.push([anio, tam, "MASCULINO", val.M, total, pctM, TAMANO_ORDER[tam] || 99]);
    });
    pivotRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      if (a[6] !== b[6]) return a[6] - b[6];
      if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
      return 0;
    });
    const shPivot = getSheetOrCreate(SHEETS.OUT_PIVOT);
    shPivot.clear();
    const headersPivot = ["ANIO", "TAMANO", "GENERO", "GERENTES", "TOTAL_TAMANO", "PCT_GENERO", "ORDEN_TAMANO"];
    const outPivot = pivotRows.length ? [headersPivot, ...pivotRows] : [headersPivot];
    shPivot.getRange(1, 1, outPivot.length, headersPivot.length).setValues(outPivot);
    shPivot.autoResizeColumns(1, headersPivot.length);
    // FORMATO #4: Aplicar formato de porcentaje a columna PCT_GENERO
    if (pivotRows.length > 0) {
      shPivot.getRange(2, 6, pivotRows.length, 1).setNumberFormat("0.00%");
    }

    // Tabla wide para uso directo en gráficos
    const wideMap = new Map(); // key anio||tam -> {F, M, total, pctF, pctM, orden}
    pivotRows.forEach((r) => {
      const [anio, tam, genero, count, total, pct, orden] = r;
      const key = `${anio}||${tam}`;
      const current = wideMap.get(key) || { F: 0, M: 0, total: 0, orden };
      if (genero === "FEMENINO") {
        current.F = count;
      } else if (genero === "MASCULINO") {
        current.M = count;
      }
      current.total = total;
      current.orden = orden;
      wideMap.set(key, current);
    });
    const wideRows = [];
    wideMap.forEach((val, key) => {
      const [anio, tam] = key.split("||");
      const total = val.total || val.F + val.M;
      const pctF = total ? val.F / total : 0;
      const pctM = total ? val.M / total : 0;
      wideRows.push([anio, tam, val.F, val.M, total, pctF, pctM, val.orden]);
    });
    wideRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      if (a[7] !== b[7]) return a[7] - b[7];
      return a[1].localeCompare(b[1]);
    });
    const shWide = getSheetOrCreate(SHEETS.OUT_WIDE);
    shWide.clear();
    const headersWide = ["ANIO", "TAMANO", "FEMENINO", "MASCULINO", "TOTAL", "PCT_FEMENINO", "PCT_MASCULINO", "ORDEN_TAMANO"];
    const outWide = wideRows.length ? [headersWide, ...wideRows] : [headersWide];
    shWide.getRange(1, 1, outWide.length, headersWide.length).setValues(outWide);
    shWide.autoResizeColumns(1, headersWide.length);
    // FORMATO #4: Aplicar formato de porcentaje a columnas PCT_FEMENINO y PCT_MASCULINO
    if (wideRows.length > 0) {
      shWide.getRange(2, 6, wideRows.length, 2).setNumberFormat("0.00%");
    }
  }

  // -------------- Entrypoint -------------- //
  function refreshDashboardGenero() {
    buildGeneroGerentes();
    Logger.log("Tablas de genero regeneradas.");
  }

  function onOpenDash3() {
    SpreadsheetApp.getUi().createMenu("DASH3").addItem("Generar tablas genero", "refreshDashboardGenero").addToUi();
  }

  // Trigger manual via checkbox (hoja REPORTE_1, celda Q105)
  const DASH3_TRIGGER_SHEET = "REPORTE_1";
  const DASH3_TRIGGER_CELL = "Q105";

  function onEditDash3(e) {
    const range = e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== DASH3_TRIGGER_SHEET) return;
    if (range.getA1Notation() !== DASH3_TRIGGER_CELL) return;

    const val = range.getValue();
    if (val === true) {
      const ss = SpreadsheetApp.getActive();
      ss.toast("Actualizando dashboard…");
      try {
        refreshDashboardGenero();
        ss.toast("Dashboard listo.");
        sheet.getRange("Q106").setValue("Ultima actualizacion: " + new Date());
      } finally {
        range.setValue(false); // deja la casilla lista para el siguiente clic
      }
    }
  }

  // Expose globals
  global.refreshDashboardGenero = refreshDashboardGenero;
  global.onOpenDash3 = onOpenDash3;
  global.onEditDash3 = onEditDash3;
})(this);

// Wrappers visibles para el menú de Apps Script si el editor no detecta las asignaciones dinámicas.
function refreshDashboardGeneroWrapper() {
  return refreshDashboardGenero();
}
function onOpenDash3Wrapper() {
  return onOpenDash3();
}
// Wrapper para que el trigger instalable detecte la función en el selector
function onEditDash3Wrapper(e) {
  return onEditDash3(e);
}

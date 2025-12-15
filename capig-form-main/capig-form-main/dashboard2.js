/**
 * Tablas para el dashboard de cambios de tamano (crecimiento/decrecimiento).
 * Ejecutar refreshDashboardTamano() para regenerar los pre-aggregados.
 */
const Dash2Module = (function () {
  // ------------------ Config ------------------ //
  const DASH2_SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    REGISTRO: "SOCIOS",
    REGISTRO_FALLBACK: "REGISTRO_AFILIADO",
    VENTAS: "VENTAS_SOCIO",
    VENTAS_FALLBACK: "VENTAS_AFILIADOS",
    OUT_DETAIL: "DASH_TAMANO_ANIO",
    OUT_TRANSITIONS: "PIVOT_CAMBIO_TAMANO_ANIO",
    OUT_2023: "DASH_CAMBIO_TAMANO_2023",
    OUT_TRANSITIONS_SUMMARY: "DASH_TRANSICIONES",
  };

  const TAMANO_BY_CODE = { 1: "MICRO", 2: "PEQUENA", 3: "MEDIANA", 4: "GRANDE" };
  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4 };
  const TAMANO_BY_MONTO = [
    { max: 100000, label: "MICRO" },
    { max: 1000000, label: "PEQUENA" },
    { max: 5000000, label: "MEDIANA" },
    { max: Number.POSITIVE_INFINITY, label: "GRANDE" },
  ];

  const FIELD_ALIASES = {
    RUC: ["RUC", "NUMERO_RUC", "NUM_RUC"],
    RAZON: ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "NOMBRE", "RAZON_SOCIA"],
    ALT_ID: [
      "ID_UNICO",
      "ID UNICO",
      "ID_INTERNO",
      "ID INTERNO",
      "ID",
      "ID_SOCIO",
      "CODIGO_SOCIO",
      "CODIGO",
      "CLAVE",
      "CLAVE_UNICA",
      // Soporte para columnas tipo "No." que ya vienen numeradas en la base
      "NO",
      "NO.",
      "NRO",
      "NUM",
      "NUMERO",
    ],
    TAMANO: ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMAï¿½'O", "TAMANO_EMP"],
    SECTOR: ["SECTOR", "SECTOR "],
    FECHA_AF: ["FECHA_AFILIACION", "FECHA AFILIACION", "FECHA_INGRESO", "FECHA DE INGRESO", "FECHA_REGISTRO"],
    VENTAS: ["VENTAS", "VENTAS_ANUAL", "VENTAS_ANUALES", "VENTAS_MONT_EST", "MONTO_ESTIMADO", "MONTO_TOTAL", "VALOR TOTAL"],
    ANIO: ["ANIO", "ANO", "Aï¿½'O", "ANIO_VENTA", "ANO_VENTA", "Aï¿½'O_VENTA", "Aï¿½'O VENTA"],
  };

  // ------------------ Utils ------------------ //
  function normalizeLabel(label) {
    let txt = (label || "").toString().trim().toUpperCase();
    txt = txt.replace(/ï¿½ï¿½ï¿½/g, "N");
    txt = txt.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    txt = txt.replace(/[^A-Z0-9]+/g, "_");
    txt = txt.replace(/^_+|_+$/g, "");
    return txt;
  }
  function normalizeName(val) {
    return (val || "").toString().trim().toUpperCase().replace(/\s+/g, " ");
  }
  function normalizeKey(val) {
    return (val || "").toString().trim().toUpperCase().replace(/\s+/g, "_");
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
  function toNumber(val) {
    if (val === null || val === undefined || val === "") return 0;
    let str = val.toString().trim().replace(/\s/g, "");
    const hasComma = str.indexOf(",") !== -1;
    const hasDot = str.indexOf(".") !== -1;
    if (hasComma && hasDot) {
      str = str.replace(/\./g, "").replace(/,/g, ".");
    } else if (hasComma) {
      str = str.replace(/,/g, ".");
    } else if (hasDot) {
      const dotCount = (str.match(/\./g) || []).length;
      if (dotCount > 1) str = str.replace(/\./g, "");
    }
    str = str.replace(/[^0-9\.-]/g, "");
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
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
  function sizeFromCode(val) {
    const code = (val || "").toString().trim();
    if (TAMANO_BY_CODE[code]) return TAMANO_BY_CODE[code];
    return normalizeTamano(val);
  }
  function entityKey(headerIndex, row) {
    const ruc = padRuc13(getVal(row, headerIndex, FIELD_ALIASES.RUC));
    const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
    const altId = normalizeKey(getVal(row, headerIndex, FIELD_ALIASES.ALT_ID || []));
    if (altId) {
      if (ruc) return `${ruc}__${altId}`;
      return `ID__${altId}`;
    }
    return ruc || razon || "";
  }
  // ------------------ Readers ------------------ //
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

  function splitBlocksByHeader(values) {
    const headerRows = [];
    for (let i = 0; i < values.length; i++) {
      const norm = values[i].map(normalizeLabel);
      const hasRuc = norm.includes("RUC");
      const hasRazon = norm.includes("RAZON_SOCIAL");
      const nonEmpty = norm.filter((c) => c).length;
      if ((hasRuc || hasRazon) && nonEmpty >= 2) headerRows.push(i);
    }
    const blocks = [];
    headerRows.forEach((idx, pos) => {
      const next = pos + 1 < headerRows.length ? headerRows[pos + 1] : values.length;
      const headerNorm = values[idx].map(normalizeLabel);
      const headerIndex = {};
      headerNorm.forEach((h, colIdx) => {
        if (h && headerIndex[h] === undefined) headerIndex[h] = colIdx;
      });
      blocks.push({ headerIndex, rows: values.slice(idx + 1, next) });
    });
    return blocks;
  }

  // ------------------ Builders ------------------ //
  function collectTamanoByYear() {
    const records = new Map(); // key => {ruc, razon, year, tamano, fuente, priority}
    function addRecord({ ruc, razon, year, tamano, fuente, priority }) {
      if (!year || !tamano) return;
      const keyEntity = ruc || normalizeName(razon);
      if (!keyEntity) return;
      const recKey = `${keyEntity}||${year}`;
      const existing = records.get(recKey);
      if (existing && existing.priority <= priority) return;
      records.set(recKey, { ruc, razon, year, tamano, fuente, priority });
    }

    // BASE (SOCIOS) con dos bloques
    const ss = SpreadsheetApp.getActive();
    const baseSheetName = [DASH2_SHEETS.BASE, DASH2_SHEETS.BASE_FALLBACK].find((n) => n && ss.getSheetByName(n));
    const shBase = baseSheetName ? ss.getSheetByName(baseSheetName) : null;
    if (shBase) {
      const values = shBase.getDataRange().getDisplayValues();
      const blocks = splitBlocksByHeader(values);
      blocks.forEach((block) => {
        const { headerIndex, rows } = block;
        const yearCols = Object.keys(headerIndex).filter((k) => /^T_?\d{4}$/.test(k));
        rows.forEach((row) => {
          const ruc = padRuc13(getVal(row, headerIndex, FIELD_ALIASES.RUC));
          const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
          yearCols.forEach((colKey) => {
            const val = row[headerIndex[colKey]];
            const tam = sizeFromCode(val);
            const year = colKey.replace(/^T_?/, "");
            if (tam) addRecord({ ruc, razon, year, tamano: tam, fuente: baseSheetName || DASH2_SHEETS.BASE, priority: 1 });
          });
          const tamExplicit = normalizeTamano(getVal(row, headerIndex, FIELD_ALIASES.TAMANO));
          const yearAf = getYear(getVal(row, headerIndex, FIELD_ALIASES.FECHA_AF));
          if (tamExplicit && yearAf) {
            addRecord({ ruc, razon, year: yearAf, tamano: tamExplicit, fuente: baseSheetName || DASH2_SHEETS.BASE, priority: 2 });
          }
        });
      });
    }

    // REGISTRO_AFILIADO (si existe) usa FECHA_AFILIACION para asignar anio
    const registro = readTableFlexibleByCandidates([DASH2_SHEETS.REGISTRO, DASH2_SHEETS.REGISTRO_FALLBACK]);
    registro.rows.forEach((row) => {
      const ruc = padRuc13(getVal(row, registro.headerIndex, FIELD_ALIASES.RUC));
      const razon = normalizeName(getVal(row, registro.headerIndex, FIELD_ALIASES.RAZON));
      const tam = normalizeTamano(getVal(row, registro.headerIndex, FIELD_ALIASES.TAMANO));
      const year = getYear(getVal(row, registro.headerIndex, FIELD_ALIASES.FECHA_AF));
      if (tam && year) addRecord({ ruc, razon, year, tamano: tam, fuente: DASH2_SHEETS.REGISTRO, priority: 3 });
    });

    // VENTAS_AFILIADOS: clasifica por monto o tamano declarado
    const ventas = readTableFlexibleByCandidates([DASH2_SHEETS.VENTAS, DASH2_SHEETS.VENTAS_FALLBACK]);
    ventas.rows.forEach((row) => {
      const ruc = padRuc13(getVal(row, ventas.headerIndex, FIELD_ALIASES.RUC));
      const razon = normalizeName(getVal(row, ventas.headerIndex, FIELD_ALIASES.RAZON));
      const year = getVal(row, ventas.headerIndex, FIELD_ALIASES.ANIO) || getYear(getVal(row, ventas.headerIndex, FIELD_ALIASES.FECHA_AF));
      let tam = normalizeTamano(getVal(row, ventas.headerIndex, FIELD_ALIASES.TAMANO));
      if (!tam) {
        const monto = toNumber(getVal(row, ventas.headerIndex, FIELD_ALIASES.VENTAS));
        if (monto > 0) {
          for (const rule of TAMANO_BY_MONTO) {
            if (monto <= rule.max) {
              tam = rule.label;
              break;
            }
          }
        }
      }
      if (tam && year) addRecord({ ruc, razon, year, tamano: tam, fuente: DASH2_SHEETS.VENTAS, priority: 4 });
    });

    const detailRows = Array.from(records.values()).map((r) => [
      r.ruc,
      r.razon,
      r.year,
      r.tamano,
      r.fuente,
    ]);
    detailRows.sort((a, b) => {
      if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
      if (a[3] !== b[3]) return a[3].localeCompare(b[3]);
      return (a[0] || "").localeCompare(b[0] || "");
    });

    const shDetail = getSheetOrCreate(DASH2_SHEETS.OUT_DETAIL);
    shDetail.clear();
    const headers = ["RUC", "RAZON_SOCIAL", "ANIO", "TAMANO", "FUENTE"];
    const out = detailRows.length ? [headers, ...detailRows] : [headers];
    shDetail.getRange(1, 1, out.length, headers.length).setValues(out);
    shDetail.autoResizeColumns(1, headers.length);

    return detailRows;
  }

  function buildTransitions(detailRows) {
    const byEntity = new Map();
    detailRows.forEach((row) => {
      const ruc = row[0];
      const razon = row[1];
      const year = parseInt(row[2], 10);
      const tam = row[3];
      if (!year || !tam) return;
      const key = ruc || normalizeName(razon);
      if (!key) return;
      if (!byEntity.has(key)) byEntity.set(key, []);
      byEntity.get(key).push({ year, tam, ruc, razon });
    });

    const transitions = [];
    byEntity.forEach((hist, key) => {
      hist.sort((a, b) => a.year - b.year);
      for (let i = 0; i < hist.length - 1; i++) {
        const a = hist[i];
        const b = hist[i + 1];
        if (!a.tam || !b.tam || a.tam === b.tam) continue;
        transitions.push({
          entity: key,
          anioInicial: a.year.toString(),
          anioFinal: b.year.toString(),
          tamInicial: a.tam,
          tamFinal: b.tam,
          delta: (TAMANO_ORDER[b.tam] || 0) - (TAMANO_ORDER[a.tam] || 0),
        });
      }
    });

    const originTotals = new Map(); // key anio|tam -> total de empresas con ese tamano en el anio
    detailRows.forEach((row) => {
      const year = row[2];
      const tam = row[3];
      if (!year || !tam) return;
      const key = `${year}||${tam}`;
      originTotals.set(key, (originTotals.get(key) || 0) + 1);
    });

    const agg = new Map(); // key anio_i|anio_f|tam_i|tam_f -> count de transiciones
    transitions.forEach((t) => {
      const key = `${t.anioInicial}||${t.anioFinal}||${t.tamInicial}||${t.tamFinal}`;
      agg.set(key, (agg.get(key) || 0) + 1);
    });

    const pivotRows = Array.from(agg.entries()).map(([k, count]) => {
      const [anioIni, anioFin, tamIni, tamFin] = k.split("||");
      const originTotal = originTotals.get(`${anioIni}||${tamIni}`) || 0;
      const pct = originTotal ? count / originTotal : 0;
      const delta = (TAMANO_ORDER[tamFin] || 0) - (TAMANO_ORDER[tamIni] || 0);
      const direccion = delta > 0 ? "CRECIMIENTO" : "DECRECIMIENTO";
      const ordenIni = TAMANO_ORDER[tamIni] || 99;
      const ordenFin = TAMANO_ORDER[tamFin] || 99;
      return [anioIni, anioFin, tamIni, tamFin, ordenIni, ordenFin, count, pct, direccion, delta];
    });

    pivotRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      if (a[1] !== b[1]) return a[1].localeCompare(b[1]);
      if (a[4] !== b[4]) return a[4] - b[4]; // ORDEN_INICIAL
      if (a[5] !== b[5]) return a[5] - b[5]; // ORDEN_FINAL
      if (a[8] !== b[8]) return a[8].localeCompare(b[8]); // DIRECCION
      if (a[9] !== b[9]) return a[9] - b[9]; // DELTA
      return a[2].localeCompare(b[2]); // TAMANO_INICIAL
    });

    const shPivot = getSheetOrCreate(DASH2_SHEETS.OUT_TRANSITIONS);
    shPivot.clear();
    const headers = [
      "ANIO_INICIAL",
      "ANIO_FINAL",
      "TAMANO_INICIAL",
      "TAMANO_FINAL",
      "ORDEN_INICIAL",
      "ORDEN_FINAL",
      "EMPRESAS",
      "PCT_ORIGEN",
      "DIRECCION",
      "DELTA",
    ];
    const out = pivotRows.length ? [headers, ...pivotRows] : [headers];
    shPivot.getRange(1, 1, out.length, headers.length).setValues(out);
    shPivot.autoResizeColumns(1, headers.length);

    const rows2023 = pivotRows.filter((r) => r[0] === "2022" && r[1] === "2023" && Math.abs(r[9]) === 1);
    const sh2023 = getSheetOrCreate(DASH2_SHEETS.OUT_2023);
    sh2023.clear();
    const headers2023 = [
      "ANIO_INICIAL",
      "ANIO_FINAL",
      "TAMANO_INICIAL",
      "TAMANO_FINAL",
      "ORDEN_INICIAL",
      "ORDEN_FINAL",
      "EMPRESAS",
      "PCT_ORIGEN",
      "DIRECCION",
      "DELTA",
    ];
    const out2023 = rows2023.length ? [headers2023, ...rows2023] : [headers2023];
    sh2023.getRange(1, 1, out2023.length, headers2023.length).setValues(out2023);
    sh2023.autoResizeColumns(1, headers2023.length);

    // VersiÃ³n completa (incluye saltos de mÃ¡s de 1 nivel) para referencias tipo PPT.
    const rows2023Full = pivotRows.filter((r) => r[0] === "2022" && r[1] === "2023");
    const sh2023Full = getSheetOrCreate("DASH_CAMBIO_TAMANO_2023_FULL");
    sh2023Full.clear();
    const headers2023Full = [...headers2023, "FLUJO"];
    const out2023Full = rows2023Full.length
      ? [headers2023Full, ...rows2023Full.map((r) => [...r, `${r[2]} -> ${r[3]}`])]
      : [headers2023Full];
    sh2023Full.getRange(1, 1, out2023Full.length, headers2023Full[0] ? headers2023Full.length : 0).setValues(out2023Full);
    sh2023Full.autoResizeColumns(1, headers2023Full.length);

    // Tabla plana para vistas rÃ¡pidas por direcciÃ³n (compatible con slicer de ANIO_INICIAL).
    const summaryRows = pivotRows.map((r) => [
      r[0], // ANIO_INICIAL
      `${r[2]} -> ${r[3]}`, // TRANSICION
      r[2], // TAMANO_INICIAL
      r[3], // TAMANO_FINAL
      r[4], // ORDEN_INICIAL
      r[5], // ORDEN_FINAL
      r[6], // EMPRESAS
      r[7], // PCT_ORIGEN
      r[8], // DIRECCION
      r[9], // DELTA
      `${r[6]} (${(r[7] * 100).toFixed(2)}%)`, // ETIQUETA_FINAL (cantidad + porcentaje)
    ]);
    summaryRows.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      if (a[8] !== b[8]) return a[8].localeCompare(b[8]); // CRECIMIENTO/DECRECIMIENTO
      if (a[4] !== b[4]) return a[4] - b[4];
      if (a[5] !== b[5]) return a[5] - b[5];
      return a[1].localeCompare(b[1]);
    });
    const shSummary = getSheetOrCreate(DASH2_SHEETS.OUT_TRANSITIONS_SUMMARY);
    shSummary.clear();
    const headersSummary = [
      "ANIO_INICIAL",
      "TRANSICION",
      "TAMANO_INICIAL",
      "TAMANO_FINAL",
      "ORDEN_INICIAL",
      "ORDEN_FINAL",
      "EMPRESAS",
      "PCT_ORIGEN",
      "DIRECCION",
      "DELTA",
      "ETIQUETA_FINAL",
    ];
    const outSummary = summaryRows.length ? [headersSummary, ...summaryRows] : [headersSummary];
    shSummary.getRange(1, 1, outSummary.length, headersSummary.length).setValues(outSummary);
    shSummary.autoResizeColumns(1, headersSummary.length);
  }

  return { collectTamanoByYear, buildTransitions };
})();

// ------------------ Entrypoint (global) ------------------ //
function refreshDashboardTamano() {
  const detail = Dash2Module.collectTamanoByYear();
  Dash2Module.buildTransitions(detail);
  Logger.log("Tablas de tamano regeneradas.");
}

function onOpenDash2() {
  SpreadsheetApp.getUi().createMenu("DASH2").addItem("Generar tablas tamano", "refreshDashboardTamano").addToUi();
}

// Trigger manual via checkbox (hoja REPORTE_1, celda Q53)
const DASH2_TRIGGER_SHEET = "REPORTE_1";
const DASH2_TRIGGER_CELL = "Q53";

function onEditDash2(e) {
  const range = e.range;
  if (!range) return;
  const sheet = range.getSheet();
  if (!sheet || sheet.getName() !== DASH2_TRIGGER_SHEET) return;
  if (range.getA1Notation() !== DASH2_TRIGGER_CELL) return;

  const val = range.getValue();
  if (val === true) {
    const ss = SpreadsheetApp.getActive();
    ss.toast("Actualizando dashboardâ€¦");
    try {
      refreshDashboardTamano();
      ss.toast("Dashboard listo.");
      sheet.getRange("Q54").setValue("Ultima actualizacion: " + new Date());
    } finally {
      range.setValue(false); // deja la casilla lista para el siguiente clic
    }
  }
}

// Wrapper opcional para asignar a un botÃ³n/imagen
function refreshDashboardTamanoWrapper() {
  return refreshDashboardTamano();
}







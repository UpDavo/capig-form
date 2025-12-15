/**
 * Dashboard de Capacitaciones - DASH5
 * Prepara tablas pre-agregadas (planas) combinando histrico + registros Django.
 * Ejecutar refreshDashboardCapacitaciones() para regenerar las tablas.
 */
(function (global) {
  const SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    CAP_HIST: "CAPACITACIONES_HISTORICAS",
    CAP_HIST_FALLBACK: "CAPACITACIONES_HISTORICO",
    CAP_NEW: "CAPACITACIONES",
    CAP_NEW_FALLBACK: "CAPACITACIONES_FINAL",
    OUT_RESUMEN: "PIVOT_CAPACITACIONES_RESUMEN_ANIO",
    OUT_RESUMEN_SOCIOS: "PIVOT_CAPACITACIONES_RESUMEN_ANIO_SOCIOS",
    OUT_TOP: "PIVOT_CAPACITACIONES_TOP_EMPRESAS",
    OUT_SOCIOS: "PIVOT_CAPACITACIONES_SOCIOS",
    OUT_SOCIOS_TAM: "PIVOT_CAPACITACIONES_SOCIOS_TAMANO",
    OUT_MASTER: "PIVOT_CAPACITACIONES_MASTER",
    OUT_MASTER_SLICER: "DASH5_MAESTRA",
  };

  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, GLOBAL: 0, SIN_TAMANO: 99, DESCONOCIDO: 99 };
  const ALT_ID_ALIASES = [
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
    "NO",
    "NO.",
    "NRO",
    "N°",
    "NUM",
    "NUMERO",
  ];
  const CAP_TRIMESTRE_ALIASES = [
    ["1ER_TRIMESTRE", "1ER TRIMESTRE", "PRIMER_TRIMESTRE", "PRIMER TRIMESTRE", "Q1_CAPACITACIONES", "Q1"],
    ["2DO_TRIMESTRE", "2DO TRIMESTRE", "SEGUNDO_TRIMESTRE", "SEGUNDO TRIMESTRE", "Q2_CAPACITACIONES", "Q2"],
    ["3ER_TRIMESTRE", "3ER TRIMESTRE", "TERCER_TRIMESTRE", "TERCER TRIMESTRE", "Q3_CAPACITACIONES", "Q3"],
    ["4TO_TRIMESTRE", "4TO TRIMESTRE", "CUARTO_TRIMESTRE", "CUARTO TRIMESTRE", "Q4_CAPACITACIONES", "Q4"],
  ];
  const VALOR_TRIMESTRE_ALIASES = [
    ["VALOR_1ER", "VALOR 1ER", "VALOR_1ER_TRIMESTRE", "VALOR 1ER TRIMESTRE", "Q1_VALOR", "VALOR_Q1"],
    ["VALOR_2DO", "VALOR 2DO", "VALOR_2DO_TRIMESTRE", "VALOR 2DO TRIMESTRE", "Q2_VALOR", "VALOR_Q2"],
    ["VALOR_3ER", "VALOR 3ER", "VALOR_3ER_TRIMESTRE", "VALOR 3ER TRIMESTRE", "Q3_VALOR", "VALOR_Q3"],
    ["VALOR_4TO", "VALOR 4TO", "VALOR_4TO_TRIMESTRE", "VALOR 4TO TRIMESTRE", "Q4_VALOR", "VALOR_Q4"],
  ];

  // ---------- Utils ----------
  function normalizeLabel(label) {
    let txt = (label || "").toString().trim().toUpperCase();
    txt = txt.replace(/\u00d1|\u00f1/g, "N");
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

  function parseMonto(val) {
    if (val === null || val === undefined || val === "") return 0;
    if (val === "-" || val === " - ") return 0;
    if (typeof val === "number") return isNaN(val) ? 0 : val;
    let str = val.toString().trim().replace(/\s/g, "");
    const lastComma = str.lastIndexOf(",");
    const lastDot = str.lastIndexOf(".");

    if (lastComma !== -1 && lastDot !== -1) {
      if (lastDot > lastComma) {
        str = str.replace(/,/g, "");
      } else {
        str = str.replace(/\./g, "").replace(/,/g, ".");
      }
    } else if (lastComma !== -1) {
      str = str.replace(/,/g, ".");
    } else if (lastDot !== -1) {
      const dotCount = (str.match(/\./g) || []).length;
      if (dotCount > 1) {
        str = str.replace(/\./g, "");
      }
    }

    str = str.replace(/[^0-9.-]/g, "");
    const n = parseFloat(str);
    return isNaN(n) ? 0 : n;
  }

  function parseDateFlexible(val) {
    if (!val) return null;
    let d = new Date(val);
    if (!isNaN(d.getTime())) return d;
    const str = val.toString().trim();
    const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
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

  function formatNumberColumn(sheet, startRow, startCol, numRows, formatPattern) {
    if (numRows <= 0) return;
    sheet.getRange(startRow, startCol, numRows, 1).setNumberFormat(formatPattern);
  }

  const MONEY_FMT = '"$"#,##0.00';

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

  function sumNumbers(row, headerIndex, aliasGroups, parserFn) {
    let total = 0;
    aliasGroups.forEach((aliases) => {
      const v = getVal(row, headerIndex, aliases);
      const n = parserFn ? parserFn(v) : parseFloat(v);
      if (!isNaN(n) && n > 0) total += n;
    });
    return total;
  }

  // ---------- Builder DASH5 ----------
  function buildCapacitacionesDashboard() {
    Logger.log("[DASH5] Iniciando preparacion de tablas de capacitaciones...");

    const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
    let capHist = readTableFlexibleByCandidates([SHEETS.CAP_HIST, SHEETS.CAP_HIST_FALLBACK]);
    let capNew = readTableFlexibleByCandidates([SHEETS.CAP_NEW, SHEETS.CAP_NEW_FALLBACK]);

    if (capHist.name && capNew.name && capHist.name === capNew.name) {
      Logger.log(`[DASH5] CAP_HIST y CAP_NEW apuntan a la misma hoja (${capNew.name}); se procesara solo como CAP_NEW para evitar duplicados.`);
      capHist = { headerIndex: {}, rows: [], name: capHist.name };
    }

    // Mapear base: RUC/ALT_ID/RAZON -> {tamano, razon, anioAfiliacion}
    const baseMap = new Map();
    const razonToBase = new Map();
    const altIdToBase = new Map();
    base.rows.forEach((row) => {
      const ruc = padRuc13(getVal(row, base.headerIndex, ["RUC"]));
      const altId = normalizeKey(getVal(row, base.headerIndex, ALT_ID_ALIASES));
      if (!ruc && !altId) return;
      const razon = normalizeName(getVal(row, base.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]));
      const tam =
        normalizeTamano(getVal(row, base.headerIndex, ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMAO", "TAMA\u00d1O"])) ||
        "SIN_TAMANO";
      const anioAf = getYear(getVal(row, base.headerIndex, ["FECHA_AFILIACION", "FECHA AFILIACION", "FECHA DE INGRESO"]));
      const data = { razon, tamano: tam || "SIN_TAMANO", anioAfiliacion: anioAf };
      if (ruc) baseMap.set(ruc, data);
      if (altId) altIdToBase.set(altId, data);
      if (razon) razonToBase.set(razon, ruc || altId || razon);
    });
    Logger.log(`[DASH5] Empresas en BASE: ${baseMap.size}`);

    // Helper para obtener identidad empresa
    function resolveEmpresaByName(name, rowForAltId, rowHeaderIndex) {
      const razon = normalizeName(name);
      let ruc = razonToBase.get(razon) || "";
      let info = ruc ? baseMap.get(ruc) : null;
      if (!info && rowForAltId) {
        const altId = normalizeKey(getVal(rowForAltId, rowHeaderIndex || {}, ALT_ID_ALIASES));
        if (altId) {
          if (altIdToBase.has(altId)) {
            info = altIdToBase.get(altId);
          }
          if (!ruc && altId) ruc = altId; // clave auxiliar para dedupe
        }
      }
      const tam = info ? info.tamano : "";
      const anioAfiliacion = info ? info.anioAfiliacion : "";
      return { razon, ruc, tam, anioAfiliacion };
    }

    const registros = [];
    const duplicados = [];
    const dedupSet = new Set();

    // Determinar ao de referencia para histrico: tomamos el mximo ao presente en CAP_NEW; si no hay, el ao actual.
    let anioHistRef = "";
    capNew.rows.forEach((row) => {
      const fecha = getVal(row, capNew.headerIndex, ["FECHA"]);
      const anio = getYear(fecha);
      if (anio) {
        if (!anioHistRef || anio > anioHistRef) anioHistRef = anio;
      }
    });
    if (!anioHistRef) {
      anioHistRef = new Date().getFullYear().toString();
    }

    // Procesar historico CAPACITACIONES (usa fecha real si existe, si no ano afiliacion/ref)
    capHist.rows.forEach((row, idx) => {
      const razonRaw = getVal(row, capHist.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]);
      const razonNorm = normalizeName(razonRaw);
      const rucHist = padRuc13(getVal(row, capHist.headerIndex, ["RUC"]));
      const fechaHistRaw = getVal(row, capHist.headerIndex, ["FECHA", "FECHA_CAP", "FECHA_CAPACITACION", "FECHA CAP"]);
      const { razon, ruc: rucBase, tam, anioAfiliacion } = resolveEmpresaByName(razonRaw, row, capHist.headerIndex);
      const ruc = rucHist || rucBase;
      const key = ruc || razon;
      if (!key) return;
      let capCount = parseFloat(getVal(row, capHist.headerIndex, ["TOTAL CAPAC.", "TOTAL_CAPAC", "TOTAL_CAPAC.", "TOTAL_CAPACITACIONES"])) || 0;
      if (!capCount || capCount <= 0) {
        capCount = sumNumbers(row, capHist.headerIndex, CAP_TRIMESTRE_ALIASES);
      }
      let valor = parseMonto(getVal(row, capHist.headerIndex, ["VALOR TOTAL", "VALOR TOTAL ", "VALOR_TOTAL", "VALOR"]));
      if (!valor || valor <= 0) {
        valor = sumNumbers(row, capHist.headerIndex, VALOR_TRIMESTRE_ALIASES, parseMonto);
      }
      if (capCount <= 0 && valor <= 0) return;
      const tamanoHist = normalizeTamano(getVal(row, capHist.headerIndex, ["TAMANO", "TAMAO", "TAMANIO", "TAMA\u00d1O"]));
      const anioFecha = getYear(fechaHistRaw);
      // Prioriza fecha real; luego anio de referencia (max de CAP_NEW); ultimo: anio de afiliacion.
      const anio = anioFecha || anioHistRef || anioAfiliacion;
      const dedupKey = `${anio}|${key}|${fechaHistRaw || "HIST"}|HIST`;
      if (dedupSet.has(dedupKey)) {
        duplicados.push({ fuente: "HIST", anio, key, razon, ruc, fecha: fechaHistRaw, valor, capCount });
        return;
      }
      dedupSet.add(dedupKey);
      registros.push({
        anio,
        key,
        ruc,
        razon,
        tamano: tam || tamanoHist || "SIN_TAMANO",
        esSocio: !!ruc && razonNorm !== "NO SOCIOS" && razonNorm !== "NO_SOCIOS",
        capCount,
        valor,
        fuente: "HIST",
      });
      if (capCount === 0 && valor === 0 && idx < 3) {
        Logger.log(`[DASH5 DEBUG] Hist cap fila ${idx + 1} sin monto ni count: razon='${razonRaw}'`);
      }
    });
    Logger.log(`[DASH5] Registros historico procesados: ${registros.length}`);

    // Procesar CAPACITACIONES_FINAL (Django)
    capNew.rows.forEach((row, idx) => {
      const razonRaw = getVal(row, capNew.headerIndex, ["RAZON SOCIAL", "RAZON_SOCIAL", "Razon Social"]);
      const rucNew = padRuc13(getVal(row, capNew.headerIndex, ["RUC"]));
      const { razon, ruc: rucBase, tam } = resolveEmpresaByName(razonRaw, row, capNew.headerIndex);
      const razonNorm = normalizeName(razonRaw);
      const ruc = rucNew || rucBase;
      const key = ruc || razon;
      if (!key) return;
      const valor = parseMonto(getVal(row, capNew.headerIndex, ["VALOR DEL PAGO", "VALOR", "VALOR_PAGO"]));
      const fecha = getVal(row, capNew.headerIndex, ["FECHA"]);
      // Si no hay fecha, usar el año de referencia detectado (o año actual) para no perder la fila en pivots por año.
      const anio = getYear(fecha) || anioHistRef;
      const dedupKey = `${anio}|${key}|${fecha || "SIN_FECHA"}|DJANGO`;
      if (dedupSet.has(dedupKey)) {
        duplicados.push({ fuente: "DJANGO", anio, key, razon, ruc, fecha, valor, capCount: 1 });
        return;
      }
      dedupSet.add(dedupKey);
      registros.push({
        anio,
        key,
        ruc,
        razon,
        tamano: tam || "SIN_TAMANO",
        esSocio: !!ruc && razonNorm !== "NO SOCIOS" && razonNorm !== "NO_SOCIOS",
        capCount: 1,
        valor,
        fuente: "DJANGO",
      });
    });
    Logger.log(`[DASH5] Registros Django procesados: ${registros.length}`);
    if (duplicados.length) {
      Logger.log(`[DASH5 WARNING] Duplicados detectados y omitidos: ${duplicados.length}`);
    }

    if (registros.length === 0) {
      Logger.log("[DASH5 WARNING] No se generaron registros de capacitaciones.");
    }

    // ---- Agregados ----
    if (duplicados.length) {
      generateDuplicadosSheet(duplicados);
    }
    generateResumen(registros);
    generateResumenSocios(registros);
    generateTop(registros);
    generateSocios(registros);
    generateSociosTamano(registros);
    generateMaster(registros);
    generateSlicerMasterSheet();

    Logger.log("[DASH5] Tablas de capacitaciones generadas.");
  }

  function generateResumen(registros) {
    const map = new Map(); // key: anio|tam|socio
    registros.forEach((r) => {
      const k = `${r.anio}|${r.tamano}|${r.esSocio ? "1" : "0"}`;
      const cur = map.get(k) || { empresas: new Set(), cap: 0, valor: 0 };
      cur.empresas.add(r.key);
      cur.cap += r.capCount;
      cur.valor += r.valor;
      map.set(k, cur);
    });
    const rows = [];
    map.forEach((v, k) => {
      const [anio, tam, socioFlag] = k.split("|");
      rows.push([anio, tam, socioFlag === "1", v.empresas.size, v.cap, v.valor]);
    });
    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
      const oA = TAMANO_ORDER[a[1]] || 99;
      const oB = TAMANO_ORDER[b[1]] || 99;
      return oA - oB;
    });
    const header = ["ANIO", "TAMANO", "ES_SOCIO", "EMPRESAS", "CAPACITACIONES", "VALOR_TOTAL"];
    const sheet = getSheetOrCreate(SHEETS.OUT_RESUMEN);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_RESUMEN}: ${rows.length} filas`);
  }

  function generateResumenSocios(registros) {
    const filtrados = registros.filter((r) => r.esSocio);
    const map = new Map(); // key: anio|tam
    filtrados.forEach((r) => {
      const k = `${r.anio}|${r.tamano}`;
      const cur = map.get(k) || { empresas: new Set(), cap: 0, valor: 0 };
      cur.empresas.add(r.key);
      cur.cap += r.capCount;
      cur.valor += r.valor;
      map.set(k, cur);
    });
    const rows = [];
    map.forEach((v, k) => {
      const [anio, tam] = k.split("|");
      rows.push([anio, tam, v.empresas.size, v.cap, v.valor]);
    });
    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
      const oA = TAMANO_ORDER[a[1]] || 99;
      const oB = TAMANO_ORDER[b[1]] || 99;
      return oA - oB;
    });
    const header = ["ANIO", "TAMANO", "EMPRESAS", "CAPACITACIONES", "VALOR_TOTAL"];
    const sheet = getSheetOrCreate(SHEETS.OUT_RESUMEN_SOCIOS);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_RESUMEN_SOCIOS}: ${rows.length} filas`);
  }

  function generateTop(registros) {
    const perAnio = new Map();
    registros.forEach((r) => {
      const k = `${r.anio}|${r.key}`;
      const cur = perAnio.get(k) || {
        anio: r.anio,
        key: r.key,
        ruc: r.ruc,
        razon: r.razon,
        tamano: r.tamano,
        esSocio: r.esSocio,
        cap: 0,
        valor: 0,
      };
      cur.cap += r.capCount;
      cur.valor += r.valor;
      perAnio.set(k, cur);
    });

    const grouped = new Map(); // anio -> list
    perAnio.forEach((v) => {
      if (!grouped.has(v.anio)) grouped.set(v.anio, []);
      grouped.get(v.anio).push(v);
    });

    const rows = [];
    grouped.forEach((list, anio) => {
      // rank by cap
      list.sort((a, b) => b.cap - a.cap);
      list.forEach((item, idx) => (item.rankCap = idx + 1));
      // rank by valor
      list.sort((a, b) => b.valor - a.valor);
      list.forEach((item, idx) => (item.rankValor = idx + 1));
      // output
      list.forEach((item) => {
        rows.push([
          item.anio,
          item.ruc,
          item.razon,
          item.tamano,
          item.esSocio,
          item.cap,
          item.valor,
          item.rankCap,
          item.rankValor,
        ]);
      });
    });

    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]); // por ano desc
      return a[8] - b[8]; // luego por rank valor asc
    });
    const header = ["ANIO", "RUC", "RAZON_SOCIAL", "TAMANO", "ES_SOCIO", "CAPACITACIONES", "VALOR_TOTAL", "RANK_CAP", "RANK_VALOR"];
    const sheet = getSheetOrCreate(SHEETS.OUT_TOP);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 7, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_TOP}: ${rows.length} filas`);
  }

  function generateSocios(registros) {
    const map = new Map(); // anio|socio
    registros.forEach((r) => {
      const k = `${r.anio}|${r.esSocio ? "1" : "0"}`;
      const cur = map.get(k) || { empresas: new Set(), cap: 0, valor: 0 };
      cur.empresas.add(r.key);
      cur.cap += r.capCount;
      cur.valor += r.valor;
      map.set(k, cur);
    });
    const rows = [];
    map.forEach((v, k) => {
      const [anio, socioFlag] = k.split("|");
      rows.push([anio, socioFlag === "1", v.empresas.size, v.cap, v.valor]);
    });
    rows.sort((a, b) => b[0].localeCompare(a[0]));
    const header = ["ANIO", "ES_SOCIO", "EMPRESAS", "CAPACITACIONES", "VALOR_TOTAL"];
    const sheet = getSheetOrCreate(SHEETS.OUT_SOCIOS);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_SOCIOS}: ${rows.length} filas`);
  }

  function generateSociosTamano(registros) {
    const map = new Map(); // anio|socio|tam
    registros.forEach((r) => {
      const k = `${r.anio}|${r.esSocio ? "1" : "0"}|${r.tamano}`;
      const cur = map.get(k) || { empresas: new Set(), cap: 0, valor: 0 };
      cur.empresas.add(r.key);
      cur.cap += r.capCount;
      cur.valor += r.valor;
      map.set(k, cur);
    });
    const rows = [];
    map.forEach((v, k) => {
      const [anio, socioFlag, tam] = k.split("|");
      rows.push([anio, socioFlag === "1", tam, v.empresas.size, v.cap, v.valor]);
    });
    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
      const oA = TAMANO_ORDER[a[2]] || 99;
      const oB = TAMANO_ORDER[b[2]] || 99;
      return oA - oB;
    });
    const header = ["ANIO", "ES_SOCIO", "TAMANO", "EMPRESAS", "CAPACITACIONES", "VALOR_TOTAL"];
    const sheet = getSheetOrCreate(SHEETS.OUT_SOCIOS_TAM);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_SOCIOS_TAM}: ${rows.length} filas`);
  }

  function generateMaster(registros) {
    const rows = registros.map((r) => [
      r.anio,
      r.key,
      r.ruc,
      r.razon,
      r.tamano,
      r.esSocio,
      r.capCount,
      r.valor,
      r.fuente,
    ]);
    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
      return b[7] - a[7];
    });
    const header = ["ANIO", "KEY", "RUC", "RAZON_SOCIAL", "TAMANO", "ES_SOCIO", "CAPACITACIONES", "VALOR_TOTAL", "FUENTE"];
    const sheet = getSheetOrCreate(SHEETS.OUT_MASTER);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    formatNumberColumn(sheet, 2, 8, rows.length, MONEY_FMT);
    Logger.log(`[DASH5] ${SHEETS.OUT_MASTER}: ${rows.length} filas`);
  }

  function generateDuplicadosSheet(duplicados) {
    const header = ["FUENTE", "ANIO", "KEY", "RUC", "RAZON_SOCIAL", "FECHA", "CAPACITACIONES", "VALOR"];
    const rows = duplicados.map((d) => [
      d.fuente,
      d.anio,
      d.key,
      d.ruc || "",
      d.razon || "",
      d.fecha || "",
      d.capCount || 0,
      d.valor || 0,
    ]);
    rows.sort((a, b) => b[1].localeCompare(a[1]));
    const sheet = getSheetOrCreate("PIVOT_CAPACITACIONES_DUPLICADOS");
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    Logger.log(`[DASH5] PIVOT_CAPACITACIONES_DUPLICADOS: ${rows.length} filas`);
  }

  function generateSlicerMasterSheet() {
    const slicerSources = [
      { name: SHEETS.OUT_RESUMEN, cols: 6 },
      { name: SHEETS.OUT_RESUMEN_SOCIOS, cols: 5 },
      { name: SHEETS.OUT_TOP, cols: 9 },
      { name: SHEETS.OUT_SOCIOS, cols: 5 },
      { name: SHEETS.OUT_SOCIOS_TAM, cols: 6 },
      { name: SHEETS.OUT_MASTER, cols: 9 },
      { name: "PIVOT_CAPACITACIONES_DUPLICADOS", cols: 8 },
    ];

    const ss = SpreadsheetApp.getActive();
    const headerSet = new Set(["SOURCE_SHEET"]);
    const sheetsData = [];

    slicerSources.forEach(({ name, cols }) => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const numRows = sh.getLastRow();
      if (numRows === 0) return;

      const values = sh.getRange(1, 1, numRows, cols).getValues(); // valores crudos para preservar numeros
      if (!values || values.length === 0) return;

      const header = values[0]
        .slice(0, cols)
        .map((h, idx) => {
          const clean = (h || "").toString().trim();
          return clean || `COL_${idx + 1}`;
        });
      header.forEach((h) => headerSet.add(h));

      const rows = values
        .slice(1)
        .map((row) => row.slice(0, cols))
        .filter((row) => row.some((cell) => (cell || "").toString().trim() !== ""));

      if (rows.length) sheetsData.push({ name, header, rows });
    });

    const masterHeaders = ["SOURCE_SHEET", ...Array.from(headerSet).filter((h) => h !== "SOURCE_SHEET")];
    const mergedRows = [];

    sheetsData.forEach(({ name, header, rows }) => {
      const indexMap = {};
      header.forEach((h, idx) => {
        if (indexMap[h] === undefined) indexMap[h] = idx;
      });

      rows.forEach((row) => {
        const mergedRow = masterHeaders.map((col) => {
          if (col === "SOURCE_SHEET") return name;
          const idx = indexMap[col];
          return idx !== undefined && idx < row.length ? row[idx] : "";
        });
        mergedRows.push(mergedRow);
      });
    });

    const sheet = getSheetOrCreate(SHEETS.OUT_MASTER_SLICER);
    sheet.clear();
    const output = mergedRows.length ? [masterHeaders, ...mergedRows] : [masterHeaders];
    sheet.getRange(1, 1, output.length, masterHeaders.length).setValues(output);
    sheet.autoResizeColumns(1, masterHeaders.length);

    if (mergedRows.length) {
      const headersWritten = sheet
        .getRange(1, 1, 1, masterHeaders.length)
        .getValues()[0]
        .map((h) => (h || "").toString().trim().toUpperCase());
      const moneyCols = [];
      const moneyMCols = [];
      headersWritten.forEach((h, idx) => {
        if (h.indexOf("VENTAS_") === 0 || h.indexOf("VALOR") === 0) moneyCols.push(idx + 1);
        if (h.indexOf("VENTAS_") === 0 && h.endsWith("_M")) moneyMCols.push(idx + 1);
      });
      moneyCols.forEach((col) => formatNumberColumn(sheet, 2, col, mergedRows.length, MONEY_FMT));
      moneyMCols.forEach((col) => formatNumberColumn(sheet, 2, col, mergedRows.length, MONEY_FMT));
    }

    Logger.log(`[DASH5] ${SHEETS.OUT_MASTER_SLICER}: ${mergedRows.length} filas (merge para slicers)`);
  }

  function refreshDashboardCapacitaciones() {
    buildCapacitacionesDashboard();
    Logger.log("Tablas de capacitaciones regeneradas.");
  }

  function onOpenDash5() {
    SpreadsheetApp.getUi().createMenu("DASH5").addItem("Generar tablas capacitaciones", "refreshDashboardCapacitaciones").addToUi();
  }

  // Trigger manual via checkbox (hoja REPORTE_1, celda Q209)
  const DASH5_TRIGGER_SHEET = "REPORTE_1";
  const DASH5_TRIGGER_CELL = "Q209";

  function onEditDash5(e) {
    const range = e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== DASH5_TRIGGER_SHEET) return;
    if (range.getA1Notation() !== DASH5_TRIGGER_CELL) return;

    const val = range.getValue();
    if (val === true) {
      const ss = SpreadsheetApp.getActive();
      ss.toast("Actualizando dashboard…");
      try {
        refreshDashboardCapacitaciones();
        ss.toast("Dashboard listo.");
        sheet.getRange("Q210").setValue("Ultima actualizacion: " + new Date());
      } finally {
        range.setValue(false); // deja la casilla lista para el siguiente clic
      }
    }
  }

  global.refreshDashboardCapacitaciones = refreshDashboardCapacitaciones;
  global.onOpenDash5 = onOpenDash5;
  global.onEditDash5 = onEditDash5;
})(this);

function refreshDashboardCapacitacionesWrapper() {
  return refreshDashboardCapacitaciones();
}

function onOpenDash5Wrapper() {
  return onOpenDash5();
}

// Wrapper para que el trigger instalable detecte la función en el selector
function onEditDash5Wrapper(e) {
  return onEditDash5(e);
}

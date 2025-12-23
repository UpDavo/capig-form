/**
 * Genera las tablas para el dashboard (ventas por año, sectores,
 * tamaños y afiliaciones). Ejecutar refreshDashboardTables() para
 * regenerar todo.
 */

// ------------------ Config ------------------ //
const DASH_SHEETS = {
  BASE: "SOCIOS",
  BASE_FALLBACK: "BASE DE DATOS",
  REGISTRO: "SOCIOS",
  REGISTRO_FALLBACK: "REGISTRO_AFILIADO",
  VENTAS: "VENTAS_SOCIO",
  VENTAS_FALLBACK: "VENTAS_AFILIADOS",
  ESTADO: "ESTADO_SOCIO",
  ESTADO_FALLBACK: "ESTADO_AFILIADOS",
  TAMANO_GLOBAL: "TAMANO_EMPRESA_GLOBAL",
  OUT_VENTAS: "DASH_VENTAS_ANIO",
  OUT_AFILIACIONES: "DASH_AFILIACIONES_ANIO",
  OUT_EMPRESAS: "DASH_EMPRESAS_TAMANO",
  OUT_PIVOT_VENTAS_TAMANO: "PIVOT_VENTAS_ANIO_TAMANO",
  OUT_PIVOT_VENTAS_SECTOR: "PIVOT_VENTAS_ANIO_SECTOR",
  OUT_PIVOT_AFILIACIONES: "PIVOT_AFILIACIONES_ANIO",
  // Hojas filtradas para los gráficos
  FILTRO_VENTAS: "FILTRO_VENTAS",
  FILTRO_SECTORES: "FILTRO_SECTORES",
  FILTRO_AFILIACIONES: "FILTRO_AFILIACIONES",
};



// Umbrales para inferir tamaño por ventas anuales
const TAMANO_BY_MONTO = [
  { max: 100000, label: "MICRO" },
  { max: 1000000, label: "PEQUENA" },
  { max: 5000000, label: "MEDIANA" },
  { max: Number.POSITIVE_INFINITY, label: "GRANDE" },
];
const TAMANO_ORDER = { GRANDE: 1, MEDIANA: 2, PEQUENA: 3, MICRO: 4 };
// Si el tamaño viene vacío en la base, asignar este valor para efectos de conteo ("" = no asignar).
const DEFAULT_TAMANO_IF_EMPTY = "";
// Si no existe hoja de registro, reutilizar la base como registro para que los conteos de afiliaciones coincidan con los pivots de usuario (duplica afiliaciones).
const INCLUDE_BASE_AS_REGISTRO_WHEN_MISSING = true;
// Contar empresas por fila en la base (en lugar de claves únicas) para cuadrar con los conteos del Excel.
const COUNT_EMPRESAS_POR_FILA = true;

// Multiplicador para cuando los montos anuales vienen en miles (p.ej. 41.702 -> 41,702,000).
// Ajusta a 1 si los valores ya vienen en monto real.
const YEAR_COL_THOUSANDS_MULTIPLIER = 1;  // Valores ya están en números reales (no en miles)

// Alias de columnas
const FIELD_ALIASES = {
  RUC: ["RUC", "NUMERO_RUC", "NUM_RUC"],
  RAZON: ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "NOMBRE", "RAZON_SOCIA"],
  ALT_ID: [
    "ID_UNICO",
    "ID ÚNICO",
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
    "N°",
    "NUM",
    "NUMERO",
  ],
  TAMANO: [
    "TAMANO",
    "TAMANIO",
    "TAMAÑO",
    "TAMANO_EMPRESA",
    "TAMANO_EMP",
    "TAMANO EMPRESA",
    "TAMANO_ACTUAL",
    "TAMANO_2023",
    "TAMANO_2022",
  ],
  SECTOR: ["SECTOR", "SECTOR "],
  FECHA_AF: [
    "FECHA_AFILIACION",
    "FECHA AFILIACION",
    "FECHA_INGRESO",
    "FECHA DE INGRESO",
    "FECHA_REGISTRO",
    "FECHA",
  ],
  VENTAS: [
    "VENTAS",
    "VENTAS_TOTAL",
    "VENTAS_ANUAL",
    "VENTAS_ANUALES",
    "VENTAS_MONT_EST",
    "MONTO_ESTIMADO",
    "MONTO_VENTAS",
    "VALOR TOTAL",
    "MONTO",
    "VALOR_APORTE",
    "MONTO_TOTAL",
  ],
  ANIO: ["ANIO", "ANO", "AÑO", "ANIO_VENTA", "ANO_VENTA", "AÑO_VENTA", "AÑO VENTA"],
};

// Mapeo de códigos de tamaño a texto (cuando solo existen columnas de códigos)
const TAMANO_CODE_TO_TEXT = { 1: "MICRO", 2: "PEQUENA", 3: "MEDIANA", 4: "GRANDE" };
function getTamanoFromCodes(headerIndex, row) {
  const candidates = ["TAMANO_COD_2023", "TAMANO_COD_2022", "T2023", "T_2023", "T2022", "T_2022"];
  for (const alias of candidates) {
    const idx = headerIndex[normalizeLabel(alias)];
    if (idx !== undefined && idx < row.length) {
      const code = Number(row[idx]);
      if (code && TAMANO_CODE_TO_TEXT[code]) return TAMANO_CODE_TO_TEXT[code];
    }
  }
  // Escaneo genérico para variantes tipo T2021
  for (const [key, idx] of Object.entries(headerIndex)) {
    if (!/^T_?20\d{2}$/.test(key)) continue;
    const code = Number(row[idx]);
    if (code && TAMANO_CODE_TO_TEXT[code]) return TAMANO_CODE_TO_TEXT[code];
  }
  return "";
}

// ------------------ Utils ------------------ //
function normalizeLabel(label) {
  let txt = (label || "").toString().trim().toUpperCase();
  txt = txt.replace(/�/g, "N");
  txt = txt.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  txt = txt.replace(/[^A-Z0-9]+/g, "_");
  txt = txt.replace(/^_+|_+$/g, "");
  return txt;
}
function normalizeKey(val) {
  return (val || "").toString().trim().toUpperCase().replace(/\s+/g, "_");
}
function normalizeName(val) {
  return (val || "").toString().trim().toUpperCase().replace(/\s+/g, " ");
}
function normalizeSector(val) {
  let s = (val || "").toString().trim().toUpperCase();
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");  // Elimina acentos
  if (!s) return "SIN CLASIFICAR";
  // Normalizar sectores comunes
  if (s.indexOf("QUIM") !== -1) return "QUIMICO";
  if (s.indexOf("METAL") !== -1) return "METALMECANICO";
  if (s.indexOf("ALIMENT") !== -1) return "ALIMENTOS";
  if (s.indexOf("AGRIC") !== -1 || s.indexOf("AGROP") !== -1) return "AGRICOLA";
  if (s.indexOf("MAQUIN") !== -1) return "MAQUINARIAS";
  if (s.indexOf("CONST") !== -1) return "CONSTRUCCION";
  if (s.indexOf("TEXT") !== -1) return "TEXTIL";
  return s;
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
function entityKey(headerIndex, row, rowIdx, allowRowFallback) {
  const ruc = padRuc13(getVal(row, headerIndex, FIELD_ALIASES.RUC));
  const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
  const altId = normalizeKey(getVal(row, headerIndex, FIELD_ALIASES.ALT_ID || []));
  if (altId) {
    // Si hay ID interno, úsalo para diferenciar aunque el RUC se repita.
    if (ruc) return `${ruc}__${altId}`;
    return `ID__${altId}`;
  }
  if (ruc) return ruc;
  if (razon) return razon;
  return allowRowFallback ? `ROW_${rowIdx || 0}` : "";
}
function normalizeTamano(val) {
  const t = (val || "").toString().trim().toUpperCase();
  if (!t) return "";
  const noise = ["TAMANO", "TAMANO_EMPRESA", "TAMANO_EMP", "TAMANO ", "TAMAÑO", "TAMAÑO ", "TAMA�O", "NAN"];
  if (noise.includes(t) || t.startsWith("TAMA")) return "";
  if (t.indexOf("MICRO") !== -1) return "MICRO";
  if (t.indexOf("PEQU") !== -1) return "PEQUENA";
  if (t.indexOf("MEDI") !== -1) return "MEDIANA";
  if (t.indexOf("GRAN") !== -1) return "GRANDE";
  return t;
}
function toNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  let str = val.toString().trim().replace(/\s/g, "");
  if (!str) return 0;

  const hasComma = str.indexOf(",") !== -1;
  const hasDot = str.indexOf(".") !== -1;
  const commaThousandsPattern = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/; // 45,717,089 or 45,717,089.50

  if (hasComma && !hasDot && commaThousandsPattern.test(str)) {
    // Solo comas y con patrón claro de miles: 45,717,089 -> 45717089
    str = str.replace(/,/g, "");
  } else if (hasComma && hasDot) {
    // Mixto coma/punto: decidir según la posición
    if (str.lastIndexOf(".") > str.lastIndexOf(",")) {
      // Formato US: 45,717,089.00
      str = str.replace(/,/g, "");
    } else {
      // Formato EU: 45.717.089,00
      str = str.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (hasComma && !hasDot) {
    // Solo comas, ambiguo. Si parece miles (1,234,567) limpiar; si no, tratar coma como decimal.
    if (commaThousandsPattern.test(str)) {
      str = str.replace(/,/g, "");
    } else {
      const parts = str.split(",");
      if (parts.length === 2 && parts[1].length <= 2) {
        str = `${parts[0]}.${parts[1]}`;
      } else {
        str = str.replace(/,/g, "");
      }
    }
  } else if (!hasComma && hasDot) {
    // Solo puntos. Si hay más de uno, asumir miles y limpiar.
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) {
      str = str.replace(/\./g, "");
    }
  }

  str = str.replace(/[^0-9\.-]/g, "");
  const num = Number(str);
  return Number.isFinite(num) ? num : 0;
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
function toDate(val) {
  const d = parseDateFlexible(val);
  return d ? Utilities.formatDate(d, "GMT", "yyyy-MM-dd") : "";
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
function getTamanoOrder(tam) {
  return TAMANO_ORDER[(tam || "").toString().trim().toUpperCase()] || 99;
}
function getYearColumns(headerIndex) {
  const yearly = {};
  Object.keys(headerIndex).forEach((key) => {
    // Omitir columnas de código de tamaño tipo T2022/T_2023
    if (/^T_?20\d{2}$/.test(key)) return;
    const ventasMatch = key.match(/^VENTAS_?(20\d{2})$/);
    if (ventasMatch) yearly[ventasMatch[1]] = key;
    const yearMatch = key.match(/^(20\d{2})$/); // 2022
    if (yearMatch) yearly[yearMatch[1]] = key;
  });
  const years = Object.keys(yearly).sort();
  return years.map((year) => ({
    year,
    colKey: yearly[year],
    multiplier: YEAR_COL_THOUSANDS_MULTIPLIER,
  }));
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
function rowHasData(row, headerIndex) {
  if (!row) return false;
  // Solo consideramos datos “reales”: RUC, razón, sector, ventas, fechas o columnas de ventas por año.
  const ruc = getVal(row, headerIndex, FIELD_ALIASES.RUC);
  const razon = getVal(row, headerIndex, FIELD_ALIASES.RAZON);
  const sector = getVal(row, headerIndex, FIELD_ALIASES.SECTOR);
  const fecha = getVal(row, headerIndex, FIELD_ALIASES.FECHA_AF);
  const ventas = getVal(row, headerIndex, FIELD_ALIASES.VENTAS);
  const yearCols = getYearColumns(headerIndex);
  const hasYearVal = yearCols.some(({ colKey }) => {
    const idx = headerIndex[colKey];
    return idx !== undefined && row[idx] !== null && row[idx] !== undefined && row[idx].toString().trim() !== "";
  });
  return Boolean(ruc || razon || sector || fecha || ventas || hasYearVal);
}
function hasKeyFallbackData(row, headerIndex, yearCols = []) {
  const tam = normalizeTamano(getVal(row, headerIndex, FIELD_ALIASES.TAMANO));
  const sec = getVal(row, headerIndex, FIELD_ALIASES.SECTOR);
  if (tam || sec) return true;
  for (const { colKey } of yearCols || []) {
    const idx = headerIndex[colKey];
    if (idx !== undefined && row[idx] !== null && row[idx] !== undefined && row[idx].toString().trim() !== "") {
      return true;
    }
  }
  return false;
}

// ------------------ Builders ------------------ //
function buildTamanoEmpresaGlobal() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const ventas = readTableFlexibleByCandidates([DASH_SHEETS.VENTAS, DASH_SHEETS.VENTAS_FALLBACK]);
  const registro = readTableFlexibleByCandidates([DASH_SHEETS.REGISTRO, DASH_SHEETS.REGISTRO_FALLBACK]);

  const yearCols = getYearColumns(base.headerIndex);

  const tamMap = new Map(); // key -> TAMANO
  const sectorMap = new Map(); // key -> SECTOR
  const ventasSum = new Map(); // key -> suma ventas

  function accVenta(key, monto) {
    if (!key) return;
    ventasSum.set(key, (ventasSum.get(key) || 0) + monto);
  }

  // 1) BASE DE DATOS
  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false); // requiere RUC o razón
    if (!key) return;
    const tam =
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(base.headerIndex, row) ||
      DEFAULT_TAMANO_IF_EMPTY;
    const sec = getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR);
    if (tam) tamMap.set(key, tam);
    if (sec) sectorMap.set(key, normalizeSector(sec));
    yearCols.forEach(({ colKey, multiplier }) => {
      const idx = base.headerIndex[colKey];
      const raw = idx !== undefined ? row[idx] : "";
      const monto = toNumber(raw) * multiplier;
      if (monto > 0) accVenta(key, monto);
    });
  });

  // 2) REGISTRO_AFILIADO
  registro.rows.forEach((row, i) => {
    if (!rowHasData(row, registro.headerIndex)) return;
    const key = entityKey(registro.headerIndex, row, i, false); // sin fallback si no hay RUC/razon
    if (!key) return;
    const tam =
      normalizeTamano(getVal(row, registro.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(registro.headerIndex, row);
    const sec = getVal(row, registro.headerIndex, FIELD_ALIASES.SECTOR);
    if (tam && !tamMap.has(key)) tamMap.set(key, tam);
    if (sec && !sectorMap.has(key)) sectorMap.set(key, normalizeSector(sec));
  });

  // 3) VENTAS_AFILIADOS
  ventas.rows.forEach((row, i) => {
    if (!rowHasData(row, ventas.headerIndex)) return;
    const key = entityKey(ventas.headerIndex, row, i, false); // sin fallback si no hay RUC/razon
    if (!key) return;
    const tam =
      normalizeTamano(getVal(row, ventas.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(ventas.headerIndex, row);
    const sec = getVal(row, ventas.headerIndex, FIELD_ALIASES.SECTOR);
    if (tam && !tamMap.has(key)) tamMap.set(key, tam);
    if (sec && !sectorMap.has(key)) sectorMap.set(key, normalizeSector(sec));
    const monto = toNumber(getVal(row, ventas.headerIndex, FIELD_ALIASES.VENTAS));
    if (monto > 0) accVenta(key, monto);
  });

  // 4) Inferir tamaño por monto
  ventasSum.forEach((monto, key) => {
    if (tamMap.has(key)) return;
    for (const rule of TAMANO_BY_MONTO) {
      if (monto <= rule.max) {
        tamMap.set(key, rule.label);
        break;
      }
    }
  });

  // 5) Hoja
  const sh = getSheetOrCreate(DASH_SHEETS.TAMANO_GLOBAL);
  sh.clear();
  const rows = [];
  tamMap.forEach((tam, key) => rows.push([key, tam]));
  rows.sort((a, b) => (a[0] > b[0] ? 1 : -1));
  const headers = ["RUC", "TAMANO"];
  const out = rows.length ? [headers, ...rows] : [headers];
  sh.getRange(1, 1, out.length, headers.length).setValues(out);
  sh.autoResizeColumns(1, headers.length);

  return { tamMap, sectorMap };
}

function buildProfiles() {
  const profiles = new Map();
  const { tamMap, sectorMap } = buildTamanoEmpresaGlobal();
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const registro = readTableFlexibleByCandidates([DASH_SHEETS.REGISTRO, DASH_SHEETS.REGISTRO_FALLBACK]);
  const yearCols = getYearColumns(base.headerIndex);

  // BASE DE DATOS
  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const razon = normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON));
    const tamBase =
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(base.headerIndex, row) ||
      DEFAULT_TAMANO_IF_EMPTY;
    const secBase = normalizeSector(getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR));
    const tam = tamBase || tamMap.get(key) || "";
    const sec = secBase || sectorMap.get(key) || "";
    profiles.set(key, { RAZON_SOCIAL: razon, TAMANO: tam, SECTOR: sec });
  });

  // Refuerzo con TAMANO_GLOBAL
  tamMap.forEach((tam, key) => {
    if (!profiles.has(key)) {
      profiles.set(key, { RAZON_SOCIAL: "", TAMANO: tam, SECTOR: sectorMap.get(key) || "" });
    } else {
      const current = profiles.get(key);
      if (!current.TAMANO && tam) current.TAMANO = normalizeTamano(tam);
      if (!current.SECTOR && sectorMap.get(key)) current.SECTOR = normalizeSector(sectorMap.get(key));
      profiles.set(key, current);
    }
  });

  // RUC nuevos del registro
  registro.rows.forEach((row, i) => {
    if (!rowHasData(row, registro.headerIndex)) return;
    const key = entityKey(registro.headerIndex, row, i, false);
    if (!key || profiles.has(key)) return;
    const razon = normalizeName(getVal(row, registro.headerIndex, FIELD_ALIASES.RAZON));
    profiles.set(key, {
      RAZON_SOCIAL: razon,
      TAMANO: normalizeTamano(tamMap.get(key) || ""),
      SECTOR: sectorMap.get(key) || "",
    });
  });

  return profiles;
}

function buildVentasAnio(profile) {
  const ventasRows = [];
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const yearCols = getYearColumns(base.headerIndex);

  // Histórico
  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const razon = normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON));
    const tam =
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(base.headerIndex, row);
    const sec = normalizeSector(getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR));
    const prof = profile.get(key) || {};
    yearCols.forEach(({ year, colKey, multiplier }) => {
      const idx = base.headerIndex[colKey];
      const raw = idx !== undefined ? row[idx] : "";
      const monto = toNumber(raw) * multiplier;
      if (monto > 0) {
        ventasRows.push([
          year,
          tam || prof.TAMANO || "",
          sec || prof.SECTOR || "",
          monto,
          padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC)),
          razon || prof.RAZON_SOCIAL || "",
          DASH_SHEETS.BASE,
        ]);
      }
    });
  });

  // Formulario VENTAS_AFILIADOS
  const ventas = readTableFlexibleByCandidates([DASH_SHEETS.VENTAS, DASH_SHEETS.VENTAS_FALLBACK]);
  ventas.rows.forEach((row, i) => {
    if (!rowHasData(row, ventas.headerIndex)) return;
    const key = entityKey(ventas.headerIndex, row, i);
    if (!key) return;
    const razon = normalizeName(getVal(row, ventas.headerIndex, FIELD_ALIASES.RAZON));
    const prof = profile.get(key) || {};
    const tam =
      normalizeTamano(getVal(row, ventas.headerIndex, FIELD_ALIASES.TAMANO) || prof.TAMANO || "") ||
      getTamanoFromCodes(ventas.headerIndex, row);
    const sec = normalizeSector(getVal(row, ventas.headerIndex, FIELD_ALIASES.SECTOR)) || prof.SECTOR || "";
    const monto = toNumber(getVal(row, ventas.headerIndex, FIELD_ALIASES.VENTAS));
    const fechaRaw = getVal(row, ventas.headerIndex, FIELD_ALIASES.FECHA_AF);
    // Priorizar el campo ANIO explícito; si no existe, derivar de la fecha.
    const anio = getVal(row, ventas.headerIndex, FIELD_ALIASES.ANIO) || getYear(fechaRaw) || "";
    if (!anio || monto === 0) return;
    ventasRows.push([anio, tam, sec, monto, padRuc13(getVal(row, ventas.headerIndex, FIELD_ALIASES.RUC)), razon || prof.RAZON_SOCIAL || "", DASH_SHEETS.VENTAS]);
  });

  const sh = getSheetOrCreate(DASH_SHEETS.OUT_VENTAS);
  sh.clear();
  const headers = ["ANIO", "TAMANO", "SECTOR", "VENTAS_MONTO", "RUC", "RAZON_SOCIAL", "FUENTE"];
  const out = ventasRows.length ? [headers, ...ventasRows] : [headers];
  sh.getRange(1, 1, out.length, headers.length).setValues(out);
  sh.autoResizeColumns(1, headers.length);
}

function buildAfiliaciones(profile) {
  const afiliaciones = [];

  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const registro = readTableFlexibleByCandidates([DASH_SHEETS.REGISTRO, DASH_SHEETS.REGISTRO_FALLBACK]);
  const registroRows = registro.rows.length === 0 && INCLUDE_BASE_AS_REGISTRO_WHEN_MISSING ? base.rows : registro.rows;

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const razon = normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON));
    const fa = toDate(getVal(row, base.headerIndex, FIELD_ALIASES.FECHA_AF));
    const anio = getYear(fa);
    if (!anio) return;
    const prof = profile.get(key) || {};
    afiliaciones.push([anio, padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC)), razon || prof.RAZON_SOCIAL || "", DASH_SHEETS.BASE]);
  });

  registroRows.forEach((row, i) => {
    if (!rowHasData(row, registro.headerIndex)) return;
    const key = entityKey(registro.headerIndex, row, i);
    if (!key) return;
    const razon = normalizeName(getVal(row, registro.headerIndex, FIELD_ALIASES.RAZON));
    const fa = toDate(getVal(row, registro.headerIndex, FIELD_ALIASES.FECHA_AF));
    const anio = getYear(fa);
    if (!anio) return;
    const prof = profile.get(key) || {};
    afiliaciones.push([anio, padRuc13(getVal(row, registro.headerIndex, FIELD_ALIASES.RUC)), razon || prof.RAZON_SOCIAL || "", DASH_SHEETS.REGISTRO]);
  });

  const sh = getSheetOrCreate(DASH_SHEETS.OUT_AFILIACIONES);
  sh.clear();
  const headers = ["ANIO_AFILIACION", "RUC", "RAZON_SOCIAL", "FUENTE"];
  const out = afiliaciones.length ? [headers, ...afiliaciones] : [headers];
  sh.getRange(1, 1, out.length, headers.length).setValues(out);
  sh.autoResizeColumns(1, headers.length);
}

function buildEmpresasTamano(profile) {
  const byTam = {};

  if (COUNT_EMPRESAS_POR_FILA) {
    // Contar cada fila de la base (sin deduplicar por RUC) para cuadrar con las tablas dinámicas originales.
    const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
    base.rows.forEach((row) => {
      if (!rowHasData(row, base.headerIndex)) return;
      const t =
        normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
        getTamanoFromCodes(base.headerIndex, row) ||
        DEFAULT_TAMANO_IF_EMPTY;
      if (!t) return; // omitir filas sin tamaño declarado
      byTam[t] = (byTam[t] || 0) + 1;
    });
  } else {
    profile.forEach((data) => {
      const t = (data.TAMANO || "").toString().trim();
      if (!t) return;
      byTam[t] = (byTam[t] || 0) + 1;
    });
  }

  const rows = Object.entries(byTam).map(([tam, count]) => [tam, count]);
  rows.sort((a, b) => a[0].localeCompare(b[0]));
  const sh = getSheetOrCreate(DASH_SHEETS.OUT_EMPRESAS);
  sh.clear();
  const headers = ["TAMANO", "EMPRESAS"];
  const out = rows.length ? [headers, ...rows] : [headers];
  sh.getRange(1, 1, out.length, headers.length).setValues(out);
  sh.autoResizeColumns(1, headers.length);

  // Limpiar cualquier formato previo en la columna EMPRESAS (debe ser número simple)
  if (rows.length > 0) {
    sh.getRange(2, 2, rows.length).setNumberFormat("0");  // Formato de número entero simple
  }
}

function buildAggregates() {
  const ventas = readTableFlexible(DASH_SHEETS.OUT_VENTAS);
  const aggTam = new Map();
  const aggSec = new Map();

  ventas.rows.forEach((row) => {
    const anio = getVal(row, ventas.headerIndex, ["ANIO"]);
    const tam = getVal(row, ventas.headerIndex, ["TAMANO"]) || "SIN_TAMANO";
    const sec = getVal(row, ventas.headerIndex, ["SECTOR"]) || "SIN_SECTOR";
    const monto = toNumber(getVal(row, ventas.headerIndex, ["VENTAS_MONTO", "VENTAS"]));
    if (!anio || monto === 0) return;
    const keyTam = `${anio}||${tam}`;
    const keySec = `${anio}||${sec}`;
    aggTam.set(keyTam, (aggTam.get(keyTam) || 0) + monto);
    aggSec.set(keySec, (aggSec.get(keySec) || 0) + monto);
  });

  const rowsTam = Array.from(aggTam.entries()).map(([k, v]) => {
    const [anio, tam] = k.split("||");
    return [anio, tam, v, getTamanoOrder(tam)];
  });
  rowsTam.sort((a, b) => {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[3] !== b[3]) return a[3] - b[3];
    return a[1].localeCompare(b[1]);
  });

  // Calcular totales por año para agregar filas de TOTAL
  const totalesPorAnio = new Map();
  rowsTam.forEach(([anio, tam, ventas, orden]) => {
    const currentTotal = totalesPorAnio.get(anio) || 0;
    totalesPorAnio.set(anio, currentTotal + ventas);
  });

  // Agregar filas de TOTAL al final de cada año
  const rowsConTotales = [];
  let lastAnio = null;
  rowsTam.forEach((row, idx) => {
    rowsConTotales.push(row);
    const currentAnio = row[0];
    const nextAnio = idx < rowsTam.length - 1 ? rowsTam[idx + 1][0] : null;
    // Si cambia el año o es la última fila, agregar TOTAL
    if (currentAnio !== nextAnio) {
      rowsConTotales.push([currentAnio, "TOTAL", totalesPorAnio.get(currentAnio), 999]);
    }
  });

  const shTam = getSheetOrCreate(DASH_SHEETS.OUT_PIVOT_VENTAS_TAMANO);
  shTam.clear();
  shTam.getRange(1, 1, rowsConTotales.length + 1, 4).setValues([["ANIO", "TAMANO", "VENTAS_TOTAL", "ORDEN"], ...rowsConTotales]);
  shTam.autoResizeColumns(1, 4);


  // Aplicar formato personalizado a las VENTAS_TOTAL (columna C) - LATAM: punto=miles, coma=decimal
  const formatoMil = '"$"#.##0,00\\ "mil"';  // $138.368,79 mil (con "mil")
  const formatoSinMil = '"$"#.##0,00';  // $1.316.520.115,49 (sin "mil" para totales)

  if (rowsConTotales.length > 0) {
    // Aplicar formato con "mil" a todas las filas primero
    const rangoVentas = shTam.getRange(2, 3, rowsConTotales.length);
    rangoVentas.setNumberFormat(formatoMil);

    // Luego aplicar formato SIN "mil" solo a las filas de TOTAL
    rowsConTotales.forEach((row, idx) => {
      if (row[1] === "TOTAL") {
        shTam.getRange(idx + 2, 3).setNumberFormat(formatoSinMil);  // +2 porque idx empieza en 0 y hay header
      }
    });
  }



  const rowsSec = Array.from(aggSec.entries()).map(([k, v]) => {
    const [anio, sec] = k.split("||");
    return [anio, sec, v];
  });
  rowsSec.sort((a, b) => (a[0] === b[0] ? a[1].localeCompare(b[1]) : a[0].localeCompare(b[0])));
  const shSec = getSheetOrCreate(DASH_SHEETS.OUT_PIVOT_VENTAS_SECTOR);
  shSec.clear();
  shSec.getRange(1, 1, rowsSec.length + 1, 3).setValues([["ANIO", "SECTOR", "VENTAS_TOTAL"], ...rowsSec]);
  shSec.autoResizeColumns(1, 3);

  const af = readTableFlexible(DASH_SHEETS.OUT_AFILIACIONES);
  const aggAf = new Map();
  af.rows.forEach((row) => {
    const anio = getVal(row, af.headerIndex, ["ANIO_AFILIACION", "ANIO"]);
    if (!anio) return;
    aggAf.set(anio, (aggAf.get(anio) || 0) + 1);
  });
  const rowsAf = Array.from(aggAf.entries())
    .map(([anio, count]) => {
      const anioNum = parseInt(anio, 10);
      return [isNaN(anioNum) ? anio : anioNum, count];
    })
    .sort((a, b) => (a[0] > b[0] ? 1 : -1));
  const shAf = getSheetOrCreate(DASH_SHEETS.OUT_PIVOT_AFILIACIONES);
  shAf.clear();
  shAf.getRange(1, 1, rowsAf.length + 1, 2).setValues([["ANIO_AFILIACION", "AFILIACIONES"], ...rowsAf]);
  shAf.getRange(2, 1, rowsAf.length, 1).setNumberFormat("0"); // Numérico para permitir filtro "Entre" (Range)
  shAf.autoResizeColumns(1, 2);
}

// ------------------ Debug helpers ------------------ //
function debugCounts() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const yearCols = getYearColumns(base.headerIndex);

  let nonEmpty = 0;
  let withKey = 0;
  const noKeyRows = [];

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    nonEmpty += 1;
    const key = entityKey(base.headerIndex, row, i, false);
    if (key) {
      withKey += 1;
      return;
    }
    // Sin RUC/razón: guardar detalle para inspección
    const tam =
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(base.headerIndex, row);
    const sec = normalizeSector(getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR));
    const ventas = yearCols
      .map(({ colKey }) => {
        const idx = base.headerIndex[colKey];
        return idx !== undefined ? row[idx] : "";
      })
      .filter((v) => (v || "").toString().trim() !== "");
    noKeyRows.push([
      i + 2, // número de fila real en la hoja (headers en 1)
      padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC)),
      normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON)),
      tam,
      sec,
      ventas.length ? ventas.join(" | ") : "",
    ]);
  });

  const sh = getSheetOrCreate("DEBUG_COUNTS");
  sh.clear();
  const rowsOut = [
    ["Metrica", "Valor"],
    ["BASE filas no vacías", nonEmpty],
    ["BASE filas con RUC/RAZON", withKey],
    ["BASE filas sin RUC/RAZON (listadas abajo)", noKeyRows.length],
  ];
  const detailHeader = ["Fila hoja", "RUC", "RAZON_SOCIAL", "TAMANO/T_CODE", "SECTOR", "VENTAS_ANYO"];
  const detail = noKeyRows.length ? [detailHeader, ...noKeyRows] : [detailHeader];
  sh.getRange(1, 1, rowsOut.length, rowsOut[0].length).setValues(rowsOut);
  sh.getRange(rowsOut.length + 2, 1, detail.length, detail[0].length).setValues(detail);
  sh.autoResizeColumns(1, detailHeader.length);
  SpreadsheetApp.flush();
  Logger.log("Debug listo en hoja DEBUG_COUNTS");
}

// Lista duplicados por clave (RUC o RAZON) en la hoja BASE
function debugDuplicates() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const duplicates = [];
  const seen = new Map();

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const idxList = seen.get(key) || [];
    idxList.push(i + 2); // fila real
    seen.set(key, idxList);
  });

  seen.forEach((rowsIdx, key) => {
    if (rowsIdx.length <= 1) return;
    const firstIdx = rowsIdx[0] - 2; // índice en base.rows
    const row = base.rows[firstIdx] || [];
    duplicates.push([
      key,
      rowsIdx.join(","),
      normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON)),
      padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC)),
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) || getTamanoFromCodes(base.headerIndex, row),
      getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR),
    ]);
  });

  const sh = getSheetOrCreate("DEBUG_DUPLICATES");
  sh.clear();
  const header = [["KEY", "FILAS", "RAZON_SOCIAL", "RUC", "TAMANO", "SECTOR"]];
  const out = duplicates.length ? header.concat(duplicates) : header;
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.autoResizeColumns(1, out[0].length);
  SpreadsheetApp.flush();
  Logger.log("Debug listo en hoja DEBUG_DUPLICATES");
}

function debugDuplicates() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const yearCols = getYearColumns(base.headerIndex);
  const keyMap = new Map();

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const tam =
      normalizeTamano(getVal(row, base.headerIndex, FIELD_ALIASES.TAMANO)) ||
      getTamanoFromCodes(base.headerIndex, row);
    const sec = normalizeSector(getVal(row, base.headerIndex, FIELD_ALIASES.SECTOR));
    const ventas = yearCols
      .map(({ colKey }) => {
        const idx = base.headerIndex[colKey];
        return idx !== undefined ? row[idx] : "";
      })
      .filter((v) => (v || "").toString().trim() !== "")
      .join(" | ");

    const info = [
      i + 2,
      padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC)),
      normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON)),
      tam,
      sec,
      ventas,
    ];
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push(info);
  });

  const dupRows = [];
  keyMap.forEach((list, key) => {
    if (list.length > 1) list.forEach((info) => dupRows.push([key, ...info]));
  });

  const sh = getSheetOrCreate("DEBUG_DUP_KEYS");
  sh.clear();
  const summary = [
    ["Metrica", "Valor"],
    ["Claves duplicadas", dupRows.length],
  ];
  const detailHeader = ["KEY", "Fila hoja", "RUC", "RAZON_SOCIAL", "TAMANO", "SECTOR", "VENTAS_ANYO"];
  const detail = dupRows.length > 0 ? [detailHeader, ...dupRows] : [detailHeader, ["SIN DUPLICADOS", "", "", "", "", "", ""]];
  sh.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  sh.getRange(summary.length + 2, 1, detail.length, detail[0].length).setValues(detail);
  sh.autoResizeColumns(1, detailHeader.length);
  SpreadsheetApp.flush();
  Logger.log("Debug de duplicados listo en hoja DEBUG_DUP_KEYS");
}

function debugDuplicates() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const seen = new Map();
  const duplicates = [];

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const rowNum = i + 2; // encabezados en fila 1
    if (seen.has(key)) {
      duplicates.push([rowNum, key, normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON))]);
    } else {
      seen.set(key, rowNum);
    }
  });

  const sh = getSheetOrCreate("DEBUG_DUP_KEYS");
  sh.clear();
  const header = [["Fila hoja", "KEY (RUC/RAZON)", "RAZON_SOCIAL"]];
  sh.getRange(1, 1, header.length, header[0].length).setValues(header);
  if (duplicates.length) {
    sh.getRange(2, 1, duplicates.length, duplicates[0].length).setValues(duplicates);
  }
  sh.autoResizeColumns(1, 3);
  SpreadsheetApp.flush();
  Logger.log("Debug duplicados listo en hoja DEBUG_DUP_KEYS");
}

// Listar claves duplicadas (mismo RUC/Razon) en la base
function debugDuplicates() {
  const base = readTableFlexibleByCandidates([DASH_SHEETS.BASE, DASH_SHEETS.BASE_FALLBACK]);
  const dupMap = new Map();

  base.rows.forEach((row, i) => {
    if (!rowHasData(row, base.headerIndex)) return;
    const key = entityKey(base.headerIndex, row, i, false);
    if (!key) return;
    const razon = normalizeName(getVal(row, base.headerIndex, FIELD_ALIASES.RAZON));
    const ruc = padRuc13(getVal(row, base.headerIndex, FIELD_ALIASES.RUC));
    const entry = dupMap.get(key) || { count: 0, rows: [], razon, ruc };
    entry.count += 1;
    entry.rows.push(i + 2); // fila real en la hoja
    // refrescar razon/ruc si estaban vacíos
    if (!entry.razon && razon) entry.razon = razon;
    if (!entry.ruc && ruc) entry.ruc = ruc;
    dupMap.set(key, entry);
  });

  const duplicates = Array.from(dupMap.entries())
    .filter(([, v]) => v.count > 1)
    .map(([key, v]) => [key, v.count, v.ruc || "", v.razon || "", v.rows.join(", ")])
    .sort((a, b) => b[1] - a[1]);

  const sh = getSheetOrCreate("DEBUG_DUP_KEYS");
  sh.clear();
  const header = [["KEY(RUC/RAZON)", "COUNT", "RUC", "RAZON_SOCIAL", "FILAS_BASE"]];
  const summary = [["Total claves", dupMap.size], ["Duplicadas", duplicates.length]];
  sh.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  sh.getRange(summary.length + 2, 1, header.length, header[0].length).setValues(header);
  if (duplicates.length) {
    sh.getRange(summary.length + 3, 1, duplicates.length, header[0].length).setValues(duplicates);
  }
  sh.autoResizeColumns(1, header[0].length);
  SpreadsheetApp.flush();
  Logger.log("Debug de duplicados listo en hoja DEBUG_DUP_KEYS");
}



// ------------------ Entrypoint ------------------ //
function refreshDashboardTables() {
  const profile = buildProfiles();
  buildVentasAnio(profile);
  buildAfiliaciones(profile);
  buildEmpresasTamano(profile);
  buildAggregates();

  Logger.log("Tablas de dashboard regeneradas (listas para Slicers).");
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("DASHBOARD")
    .addItem("Generar tablas dashboard", "refreshDashboardTables")
    .addToUi();
}

// Trigger manual via checkbox (por ejemplo celda Q1 en hoja REPORTE_1)
const DASH1_TRIGGER_SHEET = "REPORTE_1";
const DASH1_TRIGGER_CELL = "Q1";

function onEdit(e) {
  const range = e.range;
  if (!range) return;
  const sheet = range.getSheet();
  if (!sheet || sheet.getName() !== DASH1_TRIGGER_SHEET) return;
  if (range.getA1Notation() !== DASH1_TRIGGER_CELL) return;

  const val = range.getValue();
  if (val === true) {
    const ss = SpreadsheetApp.getActive();
    ss.toast("Actualizando dashboard…");
    try {
      refreshDashboardTables();
      ss.toast("Dashboard listo.");
      sheet.getRange("Q2").setValue("Ultima actualizacion: " + new Date());
    } finally {
      // Opcional: desmarcar para dejar listo el checkbox
      range.setValue(false);
    }
  }
}

// Wrapper opcional para asignar a un botón/imagen
function refreshDashboardTablesWrapper() {
  return refreshDashboardTables();
}

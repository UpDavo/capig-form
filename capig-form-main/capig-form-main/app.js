/**
 * Consolida TODAS las hojas en una sola tabla FACT_DATOS_COMPLETA.
 * Sin pivots ni gráficos. Pensado para BI externo (Data Studio, Power BI, etc.).
 */

const SHEET_NAMES = {
  FACT: "FACT_DATOS_COMPLETA",
  BASE: "SOCIOS",
  REGISTRO: "SOCIOS",
  VENTAS: "VENTAS_SOCIO",
  DIAG: "ASESORIAS",
  DIAG_FINAL: "ASESORIAS",
  CAP: "CAPACITACIONES",
  CAP_FINAL: "CAPACITACIONES",
  ESTADO: "ESTADO_SOCIO",
  SECTOR: "SECTOR",
};

const SHEET_FALLBACKS = {
  BASE: "BASE DE DATOS",
  REGISTRO: "REGISTRO_AFILIADO",
  VENTAS: "VENTAS_AFILIADOS",
  DIAG: "DIAGNOSTICOS",
  DIAG_FINAL: "DIAGNOSTICO_FINAL",
  CAP: "CAPACITACIONES",
  CAP_FINAL: "CAPACITACIONES_FINAL",
  ESTADO: "ESTADO_AFILIADOS",
  SECTOR: null,
};

const SOURCES = [
  { name: SHEET_NAMES.BASE, fallback: SHEET_FALLBACKS.BASE, tipo: "historico" },
  { name: SHEET_NAMES.REGISTRO, fallback: SHEET_FALLBACKS.REGISTRO, tipo: "registro" },
  { name: SHEET_NAMES.VENTAS, fallback: SHEET_FALLBACKS.VENTAS, tipo: "ventas" },
  { name: SHEET_NAMES.DIAG, fallback: SHEET_FALLBACKS.DIAG, tipo: "diagnostico" },
  { name: SHEET_NAMES.DIAG_FINAL, fallback: SHEET_FALLBACKS.DIAG_FINAL, tipo: "diagnostico" },
  { name: SHEET_NAMES.CAP, fallback: SHEET_FALLBACKS.CAP, tipo: "capacitacion" },
  { name: SHEET_NAMES.CAP_FINAL, fallback: SHEET_FALLBACKS.CAP_FINAL, tipo: "capacitacion" },
  { name: SHEET_NAMES.ESTADO, fallback: SHEET_FALLBACKS.ESTADO, tipo: "estado" },
];

const FACT_HEADERS = [
  "RUC",
  "RAZON_SOCIAL",
  "NOMBRE_COMERCIAL",
  "CIUDAD",
  "SECTOR",
  "TAMANO_EMPRESA",
  "ESTADO_EMPRESA",
  "TIPO_EMPRESA",
  "FECHA_AFILIACION",
  "ANIO_AFILIACION",
  "VENTAS_MONTO",
  "ANIO_VENTA",
  "CAPACITACIONES_TOMADAS",
  "CAPACITACION_FECHA",
  "CAPACITACION_ANIO",
  "CAPACITACION_VALOR",
  "CAPACITACION_TIPO",
  "DIAGNOSTICO_TOMADO",
  "DIAGNOSTICO_FECHA",
  "DIAGNOSTICO_ANIO",
  "DIAGNOSTICO_TIPO",
  "FUENTE",
  "TIPO_REGISTRO",
];

const FIELD_ALIASES = {
  RUC: ["RUC", "NUMERO_RUC", "NUM_RUC"],
  RAZON: ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "NOMBRE", "RAZON_SOCIA"],
  NOMBRE_COM: ["NOMBRE_COMERCIAL", "NOMBRE COMERCIAL", "COMERCIAL"],
  CIUDAD: ["CIUDAD"],
  SECTOR: ["SECTOR"],
  TAMANO: ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMAÑO", "TAMANO_EMP"],
  ESTADO: ["ESTADO", "ESTADO_EMPRESA", "ESTADO_AFILIADO", "PAGADO", "NO_PAGADO"],
  TIPO_EMPRESA: ["TIPO_EMPRESA", "TIPO DE EMPRESA", "TIPO_EMP"],
  FECHA_AF: ["FECHA_AFILIACION", "FECHA AFILIACION", "FECHA_INGRESO", "FECHA DE INGRESO"],
  FECHA: ["FECHA", "FECHA_EVENTO", "FECHA_DIAGNOSTICO", "FECHA_CAPACITACION", "FECHA CAPACITACION"],
  MONTO: [
    "VENTAS_MONT_EST",
    "MONTO_ESTIMADO",
    "MONTO_VENTAS",
    "MONTO",
    "VALOR TOTAL",
    "VENTAS_ESTIMADAS",
    "VALOR_DEL_PAGO",
    "VALOR_APORTE",
  ],
  ANIO: ["ANIO", "AÑO", "ANO", "ANIO_VENTA", "AÑO_VENTA", "ANIO_VENTAS", "ANIO_AFILIACION", "AÑO_AFILIACION"],
  CAP_TIPO: ["TIPO_CAPACITACION", "TIPO_CAP", "TIPO DE CAPACITACION"],
  DIAG_TIPO: [
    "TIPO_DIAGNOSTICO",
    "TIPO_DIAG",
    "TIPO_DIAGNOSTICO_FINAL",
    "TIPO DIAGNOSTICO",
    "TIPO_DE_ASESORIA",
    "TIPO DE ASESORIA",
  ],
};

// ------------------ Utils ------------------ //
function normalizeLabel(label) {
  let txt = (label || "").toString().trim().toUpperCase();
  txt = txt.replace(/�/g, "N");
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
function toDate(val) {
  if (!val) return "";
  const d = new Date(val);
  return isNaN(d.getTime()) ? "" : Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
}
function getYear(val) {
  const d = new Date(val);
  return isNaN(d.getTime()) ? "" : d.getFullYear().toString();
}
function getSheetOrCreate(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
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
    const nonEmpty = norm.filter((c) => c && c !== "NO" && c !== "LISTA_DE_AFILIADOS_ACTIVOS").length;
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

// ------------------ FACT builder ------------------ //
function generarFactDatosCompleta() {
  const factRows = [];
  const profile = new Map(); // rellena atributos faltantes por RUC
  const processedSheets = new Set();

  function updateProfile(rec) {
    if (!rec.RUC) return;
    const current = profile.get(rec.RUC) || {};
    ["RAZON_SOCIAL", "NOMBRE_COMERCIAL", "CIUDAD", "SECTOR", "TAMANO_EMPRESA", "ESTADO_EMPRESA", "TIPO_EMPRESA"].forEach(
      (k) => {
        if (rec[k]) current[k] = rec[k];
      }
    );
    profile.set(rec.RUC, current);
  }

  function applyProfile(rec) {
    if (!rec.RUC) return rec;
    const prof = profile.get(rec.RUC);
    if (!prof) return rec;
    return {
      ...rec,
      RAZON_SOCIAL: rec.RAZON_SOCIAL || prof.RAZON_SOCIAL || "",
      NOMBRE_COMERCIAL: rec.NOMBRE_COMERCIAL || prof.NOMBRE_COMERCIAL || "",
      CIUDAD: rec.CIUDAD || prof.CIUDAD || "",
      SECTOR: rec.SECTOR || prof.SECTOR || "",
      TAMANO_EMPRESA: rec.TAMANO_EMPRESA || prof.TAMANO_EMPRESA || "",
      ESTADO_EMPRESA: rec.ESTADO_EMPRESA || prof.ESTADO_EMPRESA || "",
      TIPO_EMPRESA: rec.TIPO_EMPRESA || prof.TIPO_EMPRESA || "",
    };
  }

  SOURCES.forEach((src) => {
    const { headerIndex, rows, name: usedName } = readTableFlexibleByCandidates([src.name, src.fallback]);
    if (!usedName || processedSheets.has(usedName)) return;
    processedSheets.add(usedName);

    rows.forEach((row) => {
      const ruc = cleanRuc(getVal(row, headerIndex, FIELD_ALIASES.RUC));
      const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
      const nombreCom = getVal(row, headerIndex, FIELD_ALIASES.NOMBRE_COM);
      if (!ruc && !razon && !nombreCom) return; // saltar filas basura

      const base = {
        RUC: ruc,
        RAZON_SOCIAL: razon,
        NOMBRE_COMERCIAL: nombreCom,
        CIUDAD: getVal(row, headerIndex, FIELD_ALIASES.CIUDAD),
        SECTOR: getVal(row, headerIndex, FIELD_ALIASES.SECTOR),
        TAMANO_EMPRESA: getVal(row, headerIndex, FIELD_ALIASES.TAMANO),
        ESTADO_EMPRESA: getVal(row, headerIndex, FIELD_ALIASES.ESTADO),
        TIPO_EMPRESA: getVal(row, headerIndex, FIELD_ALIASES.TIPO_EMPRESA),
      };

      updateProfile(base); // guarda atributos dim

      let fechaAf = toDate(getVal(row, headerIndex, FIELD_ALIASES.FECHA_AF));
      let anioAf = getYear(fechaAf);

      let ventasMonto = "";
      let anioVenta = "";

      let capTomadas = 0;
      let capFecha = "";
      let capAnio = "";
      let capValor = "";
      let capTipo = "";

      let diagTomado = 0;
      let diagFecha = "";
      let diagAnio = "";
      let diagTipo = "";

      const fechaEvento = toDate(getVal(row, headerIndex, FIELD_ALIASES.FECHA));
      const anioCol = getVal(row, headerIndex, FIELD_ALIASES.ANIO);

      if (src.tipo === "ventas") {
        ventasMonto = toNumber(getVal(row, headerIndex, FIELD_ALIASES.MONTO));
        anioVenta = getYear(fechaEvento) || anioCol || "";
      }

      if (src.tipo === "capacitacion") {
        capTomadas = 1;
        capValor = toNumber(getVal(row, headerIndex, FIELD_ALIASES.MONTO));
        capFecha = fechaEvento;
        capAnio = getYear(capFecha) || anioCol || "";
        capTipo = getVal(row, headerIndex, FIELD_ALIASES.CAP_TIPO);
      }

      if (src.tipo === "diagnostico") {
        diagTomado = 1;
        diagFecha = fechaEvento;
        diagAnio = getYear(diagFecha) || anioCol || "";
        diagTipo = getVal(row, headerIndex, FIELD_ALIASES.DIAG_TIPO);
      }

      const recFinal = applyProfile(base);

      factRows.push([
        recFinal.RUC,
        recFinal.RAZON_SOCIAL,
        recFinal.NOMBRE_COMERCIAL,
        recFinal.CIUDAD,
        recFinal.SECTOR,
        recFinal.TAMANO_EMPRESA,
        recFinal.ESTADO_EMPRESA,
        recFinal.TIPO_EMPRESA,
        fechaAf,
        anioAf,
        ventasMonto,
        anioVenta,
        capTomadas,
        capFecha,
        capAnio,
        capValor,
        capTipo,
        diagTomado,
        diagFecha,
        diagAnio,
        diagTipo,
        src.name,
        src.tipo,
      ]);
    });
  });

  const sh = getSheetOrCreate(SHEET_NAMES.FACT);
  sh.clear();
  const allRows = [FACT_HEADERS, ...factRows];
  if (allRows.length) {
    sh.getRange(1, 1, allRows.length, FACT_HEADERS.length).setValues(allRows);
    sh.autoResizeColumns(1, FACT_HEADERS.length);
  }
  Logger.log("FACT_DATOS_COMPLETA filas (incluye header): %s", allRows.length);
}

// Entrypoints de menú
function refreshFactDatosCompleta() {
  generarFactDatosCompleta();
}
function myFunction() {
  refreshFactDatosCompleta();
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu("FACT").addItem("Generar FACT_DATOS_COMPLETA", "refreshFactDatosCompleta").addToUi();
}

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
    OUT_MASTER: "DASH2_MAESTRA",
  };

  const TAMANO_BY_CODE = { 1: "MICRO", 2: "PEQUENA", 3: "MEDIANA", 4: "GRANDE" };
  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4 };
  const TAMANO_BY_MONTO = [
    { max: 100000, label: "MICRO" },
    { max: 1000000, label: "PEQUENA" },
    { max: 5000000, label: "MEDIANA" },
    { max: Number.POSITIVE_INFINITY, label: "GRANDE" },
  ];
  const USE_BASE_FOR_SIZE = true; // habilitado para usar columnas T202x de SOCIOS
  const USE_REGISTRO_FOR_SIZE = false; // mantenemos deshabilitado registro para no inflar anios
  const MIN_YEAR_FOR_SIZE = 2019; // limitar anios a periodos con ventas

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
    VENTAS: [
      "VENTAS",
      "VENTAS_ANUAL",
      "VENTAS_ANUALES",
      "VENTAS_MONT_EST",
      "MONTO_ESTIMADO",
      "MONTO_VENTAS",
      "MONTO_TOTAL",
      "VALOR TOTAL",
    ],
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
  function normalizeTamanoClean(val) {
    const t = normalizeTamano(val);
    if (!t) return "";
    if (t === "NO REPORTA" || t === "NO_REPORTA" || t === "NO" || t === "N O EPORTA") return "";
    if (TAMANO_ORDER[t]) return t;
    return "";
  }
  function toNumber(val) {
    if (val === null || val === undefined || val === "") return 0;
    let str = val.toString().trim().replace(/\s/g, "");
    const hasComma = str.indexOf(",") !== -1;
    const hasDot = str.indexOf(".") !== -1;
    if (hasComma && hasDot) {
      // Asumimos formato "1,234,567.89": quitar comas (miles) y dejar el punto como decimal
      str = str.replace(/,/g, "");
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
  function detectYearFromHeader(key) {
    if (!key) return "";
    const norm = key.toString().trim().toUpperCase();
    const match = norm.match(/^(?:T_?)?(\d{4})(?:\D.*)?$/); // acepta 2019, T2019, T_2019, 2019_USD
    return match ? match[1] : "";
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
    return normalizeTamanoClean(val);
  }
  function entityKey(headerIndex, row, rowIdx) {
    const ruc = padRuc13(getVal(row, headerIndex, FIELD_ALIASES.RUC));
    const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
    const altId = normalizeKey(getVal(row, headerIndex, FIELD_ALIASES.ALT_ID || []));
    if (altId) {
      if (ruc) return `${ruc}__${altId}`;
      return `ID__${altId}`;
    }
    if (ruc) return ruc;
    if (razon) return razon;
    // Fallback para filas sin identificadores (contar por fila para no perder registros)
    return rowIdx !== undefined ? `ROW_${rowIdx}` : "";
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
    const records = new Map(); // key => {key, ruc, razon, altId, year, tamano, fuente, priority}
    function addRecord({ ruc, razon, altId, year, tamano, fuente, priority }) {
      if (!year || !tamano) return;
      const keyEntity = ruc || normalizeName(razon) || altId;
      if (!keyEntity) return;
      const recKey = `${keyEntity}||${year}`;
      const existing = records.get(recKey);
      if (existing && existing.priority <= priority) return;
      records.set(recKey, { key: keyEntity, ruc, razon, altId, year, tamano, fuente, priority });
    }

    // BASE (SOCIOS) con dos bloques - deshabilitado para limitar a años con ventas
    if (USE_BASE_FOR_SIZE) {
      const ss = SpreadsheetApp.getActive();
      const baseSheetName = [DASH2_SHEETS.BASE, DASH2_SHEETS.BASE_FALLBACK].find((n) => n && ss.getSheetByName(n));
      const shBase = baseSheetName ? ss.getSheetByName(baseSheetName) : null;
      if (shBase) {
        const values = shBase.getDataRange().getDisplayValues();
        const blocks = splitBlocksByHeader(values);
        blocks.forEach((block) => {
          const { headerIndex, rows } = block;
          const yearCols = Object.keys(headerIndex)
            .map((k) => ({ key: k, year: detectYearFromHeader(k) }))
            .filter((y) => y.year);
          rows.forEach((row, idx) => {
              const ruc = padRuc13(getVal(row, headerIndex, FIELD_ALIASES.RUC));
              const razon = normalizeName(getVal(row, headerIndex, FIELD_ALIASES.RAZON));
              const altId = normalizeKey(getVal(row, headerIndex, FIELD_ALIASES.ALT_ID || []));
              const keyRow = entityKey(headerIndex, row, idx);
              yearCols.forEach(({ key: colKey, year }) => {
                const val = row[headerIndex[colKey]];
                let tam = sizeFromCode(val);
              if (!tam) {
                const monto = toNumber(val);
                if (monto > 0) {
                  for (const rule of TAMANO_BY_MONTO) {
                    if (monto <= rule.max) {
                      tam = rule.label;
                      break;
                    }
                  }
                }
              }
              if (tam && parseInt(year, 10) >= MIN_YEAR_FOR_SIZE) {
                addRecord({ ruc, razon, altId: altId || keyRow, year, tamano: tam, fuente: baseSheetName || DASH2_SHEETS.BASE, priority: 1 });
              }
            });
            const tamExplicit = normalizeTamanoClean(getVal(row, headerIndex, FIELD_ALIASES.TAMANO));
            const yearAf = getYear(getVal(row, headerIndex, FIELD_ALIASES.FECHA_AF));
            if (tamExplicit && yearAf && parseInt(yearAf, 10) >= MIN_YEAR_FOR_SIZE) {
              addRecord({ ruc, razon, altId: altId || keyRow, year: yearAf, tamano: tamExplicit, fuente: baseSheetName || DASH2_SHEETS.BASE, priority: 2 });
            }
          });
        });
      }
    }

    // REGISTRO_AFILIADO (si existe) usa FECHA_AFILIACION para asignar anio - deshabilitado para limitar a años con ventas
    if (USE_REGISTRO_FOR_SIZE) {
      const registro = readTableFlexibleByCandidates([DASH2_SHEETS.REGISTRO, DASH2_SHEETS.REGISTRO_FALLBACK]);
      registro.rows.forEach((row, idx) => {
        const ruc = padRuc13(getVal(row, registro.headerIndex, FIELD_ALIASES.RUC));
        const razon = normalizeName(getVal(row, registro.headerIndex, FIELD_ALIASES.RAZON));
        const altId = normalizeKey(getVal(row, registro.headerIndex, FIELD_ALIASES.ALT_ID || []));
        const tam = normalizeTamano(getVal(row, registro.headerIndex, FIELD_ALIASES.TAMANO));
        const year = getYear(getVal(row, registro.headerIndex, FIELD_ALIASES.FECHA_AF));
        const keyRow = entityKey(registro.headerIndex, row, idx);
        if (tam && year) addRecord({ ruc, razon, altId: altId || keyRow, year, tamano: tam, fuente: DASH2_SHEETS.REGISTRO, priority: 3 });
      });
    }

    // VENTAS_AFILIADOS: clasifica por monto o tamano declarado
    const ventas = readTableFlexibleByCandidates([DASH2_SHEETS.VENTAS, DASH2_SHEETS.VENTAS_FALLBACK]);
    ventas.rows.forEach((row, idx) => {
      const ruc = padRuc13(getVal(row, ventas.headerIndex, FIELD_ALIASES.RUC));
      const razon = normalizeName(getVal(row, ventas.headerIndex, FIELD_ALIASES.RAZON));
      const altId = normalizeKey(getVal(row, ventas.headerIndex, FIELD_ALIASES.ALT_ID || []));
      let year = getVal(row, ventas.headerIndex, FIELD_ALIASES.ANIO) || getYear(getVal(row, ventas.headerIndex, FIELD_ALIASES.FECHA_AF));
      let montoCell = getVal(row, ventas.headerIndex, FIELD_ALIASES.VENTAS);

      // Fallback: si no hay año declarado, intentar columnas numéricas (ej. 2019, 2020) con valores
      if (!year) {
        for (const [key, idx] of Object.entries(ventas.headerIndex)) {
          const detectedYear = detectYearFromHeader(key);
          if (detectedYear && idx < row.length && row[idx] !== "") {
            year = detectedYear;
            montoCell = row[idx];
            break;
          }
        }
      }

      let tam = normalizeTamanoClean(getVal(row, ventas.headerIndex, FIELD_ALIASES.TAMANO));
      if (!tam) {
        const monto = toNumber(montoCell);
        if (monto > 0) {
          for (const rule of TAMANO_BY_MONTO) {
            if (monto <= rule.max) {
              tam = rule.label;
              break;
            }
          }
        }
      }
      if (tam && year && parseInt(year, 10) >= MIN_YEAR_FOR_SIZE) {
        const keyRow = entityKey(ventas.headerIndex, row, idx);
        addRecord({ ruc, razon, altId: altId || keyRow, year, tamano: tam, fuente: DASH2_SHEETS.VENTAS, priority: 4 });
      }
    });

    const detailRecords = Array.from(records.values());
    detailRecords.sort((a, b) => {
      const yearA = String(a.year || "");
      const yearB = String(b.year || "");
      if (yearA !== yearB) return yearA.localeCompare(yearB);
      const tamA = String(a.tamano || "");
      const tamB = String(b.tamano || "");
      if (tamA !== tamB) return tamA.localeCompare(tamB);
      return String(a.ruc || "").localeCompare(String(b.ruc || ""));
    });

    const detailRowsForSheet = detailRecords.map((r) => [r.ruc, r.razon, r.year, r.tamano, r.fuente]);
    const shDetail = getSheetOrCreate(DASH2_SHEETS.OUT_DETAIL);
    shDetail.clear();
    const headers = ["RUC", "RAZON_SOCIAL", "ANIO", "TAMANO", "FUENTE"];
    const out = detailRowsForSheet.length ? [headers, ...detailRowsForSheet] : [headers];
    shDetail.getRange(1, 1, out.length, headers.length).setValues(out);
    shDetail.autoResizeColumns(1, headers.length);

    return detailRecords;
  }

  function buildTransitions(detailRows) {
    const byEntity = new Map();
    const isArrayRow = Array.isArray(detailRows[0]);
    detailRows.forEach((row, idx) => {
      const ruc = isArrayRow ? row[0] : row.ruc;
      const razon = isArrayRow ? row[1] : row.razon;
      const tam = isArrayRow ? row[3] : row.tamano;
      const year = parseInt(isArrayRow ? row[2] : row.year, 10);
      if (!year || !tam) return;
      const key = isArrayRow
        ? ruc || normalizeName(razon) || `ROW_${idx}`
        : row.key || row.altId || ruc || normalizeName(razon) || `ROW_${idx}`;
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
        if (!a.tam || !b.tam) continue;
        transitions.push({
          entity: key,
          anioInicial: a.year.toString(),
          anioFinal: b.year.toString(),
          tamInicial: a.tam,
          tamFinal: b.tam,
          delta: (TAMANO_ORDER[b.tam] || 0) - (TAMANO_ORDER[a.tam] || 0),
        });
      }
      // Asegurar que el último año también aparezca como transición consigo mismo (para que salga en el slicer)
      const last = hist[hist.length - 1];
      if (last && last.tam) {
        transitions.push({
          entity: key,
          anioInicial: last.year.toString(),
          anioFinal: last.year.toString(),
          tamInicial: last.tam,
          tamFinal: last.tam,
          delta: 0,
        });
      }
    });

    const originTotals = new Map(); // key anio|tam -> total de empresas con ese tamano en el anio
    detailRows.forEach((row) => {
      const year = isArrayRow ? row[2] : row.year;
      const tam = isArrayRow ? row[3] : row.tamano;
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
      const direccion = delta > 0 ? "CRECIMIENTO" : delta < 0 ? "DECRECIMIENTO" : "SIN_CAMBIO";
      const ordenIni = TAMANO_ORDER[tamIni] || 99;
      const ordenFin = TAMANO_ORDER[tamFin] || 99;
      return [anioIni, anioFin, tamIni, tamFin, ordenIni, ordenFin, count, pct, direccion, delta];
    });

    pivotRows.sort((a, b) => {
      const anioIniA = String(a[0] || "");
      const anioIniB = String(b[0] || "");
      if (anioIniA !== anioIniB) return anioIniA.localeCompare(anioIniB);
      const anioFinA = String(a[1] || "");
      const anioFinB = String(b[1] || "");
      if (anioFinA !== anioFinB) return anioFinA.localeCompare(anioFinB);
      if (a[4] !== b[4]) return a[4] - b[4]; // ORDEN_INICIAL
      if (a[5] !== b[5]) return a[5] - b[5]; // ORDEN_FINAL
      const dirA = String(a[8] || "");
      const dirB = String(b[8] || "");
      if (dirA !== dirB) return dirA.localeCompare(dirB); // DIRECCION
      if (a[9] !== b[9]) return a[9] - b[9]; // DELTA
      const tamA = String(a[2] || "");
      const tamB = String(b[2] || "");
      return tamA.localeCompare(tamB); // TAMANO_INICIAL
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
      const anioA = String(a[0] || "");
      const anioB = String(b[0] || "");
      if (anioA !== anioB) return anioA.localeCompare(anioB);

      const dirA = String(a[8] || "");
      const dirB = String(b[8] || "");
      if (dirA !== dirB) return dirA.localeCompare(dirB); // CRECIMIENTO/DECRECIMIENTO/SIN_CAMBIO

      if (a[4] !== b[4]) return (a[4] || 0) - (b[4] || 0);
      if (a[5] !== b[5]) return (a[5] || 0) - (b[5] || 0);

      const transA = String(a[1] || "");
      const transB = String(b[1] || "");
      return transA.localeCompare(transB);
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

    // Hoja maestra con anio inicial y final (para usar ambos slicers en una sola fuente)
    const masterRows = pivotRows.map((r) => {
      const trans = `${r[2]} -> ${r[3]}`;
      const etiqueta = `${r[6]} (${(r[7] * 100).toFixed(2)}%)`;
      return [r[0], r[1], trans, r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], etiqueta];
    });
    masterRows.sort((a, b) => {
      const anioFinA = String(a[1] || "");
      const anioFinB = String(b[1] || "");
      if (anioFinA !== anioFinB) return anioFinA.localeCompare(anioFinB); // ANIO_FINAL

      const anioIniA = String(a[0] || "");
      const anioIniB = String(b[0] || "");
      if (anioIniA !== anioIniB) return anioIniA.localeCompare(anioIniB); // ANIO_INICIAL

      const dirA = String(a[8] || "");
      const dirB = String(b[8] || "");
      if (dirA !== dirB) return dirA.localeCompare(dirB); // DIRECCION

      if (a[5] !== b[5]) return (a[5] || 0) - (b[5] || 0); // ORDEN_FINAL
      return (a[4] || 0) - (b[4] || 0); // ORDEN_INICIAL
    });
    const shMaster = getSheetOrCreate(DASH2_SHEETS.OUT_MASTER);
    shMaster.clear();
    const headersMaster = [
      "ANIO_INICIAL",
      "ANIO_FINAL",
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
    const outMaster = masterRows.length ? [headersMaster, ...masterRows] : [headersMaster];
    shMaster.getRange(1, 1, outMaster.length, headersMaster.length).setValues(outMaster);
    shMaster.autoResizeColumns(1, headersMaster.length);
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







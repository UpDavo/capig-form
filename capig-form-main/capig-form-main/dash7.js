/**
 * Dashboard de Asesorias Legales - DASH7
 * Integra historico (hoja DIAGNOSTICO - columna LEGAL) + registros Django en DIAGNOSTICO_FINAL.
 * Enfocado solo en tipo LEGAL y sus subtipos (laboral, societario, propiedad intelectual, contacto, otros).
 */
(function (global) {
  const SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    DIAG_DJANGO: "DIAGNOSTICOS",
    DIAG_DJANGO_FALLBACK: "DIAGNOSTICO_FINAL",
    OUT_RESUMEN: "PIVOT_ASESORIAS_RESUMEN_ANIO",
    OUT_SUBTIPO: "PIVOT_ASESORIAS_SUBTIPO_ANIO",
    OUT_EMPRESAS: "PIVOT_ASESORIAS_POR_EMPRESA",
    OUT_MASTER_SLICER: "DASH7_MAESTRA",
  };

  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, SIN_TAMANO: 99 };
  const HIST_YEAR = "HISTORICO";
  const NO_DATE_YEAR = "SIN_FECHA";
  const PROFILES = {
    FULL: {
      // Todo lo legal sin descartar
      requireFechaDjango: false,
      dropSinSubtipo: false,
      allowedSubtipos: [],
      soloSocios: false,
      socioFieldNames: ["SOCIO", "SOCIOS", "ES_SOCIO", "TIPO_SOCIO", "ESTADO_SOCIO", "ESTADO AFILIADO", "ESTADO_AFILIADO"],
      socioYesValues: ["SI", "1", "SOCIO", "SOCIOS", "ACTIVO"],
      socioNoValues: ["NO", "NO SOCIO", "NO SOCIOS", "NO SOCIA", "NO SOCIAS"],
      dedupStrategy: "none", // none | emp-subtipo-anio | emp-anio
    },
    PPT: {
      // Acercar al PPT (36/21 aprox) - INCLUYE NO SOCIOS
      requireFechaDjango: false, // permitir sin fecha, se anclan al anio referencia
      dropSinSubtipo: true, // descartar SIN SUBTIPO/PENDIENTE
      allowedSubtipos: ["LABORAL", "SOCIETARIO", "PROPIEDAD INTELECTUAL", "CONTACTO", "OTROS"],
      soloSocios: false, // INCLUIR NO SOCIOS
      socioFieldNames: ["SOCIO", "SOCIOS", "ES_SOCIO", "TIPO_SOCIO", "ESTADO_SOCIO", "ESTADO AFILIADO", "ESTADO_AFILIADO"],
      socioYesValues: ["SI", "1", "SOCIO", "SOCIOS", "ACTIVO"],
      socioNoValues: ["NO", "NO SOCIO", "NO SOCIOS", "NO SOCIA", "NO SOCIAS"],
      dedupStrategy: "none", // CONTAR TODAS las asesorias
    },
  };
  const PROFILE = "PPT";
  const FILTERS = PROFILES[PROFILE];

  // ---------- Utils (copiados de dash6) ----------
  function normalizeLabel(label) {
    let txt = (label || "").toString().trim().toUpperCase();
    txt = txt.replace(/\u00d1/g, "N");
    txt = txt.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    txt = txt.replace(/[^A-Z0-9]+/g, "_");
    txt = txt.replace(/^_+|_+$/g, "");
    return txt;
  }

  function normalizeName(val) {
    let name = (val || "").toString().trim().toUpperCase();
    name = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    name = name.replace(/\u00d1/g, "N");
    name = name.replace(/[.,\-\/()'"`]/g, " ");
    name = name.replace(/\b(S\s*A|C\s*I\s*A|L\s*T\s*D\s*A|L\s*T\s*D|C\s*A|S\s*A\s*S|SOCIEDAD|ANONIMA|COMPANIA|LIMITADA)\b/gi, " ");
    name = name.replace(/\s+/g, " ").trim();
    return name;
  }

  function cleanRuc(val) {
    return (val || "").toString().replace(/[^0-9]/g, "");
  }

  function isTruthyMark(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return false;
    return t === "X" || t === "S" || t === "SI" || t === "1";
  }

  function isFalsyMark(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return true;
    return t === "0" || t === "-" || t === "NO" || t === "N" || t === "NULL";
  }

  function isYes(val, yesList) {
    const norm = normalizeLabel(val);
    return yesList.some((y) => normalizeLabel(y) === norm);
  }

  function isNo(val, noList) {
    const norm = normalizeLabel(val);
    return noList.some((n) => normalizeLabel(n) === norm);
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

  function getTrimestre(val) {
    const d = parseDateFlexible(val);
    if (!d) return "";
    const month = d.getMonth() + 1;
    return "Q" + Math.ceil(month / 3).toString();
  }

  function normalizeTipoDiagnostico(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return "SIN TIPO";
    const clean = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (clean.indexOf("LEGAL") !== -1) return "LEGAL";
    if (clean.indexOf("ESTRAT") !== -1) return "ESTRATEGIA";
    if (clean.indexOf("LEAN") !== -1) return "LEAN";
    if (clean.indexOf("AMBI") !== -1) return "AMBIENTE";
    if (clean.indexOf("RRHH") !== -1 || clean.indexOf("RH") !== -1 || clean.indexOf("RECURSO") !== -1) return "RRHH";
    if (clean === "NINGUNO") return "NINGUNO";
    return clean;
  }

  function normalizeSubtipoDiagnostico(val) {
    const s = (val || "").toString().trim().toUpperCase();
    if (!s) return "SIN SUBTIPO";
    const clean = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const normalized = clean.replace(/_/g, " ");
    if (normalized.indexOf("PEND") !== -1) return "PENDIENTE";
    if (normalized.indexOf("LABOR") !== -1) return "LABORAL";
    if (normalized.indexOf("PROPIEDAD") !== -1 || normalized.indexOf("INTELECT") !== -1) return "PROPIEDAD INTELECTUAL";
    if (normalized.indexOf("SOCIETA") !== -1) return "SOCIETARIO";
    if (normalized.indexOf("CONTACT") !== -1) return "CONTACTO";
    if (normalized.indexOf("OTRO") !== -1) return "OTROS";
    return normalized;
  }

  function normalizeSector(val) {
    const s = (val || "").toString().trim().toUpperCase();
    if (!s) return "SIN CLASIFICAR";
    if (s.indexOf("QUIM") !== -1) return "QUIMICO";
    if (s.indexOf("METAL") !== -1) return "METALMECANICO";
    if (s.indexOf("ALIMENT") !== -1) return "ALIMENTOS";
    if (s.indexOf("AGRIC") !== -1 || s.indexOf("AGROP") !== -1) return "AGRICOLA";
    if (s.indexOf("MAQUIN") !== -1) return "MAQUINARIAS";
    if (s.indexOf("CONST") !== -1) return "CONSTRUCCION";
    if (s.indexOf("TEXT") !== -1) return "TEXTIL";
    return s;
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
      const nonEmpty = norm.filter((c) => c && c !== "NO").length;
      if ((hasRuc || hasRazon) && nonEmpty >= 2) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) headerRowIdx = 0;

    let headersNorm = values[headerRowIdx].map(normalizeLabel);
    let headerIndex = {};
    headersNorm.forEach((h, idx) => {
      if (h && headerIndex[h] === undefined) headerIndex[h] = idx;
    });

    // Fallback: si no encontramos RAZON_SOCIAL ni RUC, reintentar buscando la primera fila que los tenga
    if (headerIndex["RAZON_SOCIAL"] === undefined && headerIndex["RUC"] === undefined) {
      for (let i = 0; i < values.length; i++) {
        const norm = values[i].map(normalizeLabel);
        if (norm.includes("RAZON_SOCIAL") || norm.includes("RUC")) {
          headerRowIdx = i;
          headersNorm = norm;
          headerIndex = {};
          headersNorm.forEach((h, idx) => {
            if (h && headerIndex[h] === undefined) headerIndex[h] = idx;
          });
          break;
        }
      }
    }

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

  /**
   * Intentar leer la primera hoja disponible de la lista de candidatos.
   */
  function readTableFlexibleByCandidates(names = []) {
    const ss = SpreadsheetApp.getActive();
    for (const name of names) {
      if (!name) continue;
      const sh = ss.getSheetByName(name);
      if (sh) {
        const data = readTableFlexible(name);
        Logger.log(`[DASH7] Usando hoja ${name}: ${data.rows.length} filas`);
        return data;
      }
    }
    Logger.log(`[DASH7] No se encontraron hojas candidatas: ${names.join(", ")}`);
    return { headerIndex: {}, rows: [] };
  }

  // ---------- Core ----------
  function buildAsesoriasLegalesData() {
    Logger.log("[DASH7] Iniciando generacion...");

    const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
    const diagDjango = readTableFlexibleByCandidates([SHEETS.DIAG_DJANGO, SHEETS.DIAG_DJANGO_FALLBACK, "DIAGNOSTICO"]);

    Logger.log(`[DASH7] Filas diagnosticos: ${diagDjango.rows.length} | Filas base: ${base.rows.length}`);

    const baseMap = new Map();
    const aliasToRuc = new Map();

    base.rows.forEach((row) => {
      const ruc = padRuc13(getVal(row, base.headerIndex, ["RUC"]));
      const razon = normalizeName(getVal(row, base.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]));
      if (!ruc && !razon) return;
      const tam = normalizeTamano(getVal(row, base.headerIndex, ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMANO"]));
      const sector = normalizeSector(getVal(row, base.headerIndex, ["SECTOR", "ACTIVIDAD", "SECTOR_PRODUCTIVO"]));
      const socioRaw = getVal(row, base.headerIndex, FILTERS.socioFieldNames || []);
      let socio = true;
      if (FILTERS.soloSocios) {
        if (socioRaw === "") {
          socio = true;
        } else if (isNo(socioRaw, FILTERS.socioNoValues || [])) {
          socio = false;
        } else {
          socio = isYes(socioRaw, FILTERS.socioYesValues || []);
        }
      }
      const key = ruc || razon;
      baseMap.set(key, { ruc, razonSocial: razon || "SIN_NOMBRE", tamano: tam || "SIN_TAMANO", sector: sector || "SIN CLASIFICAR", socio });
      if (razon) aliasToRuc.set(razon, key);
    });
    Logger.log(`[DASH7] Empresas base: ${baseMap.size}`);

    const diagnosticos = [];
    const years = new Set();
    let djangoProcesados = 0;
    let filtNo = 0;
    let filtNinguno = 0;
    let noFound = 0;
    let sinFechaRecuperados = 0;
    let sinFechaDescartados = 0;
    let subtipoDescartados = 0;
    let subtipoFueraLista = 0;
    let noSocioSoloEmpresas = 0;
    let dedupDescartadosSocios = 0;
    let dedupDescartadosNoSocios = 0;
    const dedupSeenSocios = new Set();
    const dedupSeenNoSocios = new Set();
    const diagNoSocios = [];

    // Ano de referencia para filas sin fecha explicita
    let djangoFallbackYear = "";
    diagDjango.rows.forEach((row) => {
      const y = getYear(getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]));
      if (y && (!djangoFallbackYear || y > djangoFallbackYear)) djangoFallbackYear = y;
    });
    if (!djangoFallbackYear) djangoFallbackYear = new Date().getFullYear().toString();
    Logger.log(`[DASH7] Ano referencia DIAGNOSTICOS: ${djangoFallbackYear}`);

    function resolveEmpresa(name) {
      const norm = normalizeName(name);
      const key = aliasToRuc.get(norm) || norm;
      const info = baseMap.get(key);
      if (info) return info;
      const placeholder = { ruc: "", razonSocial: norm || "SIN_NOMBRE", tamano: "SIN_TAMANO", sector: "SIN CLASIFICAR" };
      baseMap.set(key, placeholder);
      if (norm) aliasToRuc.set(norm, key);
      return placeholder;
    }

    // Procesar hoja DIAGNOSTICOS (contiene historico + nuevos)
    diagDjango.rows.forEach((row) => {
      const seDiagRaw = getVal(row, diagDjango.headerIndex, ["SE_DIAGNOSTICO", "SE DIAGNOSTICO", "SE DIAGNOSTICO"]);
      if (!isTruthyMark(seDiagRaw) || isFalsyMark(seDiagRaw)) {
        filtNo++;
        return;
      }
      const tipoRaw = normalizeTipoDiagnostico(getVal(row, diagDjango.headerIndex, ["TIPO_DE_DIAGNOSTICO", "TIPO", "TIPO DIAGNOSTICO", "TIPO DE DIAGNOSTICO"]));
      if (tipoRaw !== "LEGAL") {
        return;
      }

      const fechaRaw = getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]);
      const anioFecha = getYear(fechaRaw);
      if (!anioFecha) sinFechaRecuperados++;
      const subtipoRaw = normalizeSubtipoDiagnostico(
        getVal(row, diagDjango.headerIndex, ["SUBTIPO_DIAGNOSTICO", "SUBTIPO_DE_DIAGNOSTICO", "SUBTIPO", "SUBTIPO DIAGNOSTICO", "SUBTIPO DE DIAGNOSTICO", "OTROS_SUBTIPO", "OTROS SUBTIPO"])
      );
      if (FILTERS.dropSinSubtipo && (!subtipoRaw || subtipoRaw === "SIN SUBTIPO" || subtipoRaw === "PENDIENTE")) {
        subtipoDescartados++;
        return;
      }
      if (FILTERS.allowedSubtipos.length > 0 && !FILTERS.allowedSubtipos.includes(subtipoRaw)) {
        subtipoFueraLista++;
        return;
      }

      const razonRaw = getVal(row, diagDjango.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]);
      if (!razonRaw) {
        noFound++;
        return;
      }

      const empresa = resolveEmpresa(razonRaw);
      if (FILTERS.soloSocios && !empresa.socio) {
        noSocioSoloEmpresas++;
        return;
      }

      const diagObj = {
        ruc: empresa.ruc,
        razonSocial: empresa.razonSocial,
        tamano: empresa.tamano,
        sector: empresa.sector,
        tipoDiagnostico: "LEGAL",
        subtipoDiagnostico: subtipoRaw,
        anio: anioFecha || djangoFallbackYear,
        trimestre: "",
        fuente: "DIAGNOSTICOS",
      };

      diagnosticos.push(diagObj);
      years.add(diagObj.anio);
      djangoProcesados++;
    });

    Logger.log(`[DASH7] Diagnosticos procesados: ${djangoProcesados}`);
    Logger.log("[DASH7] ===== RESUMEN =====");
    Logger.log(`Django procesados: ${djangoProcesados}`);
    Logger.log(`Filtrados 'No': ${filtNo}`);
    Logger.log(`Filtrados 'ninguno': ${filtNinguno}`);
    Logger.log(`No encontrados en base: ${noFound}`);
    Logger.log(`Descartados sin fecha (Django): ${sinFechaDescartados}`);
    Logger.log(`Descartados sin subtipo/pendiente: ${subtipoDescartados}`);
    Logger.log(`Descartados fuera de lista permitida: ${subtipoFueraLista}`);
    Logger.log(`No socios (solo tabla empresas): ${noSocioSoloEmpresas}`);
    Logger.log(`Descartados por dedup socios (${FILTERS.dedupStrategy}): ${dedupDescartadosSocios}`);
    Logger.log(`Descartados por dedup no socios (${FILTERS.dedupStrategy}): ${dedupDescartadosNoSocios}`);
    Logger.log(`Sin fecha (recuperados con fallback): ${sinFechaRecuperados}`);
    Logger.log(`Total asesorias: ${diagnosticos.length}`);

    const yearsList = Array.from(years).sort((a, b) => yearRank(b) - yearRank(a));
    Logger.log(`Anios detectados: ${yearsList.join(", ")}`);

    generateAsesoriasResumenTable(baseMap, diagnosticos, yearsList);
    generateAsesoriasSubtipoTable(diagnosticos, yearsList);
    generateAsesoriasPorEmpresaTable(diagnosticos, yearsList, diagNoSocios);
    generateSlicerMasterSheet();

    Logger.log("[DASH7] Completado");
  }

  function yearRank(val) {
    const num = parseInt(val, 10);
    if (!isNaN(num)) return num;
    if (val === HIST_YEAR) return -1;
    if (val === NO_DATE_YEAR) return -2;
    return -3;
  }

  function generateAsesoriasResumenTable(baseMap, diagnosticos, years) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "TOTAL_ASESORIAS", "EMPRESAS_CON_ASE", "EMPRESAS_SIN_ASE", "EMPRESAS_TOTALES"];
    const empresasPorTam = new Map();
    baseMap.forEach((info) => {
      const tam = info.tamano || "SIN_TAMANO";
      if (!empresasPorTam.has(tam)) empresasPorTam.set(tam, new Set());
      empresasPorTam.get(tam).add(info.ruc || info.razonSocial);
    });
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
        const empresasSet = new Set(diags.map((d) => d.ruc || d.razonSocial));
        const totalEmp = empresasPorTam.get(tam)?.size || 0;
        rows.push([anio, tam, diags.length, empresasSet.size, Math.max(0, totalEmp - empresasSet.size), totalEmp]);
      });
    });

    rows.sort((a, b) => {
      const yr = yearRank(b[0]) - yearRank(a[0]);
      if (yr !== 0) return yr;
      const oA = TAMANO_ORDER[a[1]] || 99;
      const oB = TAMANO_ORDER[b[1]] || 99;
      return oA - oB;
    });

    const sheet = getSheetOrCreate(SHEETS.OUT_RESUMEN);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    Logger.log(`[DASH7] ${SHEETS.OUT_RESUMEN}: ${rows.length} filas`);
  }

  function generateAsesoriasSubtipoTable(diagnosticos, years) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "SUBTIPO", "CANTIDAD", "PCT"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
        const total = diags.length;
        const porSubtipo = new Map();
        diags.forEach((d) => {
          const key = d.subtipoDiagnostico || "SIN SUBTIPO";
          porSubtipo.set(key, (porSubtipo.get(key) || 0) + 1);
        });
        porSubtipo.forEach((count, key) => {
          rows.push([anio, tam, key, count, total > 0 ? (count / total) * 100 : 0]);
        });
      });
    });

    rows.sort((a, b) => {
      const yr = yearRank(b[0]) - yearRank(a[0]);
      if (yr !== 0) return yr;
      const oA = TAMANO_ORDER[a[1]] || 99;
      const oB = TAMANO_ORDER[b[1]] || 99;
      return oA - oB;
    });

    const sheet = getSheetOrCreate(SHEETS.OUT_SUBTIPO);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    Logger.log(`[DASH7] ${SHEETS.OUT_SUBTIPO}: ${rows.length} filas`);
  }

  function generateAsesoriasPorEmpresaTable(diagnosticos, years, diagExtraNoSocios = []) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "RUC", "RAZON_SOCIAL", "SECTOR", "TOTAL_ASESORIAS", "SUBTIPOS_TOMADOS", "RANK"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
        const diagsExtra = diagExtraNoSocios.filter((d) => d.anio === anio && d.tamano === tam);
        const source = diags.concat(diagsExtra);
        const porEmp = new Map();
        source.forEach((d) => {
          const key = d.ruc || d.razonSocial;
          if (!porEmp.has(key)) {
            porEmp.set(key, { ruc: d.ruc, razonSocial: d.razonSocial, sector: d.sector, count: 0, subtipos: new Set() });
          }
          const item = porEmp.get(key);
          item.count++;
          item.subtipos.add(d.subtipoDiagnostico);
        });
        const lista = Array.from(porEmp.values()).sort((a, b) => b.count - a.count);
        lista.forEach((item, idx) => {
          rows.push([anio, tam, item.ruc, item.razonSocial, item.sector, item.count, Array.from(item.subtipos).join(", "), idx + 1]);
        });
      });
    });

    rows.sort((a, b) => {
      const yr = yearRank(b[0]) - yearRank(a[0]);
      if (yr !== 0) return yr;
      const oA = TAMANO_ORDER[a[1]] || 99;
      const oB = TAMANO_ORDER[b[1]] || 99;
      return oA - oB;
    });

    const sheet = getSheetOrCreate(SHEETS.OUT_EMPRESAS);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    Logger.log(`[DASH7] ${SHEETS.OUT_EMPRESAS}: ${rows.length} filas`);
  }

  function generateSlicerMasterSheet() {
    const slicerSources = [
      { name: SHEETS.OUT_RESUMEN, cols: 6 },
      { name: SHEETS.OUT_SUBTIPO, cols: 5 },
      { name: SHEETS.OUT_EMPRESAS, cols: 8 },
    ];

    const ss = SpreadsheetApp.getActive();
    const headerSet = new Set(["SOURCE_SHEET"]);
    const sheetsData = [];

    slicerSources.forEach(({ name, cols }) => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const numRows = sh.getLastRow();
      if (numRows === 0) return;

      const values = sh.getRange(1, 1, numRows, cols).getValues();
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
      headersWritten.forEach((h, idx) => {
        if (h.indexOf("VENTAS_") === 0) moneyCols.push(idx + 1);
        if (h.indexOf("VENTAS_") === 0 && h.endsWith("_M")) moneyCols.push(idx + 1);
      });
      moneyCols.forEach((col) => {
        sheet.getRange(2, col, mergedRows.length, 1).setNumberFormat('"$"#,##0.00');
      });
    }

    Logger.log(`[DASH7] ${SHEETS.OUT_MASTER_SLICER}: ${mergedRows.length} filas (merge para slicers)`);
  }

  function getSheetNameFlexible(preferred) {
    const ss = SpreadsheetApp.getActive();
    if (ss.getSheetByName(preferred)) return preferred;
    const sheets = ss.getSheets();
    const preferredNorm = normalizeLabel(preferred);
    let best = null;
    let bestScore = -1;

    function scoreSheet(sheetName) {
      const info = readTableFlexible(sheetName);
      const rowCount = info.rows.length;
      if (rowCount === 0) return null;
      const hasLegalCols =
        info.headerIndex["LEGAL"] !== undefined ||
        info.headerIndex["LEGALES"] !== undefined ||
        info.headerIndex["DIAGNOSTICO_LEGAL"] !== undefined ||
        info.headerIndex["ASPECTOS LEGALES"] !== undefined ||
        info.headerIndex["JURIDICO"] !== undefined ||
        info.headerIndex["D_LEGAL"] !== undefined;
      return { rowCount, hasLegalCols };
    }

    sheets.forEach((s) => {
      const name = s.getName();
      const norm = normalizeLabel(name);
      if (norm.indexOf("FINAL") !== -1) return;
      if (norm.indexOf("PIVOT") !== -1) return;
      const meta = scoreSheet(name);
      if (!meta) return;
      const { rowCount, hasLegalCols } = meta;
      if ((norm === preferredNorm || norm === "DIAGNOSTICOS") && rowCount > 0 && hasLegalCols) {
        best = name;
        bestScore = rowCount + 1000;
        return;
      }
      if (norm.indexOf("DIAGNOST") !== -1 && hasLegalCols && rowCount >= 5) {
        const score = rowCount;
        if (score > bestScore) {
          best = name;
          bestScore = score;
        }
      }
    });
    return best || null;
  }

  function refreshDashboardAsesorias() {
    buildAsesoriasLegalesData();
  }

  function onOpenDash7() {
    SpreadsheetApp.getUi().createMenu("DASH7").addItem("Generar asesorias", "refreshDashboardAsesorias").addToUi();
  }

  // Trigger manual via checkbox (hoja REPORTE_2, celda Q53)
  const DASH7_TRIGGER_SHEET = "REPORTE_2";
  const DASH7_TRIGGER_CELL = "Q53";

  function onEditDash7(e) {
    const range = e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== DASH7_TRIGGER_SHEET) return;
    if (range.getA1Notation() !== DASH7_TRIGGER_CELL) return;

    const val = range.getValue();
    if (val === true) {
      const ss = SpreadsheetApp.getActive();
      ss.toast("Actualizando dashboard…");
      try {
        refreshDashboardAsesorias();
        ss.toast("Dashboard listo.");
        sheet.getRange("Q54").setValue("Ultima actualizacion: " + new Date());
      } finally {
        range.setValue(false); // deja la casilla lista para el siguiente clic
      }
    }
  }

  global.refreshDashboardAsesorias = refreshDashboardAsesorias;
  global.onOpenDash7 = onOpenDash7;
  global.onEditDash7 = onEditDash7;
})(this);

function refreshDashboardAsesoriasWrapper() {
  return refreshDashboardAsesorias();
}

// Wrapper para que el trigger instalable detecte la función onEditDash7
function onEditDash7Wrapper(e) {
  return onEditDash7(e);
}

// Wrapper opcional para onOpenDash7 si se requiere trigger instalable
function onOpenDash7Wrapper() {
  return onOpenDash7();
}









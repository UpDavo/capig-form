/**
 * Dashboard de Asesorias Legales - DASH7
 * Integra historico (DIAGNOSTICOS_HISTORICOS/DIAGNOSTICO) + registros Django (DIAGNOSTICOS/DIAGNOSTICO_FINAL).
 * Solo tipo LEGAL (asesorias) con subtipos laborales/societario/propiedad intelectual/contacto/otros.
 * Denominador: empresas presentes en diagnosticos/historicos (incluye NO); Numerador: filas LEGAL con conteo positivo.
 */
(function (global) {
  const SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    DIAG_DJANGO: "ASESORIAS",
    DIAG_DJANGO_FALLBACK: "DIAGNOSTICOS",
    DIAG_HIST: "DIAGNOSTICOS_HISTORICOS",
    LEGAL_UNIFICADO: "LEGAL_UNIFICADO",
    OUT_RESUMEN: "PIVOT_ASESORIAS_RESUMEN_ANIO",
    OUT_SUBTIPO: "PIVOT_ASESORIAS_SUBTIPO_ANIO",
    OUT_EMPRESAS: "PIVOT_ASESORIAS_POR_EMPRESA",
    OUT_MASTER_SLICER: "DASH7_MAESTRA",
  };

  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, SIN_TAMANO: 99 };
  const HIST_YEAR = "HISTORICO";
  const NO_DATE_YEAR = "SIN_FECHA";
  const DEFAULT_HIST_YEAR = "2025";
  const TYPE_COLUMNS = ["LEAN", "ESTRATEGIA", "LEGAL", "AMBIENTE", "RRHH"];
  const ALLOWED_SUBTIPOS = ["LABORAL", "SOCIETARIO", "PROPIEDAD INTELECTUAL", "OTROS"];
  const LEGAL_SHEETS = ["LEGAL 1", "LEGAL1", "LEGAL_1", "LEGAL 2", "LEGAL2", "LEGAL_2"];
  const LEGAL_PROP_INT_ALIASES = ["PROPIEDAD_INTELECTUAL", "PROPIEDAD INTELECTUAL", "INTELECTUAL"];
  const LEGAL_TAMANO_ALIASES = ["TAMANO", "TAMANIO"];
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
  const TOTAL_ASESORIAS_ALIASES = [
    "TOTAL",
    "TOTAL_ASESORIAS",
    "TOTAL ASESORIAS",
    "TOTAL_ASESORIA",
    "TOTAL SERVICIO LEGAL",
    "TOTAL_SERVICIO_LEGAL",
    "TOTAL LEGAL",
  ];
  const LEGAL_MARKER_ALIASES = [
    "LEGAL",
    "LEGAL_",
    "LEGAL1",
    "LEGAL 1",
    "LEGAL_1",
    "LEGAL2",
    "LEGAL 2",
    "LEGAL_2",
    "LEGAL DIAGNOSTICO",
    "LEGAL_DIAGNOSTICO",
    "LEGAL_SERVICIO",
    "LEGAL SERVICIO",
    "LEGAL_SERVICIOS",
    "LEGAL SERVICIOS",
    "ASESORIA_LEGAL",
    "ASESORIA LEGAL",
    "ASESORIAS_LEGALES",
    "ASESORIA LEGAL 1",
    "ASESORIA LEGAL 2",
  ];
  const SERVICIO_LEGAL_ALIASES = [
    "SERVICIO_LEGAL",
    "SERVICIO LEGAL",
    "SERVICIO LEGAL?",
    "SERVICIOS_LEGALES",
    "SERVICIO LEGAL 1",
    "SERVICIO LEGAL 2",
  ];
    const SERVICIO_FLAG_ALIASES = ["SERVICIO", "SERVICIO LEGAL", "SERVICIO LEGAL?", "DIAGNOSTICO", "DIAGNOSTICO", "DIAGNOSTICO?"];
  const SUBTIPO_ALIASES = [
    "SUBTIPO_DIAGNOSTICO",
    "SUBTIPO_DE_DIAGNOSTICO",
    "SUBTIPO",
    "SUBTIPO DIAGNOSTICO",
    "SUBTIPO DE DIAGNOSTICO",
    "SUBTIPO_DE_ASESORIA",
    "SUBTIPO DE ASESORIA",
    "SUBTIPO ASESORIA",
    "OTROS_SUBTIPO",
    "OTROS SUBTIPO",
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

  function normalizeKey(val) {
    return (val || "").toString().trim().toUpperCase().replace(/\s+/g, "_");
  }

  function normalizeName(val) {
    let name = (val || "").toString().trim().toUpperCase();
    name = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    name = name.replace(/\u00d1|\u00f1/g, "N");
    name = name.replace(/[.,\-\/()'"`]/g, " ");
    name = name.replace(/\b(S\s*A|C\s*I\s*A|L\s*T\s*D\s*A|L\s*T\s*D|C\s*A|S\s*A\s*S|SOCIEDAD|ANONIMA|COMPANIA|LIMITADA)\b/gi, " ");
    name = name.replace(/\s+/g, " ").trim();
    return name;
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
    const compact = clean.replace(/\s+/g, "");
    if (compact === "NINGUNO" || compact === "NINGUNA" || compact === "NINGUN") return "NINGUNO";
    if (clean.indexOf("LEGAL") !== -1) return "LEGAL";
    if (clean.indexOf("ESTRAT") !== -1) return "ESTRATEGIA";
    if (clean.indexOf("LEAN") !== -1) return "LEAN";
    if (clean.indexOf("AMBI") !== -1) return "AMBIENTE";
    if (clean.indexOf("RRHH") !== -1 || clean.indexOf("RH") !== -1 || clean.indexOf("RECURSO") !== -1) return "RRHH";
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

  function normalizeSubtipoLegal(val) {
    const norm = normalizeSubtipoDiagnostico(val);
    if (!norm || norm === "SIN SUBTIPO" || norm === "PENDIENTE") return "SIN SUBTIPO";
    if (ALLOWED_SUBTIPOS.includes(norm)) return norm;
    return "OTROS";
  }

  function isMarkerEmpty(val) {
    const t = (val || "").toString().trim().toUpperCase();
    if (!t) return true;
    const compact = t.replace(/\s+/g, "");
    return compact === "X" || compact === "-" || compact === "N/A" || compact === "NA" || compact === "N\\A";
  }

  function shouldSkipDiagnosticoFlag(flag) {
    const t = (flag || "").toString().trim().toUpperCase();
    if (!t) return true;
    if (isMarkerEmpty(t)) return true;
    return t.startsWith("NO");
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

  function parseCountFlexible(val) {
    if (val === null || val === undefined) return 0;
    const raw = val.toString().trim();
    if (!raw) return 0;
    const num = Number(raw);
    if (!isNaN(num)) return num > 0 ? num : 0;
    const upperNorm = raw.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (upperNorm === "SI" || upperNorm === "TRUE" || upperNorm === "X") return 1;
    return 0;
  }

  function parseCountNumber(val) {
    if (val === null || val === undefined) return 0;
    const raw = val.toString().trim();
    if (!raw) return 0;
    const num = Number(raw);
    if (!isNaN(num)) return num > 0 ? num : 0;
    const upperNorm = raw.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (upperNorm === "SI" || upperNorm === "TRUE" || upperNorm === "X") return 1;
    return 0;
  }

  function collectLegalCount(row, headerIndex, tipoRaw) {
    if (tipoRaw !== "LEGAL") {
      return { total: 0, fromTotal: 0, fromLegal: 0, fromServicio: 0, hasServicio: false };
    }
    const fromTotal = parseCountFlexible(getVal(row, headerIndex, TOTAL_ASESORIAS_ALIASES));
    let fromLegal = 0;
    LEGAL_MARKER_ALIASES.forEach((alias) => {
      const v = getVal(row, headerIndex, [alias]);
      fromLegal += parseCountFlexible(v);
    });
    let fromServicio = 0;
    SERVICIO_LEGAL_ALIASES.forEach((alias) => {
      const v = getVal(row, headerIndex, [alias]);
      fromServicio += parseCountFlexible(v);
    });
    let total = fromTotal > 0 ? fromTotal : fromLegal + fromServicio;
    const hasServicio = fromServicio > 0 || fromTotal > 0 || fromLegal > 0;
    if (total < 0) total = 0;
    return { total, fromTotal, fromLegal, fromServicio, hasServicio };
  }

  function applyManualAliases(aliasToRuc, baseMap) {
    const manual = [
      { ruc: "0991300333001", aliases: ["JAZUL", "TONISA", "TONISA S.A.", "TONISA SA"] },
      { ruc: "0992257946001", aliases: ["MUNDOCARE", "MUNDOCARE S.A.", "MUNDOCARE SA", "ECUASERVIGLOBAL", "ECUASERVIGLOBAL S.A."] },
      { ruc: "0991318380001", aliases: ["CORDOVA DONOSO SONIA SALOME", "CONSTRUME", "CONSTRUCCIONES CIVILES Y METALICAS CONSTRUME", "CONSTRUCCIONES CIVILES Y METALICAS CONSTRUME S.A."] },
    ];
    manual.forEach((entry) => {
      const info = baseMap.get(entry.ruc);
      if (!info) return;
      entry.aliases.forEach((a) => {
        const norm = normalizeName(a);
        aliasToRuc.set(norm, entry.ruc);
      });
    });
  }

  // ---------- Core ----------
  function buildAsesoriasLegalesData() {
    Logger.log("[DASH7] Iniciando generacion...");

    const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
    const diagDjango = readTableFlexibleByCandidates([
      SHEETS.DIAG_DJANGO,
      SHEETS.DIAG_DJANGO_FALLBACK,
      "DIAGNOSTICO_FINAL",
      "DIAGNOSTICO",
    ]);
    const diagHist = { headerIndex: {}, rows: [] }; // historico fuera de uso (ya consolidado en DIAGNOSTICOS)
    const legalUnified = readTableFlexibleByCandidates([SHEETS.LEGAL_UNIFICADO, ...LEGAL_SHEETS]);
    const legalTables = LEGAL_SHEETS.map((name) => ({ name, ...readTableFlexible(name) })).filter((t) => t.rows.length);
    const legalSources = legalUnified.rows.length ? [legalUnified] : legalTables;
    const useLegalSheetsOnly = legalSources.some((t) => t.rows.length > 0);

    Logger.log(`[DASH7] Filas ${SHEETS.DIAG_DJANGO}: ${diagDjango.rows.length} | Filas historico: ${diagHist.rows.length}`);

    const baseMap = new Map();
    const aliasToRuc = new Map();
    const altIdToBase = new Map();
    const missingEmpresas = new Map();
    const empresaTamAsignado = new Map();
    const empresasBaseAsePorTam = new Map();
    const empresasRawAsePorTam = new Map();
    const empresasLegalDenomPorTam = new Map();

    base.rows.forEach((row) => {
      const ruc = padRuc13(getVal(row, base.headerIndex, ["RUC"]));
      const altId = normalizeKey(getVal(row, base.headerIndex, ALT_ID_ALIASES));
      const razon = normalizeName(getVal(row, base.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]));
      if (!ruc && !altId && !razon) return;
      const tam = normalizeTamano(getVal(row, base.headerIndex, ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMANO_EMPRESA_", "TAMANO"]));
      const sector = normalizeSector(getVal(row, base.headerIndex, ["SECTOR", "ACTIVIDAD", "SECTOR_PRODUCTIVO"]));
      const key = ruc || altId || razon;
      const data = {
        baseKey: key,
        ruc,
        altId,
        razonSocial: razon || "SIN_NOMBRE",
        tamano: tam || "SIN_TAMANO",
        sector: sector || "SIN CLASIFICAR",
      };
      baseMap.set(key, data);
      if (altId) altIdToBase.set(altId, data);
      if (razon) aliasToRuc.set(razon, key);
    });
    Logger.log(`[DASH7] Empresas base: ${baseMap.size}`);

    // En modo solo LEGAL evitamos alias manuales para no colapsar empresas distintas.
    if (!useLegalSheetsOnly) {
      applyManualAliases(aliasToRuc, baseMap);
    }

    const asesoriasBuckets = new Map(); // key: entity|anio (dedup por empresa y anio)
    const asesorias = [];
    const years = new Set();
    let registrosProcesados = 0;
    let filtNo = 0;
    let filtNinguno = 0;
    let noFound = 0;
    let sinFechaRecuperados = 0;
    let bucketRowsMerged = 0;
    let debugUsefulRowsDjango = 0;
    let debugUsefulRowsHist = 0;
    let legalRowsProcessed = 0;
    let djangoFallbackYear = "";
    diagDjango.rows.forEach((row) => {
      const y = getYear(getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]));
      if (y && (!djangoFallbackYear || y > djangoFallbackYear)) djangoFallbackYear = y;
    });
    if (!djangoFallbackYear) djangoFallbackYear = DEFAULT_HIST_YEAR;
    const histFallbackYear = DEFAULT_HIST_YEAR;

    const hasHistInsideDjango = diagDjango.rows.some(
      (row) => !getYear(getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]))
    );
    const shouldProcessHist = false;
    Logger.log(
      `[DASH7] Historico embebido en ${SHEETS.DIAG_DJANGO}: ${hasHistInsideDjango ? "SI" : "NO"} | Procesar hoja historica aparte: NO (solo ${SHEETS.DIAG_DJANGO})`
    );
    Logger.log(`[DASH7] Ano fallback Django: ${djangoFallbackYear} | Ano fallback Historico: ${histFallbackYear}`);

    function yearOrDefault(fechaRaw, fuente, fallbackYear) {
      const y = getYear(fechaRaw);
      if (y) return y;
      if (fuente === "HIST") return fallbackYear || HIST_YEAR;
      if (fuente === "DJANGO") return fallbackYear || NO_DATE_YEAR;
      return fallbackYear || NO_DATE_YEAR;
    }

  function resolveEmpresa(name, rowForAltId, rowHeaderIndex, preferRawForLegal = false) {
    const norm = normalizeName(name);
    // En modo LEGAL priorizamos RUC/ALT, pero luego usamos las mismas concordancias por nombre/alias
    if (preferRawForLegal && rowForAltId && rowHeaderIndex) {
      const rawRuc = padRuc13(getVal(rowForAltId, rowHeaderIndex, ["RUC"]));
      if (rawRuc) {
        const byRuc = baseMap.get(rawRuc);
        if (byRuc) return byRuc;
      }
      const rawAlt = normalizeKey(getVal(rowForAltId, rowHeaderIndex, ALT_ID_ALIASES));
      if (rawAlt) {
        const byAlt = baseMap.get(rawAlt) || altIdToBase.get(rawAlt);
        if (byAlt) return byAlt;
      }
      // si no hay RUC/ALT, continuamos al flujo normal (alias/nombre) antes de crear placeholder
    }

    const mappedKey = aliasToRuc.get(norm);
    if (mappedKey) {
      const infoMapped = baseMap.get(mappedKey) || missingEmpresas.get(mappedKey);
      if (infoMapped) return infoMapped;
    }

    const infoByName = baseMap.get(norm);
    if (infoByName) {
      aliasToRuc.set(norm, infoByName.baseKey);
      return infoByName;
    }

    if (rowForAltId) {
      const altId = normalizeKey(getVal(rowForAltId, rowHeaderIndex || {}, ALT_ID_ALIASES));
      if (altId) {
        const infoAlt = altIdToBase.get(altId);
        if (infoAlt) {
          aliasToRuc.set(norm, infoAlt.baseKey);
          return infoAlt;
        }
      }
    }

    const placeholder = {
      baseKey: norm,
      ruc: "",
      altId: "",
      razonSocial: norm || "SIN_NOMBRE",
      tamano: "SIN_TAMANO",
      sector: "SIN CLASIFICAR",
      isPlaceholder: true,
    };
    if (norm) {
      missingEmpresas.set(norm, placeholder);
      aliasToRuc.set(norm, placeholder.baseKey);
    }
    return placeholder;
  }

    function addAsesoriaAggregated(empresa, count, fechaRaw, fuente, fallbackYear, subtipoRaw) {
      let subtipo = normalizeSubtipoLegal(subtipoRaw);
      if (!subtipo || subtipo === "SIN SUBTIPO" || subtipo === "PENDIENTE") subtipo = "OTROS"; // reubicar sin sub
      const anio = yearOrDefault(fechaRaw, fuente, fallbackYear);
      const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
      const key = `${entityKey}|${anio}`;
      let bucket = asesoriasBuckets.get(key);
      const prevTotal = bucket ? bucket.total || 0 : 0;
      if (!bucket) {
        bucket = {
          ruc: empresa.ruc,
          razonSocial: empresa.razonSocial,
          tamano: empresa.tamano,
          sector: empresa.sector,
          anio,
          trimestre: getTrimestre(fechaRaw),
          fuente,
          total: 0,
          subtipos: new Map(),
        };
        asesoriasBuckets.set(key, bucket);
        years.add(anio);
      } else {
        if (!bucket.trimestre) bucket.trimestre = getTrimestre(fechaRaw);
        if (!bucket.fuente && fuente) bucket.fuente = fuente;
      }
      const newTotal = Math.max(bucket.total || 0, count);
      const added = Math.max(0, newTotal - (bucket.total || 0));
      bucket.total = newTotal;

      const prevSub = bucket.subtipos.get(subtipo) || 0;
      const newSub = Math.max(prevSub, count);
      bucket.subtipos.set(subtipo, newSub);

      return { added, merged: count - added > 0 ? count - added : 0 };
    }

    function addAsesoriaAggregatedFromSubtipos(empresa, subtiposMap, totalRaw, fechaRaw, fuente, fallbackYear) {
      const totalSub = Array.from(subtiposMap.values()).reduce((acc, v) => acc + (v || 0), 0);
      let total = parseCountNumber(totalRaw);
      if (total <= 0) total = totalSub;
      if (total <= 0 && totalSub <= 0) return { added: 0, merged: 0 };

      const anio = yearOrDefault(fechaRaw, fuente, fallbackYear);
      const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
      const key = `${entityKey}|${anio}`;
      let bucket = asesoriasBuckets.get(key);
      const prevTotal = bucket ? bucket.total || 0 : 0;
      if (!bucket) {
        bucket = {
          ruc: empresa.ruc,
          razonSocial: empresa.razonSocial,
          tamano: empresa.tamano,
          sector: empresa.sector,
          anio,
          trimestre: getTrimestre(fechaRaw),
          fuente,
          total: 0,
          subtipos: new Map(),
        };
        asesoriasBuckets.set(key, bucket);
        years.add(anio);
      } else {
        if (!bucket.trimestre) bucket.trimestre = getTrimestre(fechaRaw);
        if (!bucket.fuente && fuente) bucket.fuente = fuente;
      }

      const newTotal = Math.max(bucket.total || 0, total);
      const added = Math.max(0, newTotal - (bucket.total || 0));
      bucket.total = newTotal;

      subtiposMap.forEach((cnt, subtipoRaw) => {
        const subtipo = normalizeSubtipoLegal(subtipoRaw);
        if (!subtipo || subtipo === "SIN SUBTIPO") return;
        const prevSub = bucket.subtipos.get(subtipo) || 0;
        bucket.subtipos.set(subtipo, Math.max(prevSub, cnt || 0));
      });

      return { added, merged: total - added > 0 ? total - added : 0 };
    }

    function markEmpresaDenominador(empresa, row, headerIndex, hasUsefulData, collectLegalDenom = false) {
      if (!empresa || !hasUsefulData) return false;
      const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
      const tamAsignado = empresaTamAsignado.get(entityKey) || empresa.tamano || "SIN_TAMANO";
      if (!empresaTamAsignado.has(entityKey)) empresaTamAsignado.set(entityKey, tamAsignado);
      const tamKey = tamAsignado || "SIN_TAMANO";
      if (collectLegalDenom) {
        if (!empresasLegalDenomPorTam.has(tamKey)) empresasLegalDenomPorTam.set(tamKey, new Set());
        empresasLegalDenomPorTam.get(tamKey).add(entityKey);
      }
      if (!empresa.isPlaceholder) {
        if (!empresasBaseAsePorTam.has(tamKey)) empresasBaseAsePorTam.set(tamKey, new Set());
        empresasBaseAsePorTam.get(tamKey).add(entityKey);
      }

      let rawRuc = "";
      let rawAlt = "";
      let rawName = "";
      if (row) {
        rawRuc = padRuc13(getVal(row, headerIndex, ["RUC"]));
        rawAlt = normalizeKey(getVal(row, headerIndex, ALT_ID_ALIASES));
        rawName = normalizeName(getVal(row, headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "Razon Social"]));
      } else {
        rawRuc = padRuc13(empresa.ruc || "");
        rawName = normalizeName(empresa.razonSocial || "");
      }
      const rawKey = rawRuc || rawAlt || rawName;
      // Normalizamos el identificador crudo para evitar duplicados por variaciones de escritura.
      const keyForRaw = normalizeName(rawKey || entityKey);
      if (keyForRaw) {
        if (!empresasRawAsePorTam.has(tamKey)) empresasRawAsePorTam.set(tamKey, new Set());
        empresasRawAsePorTam.get(tamKey).add(keyForRaw);
      }
      return true;
    }

    if (useLegalSheetsOnly && legalSources.length) {
      Logger.log(`[DASH7] Usando hojas legales: fuentes=${legalSources.length}`);
      const legalMerged = new Map(); // key: empresa -> maxima data combinando todas las hojas legales

      legalSources.forEach((table) => {
        Logger.log(`[DASH7] Procesando hoja legal ${table.name || "LEGAL_UNIFICADO"}: ${table.rows.length} filas`);
        table.rows.forEach((row) => {
          let razonRaw = getVal(row, table.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA"]);
          if (!razonRaw) razonRaw = "SIN_NOMBRE_LEGAL";
          const empresa = resolveEmpresa(razonRaw, row, table.headerIndex, true);
          if (!empresa) {
            noFound++;
            return;
          }
          if (empresa.isPlaceholder) {
            const tamLegal = normalizeTamano(getVal(row, table.headerIndex, LEGAL_TAMANO_ALIASES));
            if (tamLegal) empresa.tamano = tamLegal;
          }
          const subtiposMap = new Map();
          subtiposMap.set("LABORAL", parseCountNumber(getVal(row, table.headerIndex, ["LABORAL"])) || 0);
          subtiposMap.set("SOCIETARIO", parseCountNumber(getVal(row, table.headerIndex, ["SOCIETARIO"])) || 0);
          subtiposMap.set("PROPIEDAD INTELECTUAL", parseCountNumber(getVal(row, table.headerIndex, LEGAL_PROP_INT_ALIASES)) || 0);
          subtiposMap.set("OTROS", parseCountNumber(getVal(row, table.headerIndex, ["OTROS"])) || 0);
          const totalRaw = getVal(row, table.headerIndex, TOTAL_ASESORIAS_ALIASES);
          const servicioRaw = getVal(row, table.headerIndex, [...SERVICIO_LEGAL_ALIASES, "SERVICIO"]);
          const diagFlagRaw = getVal(row, table.headerIndex, ["DIAGNOSTICO", "DIAGNÓSTICO"]);
          let servicioNorm = (servicioRaw || "").toString().trim().toUpperCase();
          const diagFlagNorm = (diagFlagRaw || "").toString().trim().toUpperCase();
          if (!servicioNorm && diagFlagNorm) servicioNorm = diagFlagNorm;
          const totalNumRaw = parseCountNumber(totalRaw);
          const sumSubtipos = Array.from(subtiposMap.values()).reduce((acc, v) => acc + (v || 0), 0);
          const totalNum = totalNumRaw > 0 ? totalNumRaw : sumSubtipos;
          const hasAnyLegalData = (servicioNorm && servicioNorm.length) || totalNum > 0 || sumSubtipos > 0;
          if (!hasAnyLegalData) return;
          // Denominador
          markEmpresaDenominador(empresa, row, table.headerIndex, true, true);
          // Numerador: consolidar por empresa tomando max por subtipo y total para evitar duplicados entre hojas
          if (servicioNorm && servicioNorm.startsWith("NO")) return;
          if (!servicioNorm && totalNum > 0) servicioNorm = "SI";
          if (servicioNorm && (servicioNorm.startsWith("NO") || servicioNorm === "0")) return;
          if (servicioNorm && servicioNorm !== "SI") return;
          if (totalNum <= 0 && sumSubtipos <= 0) return;
          const fechaRaw = getVal(row, table.headerIndex, ["FECHA", "ANIO", "ANIO_SERVICIO", "ANO"]);
          const key = empresa.baseKey || empresa.ruc || empresa.razonSocial;
          if (!key) return;
          if (!legalMerged.has(key)) {
            legalMerged.set(key, {
              empresa,
              subtipos: new Map(),
              total: 0,
              fecha: fechaRaw,
            });
          }
          const target = legalMerged.get(key);
          target.total = Math.max(target.total || 0, totalNum);
          subtiposMap.forEach((cnt, st) => {
            const prev = target.subtipos.get(st) || 0;
            target.subtipos.set(st, Math.max(prev, cnt || 0));
          });
          if (!target.fecha && fechaRaw) target.fecha = fechaRaw;
        });
      });

      legalMerged.forEach((entry) => {
        const res = addAsesoriaAggregatedFromSubtipos(
          entry.empresa,
          entry.subtipos,
          entry.total,
          entry.fecha,
          legalSources[0].name || "LEGAL_UNIFICADO",
          DEFAULT_HIST_YEAR
        );
        registrosProcesados += res.added;
        bucketRowsMerged += res.merged;
        legalRowsProcessed += res.added;
        const hasData = res.added > 0 || Array.from(entry.subtipos.values()).some((v) => v > 0);
        if (hasData) markEmpresaDenominador(entry.empresa, null, {}, true);
      });
    }

    if (!useLegalSheetsOnly) {
      diagDjango.rows.forEach((row) => {
      const razonRaw = getVal(row, diagDjango.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "Razon Social"]);
      const empresa = resolveEmpresa(razonRaw, row, diagDjango.headerIndex);
      if (!empresa) {
        noFound++;
        return;
      }

      const seDiagRaw = (
        getVal(row, diagDjango.headerIndex, [
          "SE_DIAGNOSTICO",
          "SE DIAGNOSTICO",
          "SE_DIAGNOSTICO_",
          "SE DIAGNOSTICO_",
          "TOTAL_DIAGNOSTICO",
          "TOTAL DIAGNOSTICO",
          "TOTAL DIAGNOSTICOS",
          "TOTAL DIAGNOSTICOS",
          "DIAGNOSTICO",
        ]) || ""
      )
        .toString()
        .trim()
        .toUpperCase();

      const tipoRaw = normalizeTipoDiagnostico(
        getVal(row, diagDjango.headerIndex, [
          "TIPO_DE_DIAGNOSTICO",
          "TIPO_DE_ASESORIA",
          "TIPO",
          "TIPO DIAGNOSTICO",
          "TIPO DE DIAGNOSTICO",
          "TIPO DE ASESORIA",
        ])
      );
      if (tipoRaw !== "LEGAL") {
        if (tipoRaw === "NINGUNO") filtNinguno++;
        return;
      }
      const subtipoRaw = getVal(row, diagDjango.headerIndex, SUBTIPO_ALIASES);
      const legalCounts = collectLegalCount(row, diagDjango.headerIndex, tipoRaw);
      const servicioMarcado =
        legalCounts.hasServicio ||
        parseCountFlexible(seDiagRaw) > 0 ||
        parseCountFlexible(getVal(row, diagDjango.headerIndex, SERVICIO_FLAG_ALIASES)) > 0;
      // Si es LEGAL, contar salvo que explicitamente marque NO (o similar).
      const tookLegal = tipoRaw === "LEGAL" && (servicioMarcado || !shouldSkipDiagnosticoFlag(seDiagRaw));
      const fechaRaw = getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]);
      const hasDate = !!getYear(fechaRaw);

      const tiposDet = TYPE_COLUMNS.map((col) => parseCountFlexible(getVal(row, diagDjango.headerIndex, [col]))).filter((v) => v > 0);
      const hasLegalData = tookLegal; // Solo contamos si realmente tomo servicio legal
      const hasUsefulData = hasLegalData || !!seDiagRaw || tiposDet.length > 0;

      if (markEmpresaDenominador(empresa, row, diagDjango.headerIndex, hasUsefulData)) debugUsefulRowsDjango++;

      if (!hasLegalData && shouldSkipDiagnosticoFlag(seDiagRaw)) {
        filtNo++;
        return;
      }
      if (!hasLegalData) return;

      const count = tookLegal ? Math.max(legalCounts.total || 0, 1) : 0;
      if (!hasDate) sinFechaRecuperados += count;
      if (count === 0) return;
      const result = addAsesoriaAggregated(
        empresa,
        count,
        fechaRaw,
        hasDate ? "DJANGO" : "HIST",
        hasDate ? djangoFallbackYear : histFallbackYear,
        subtipoRaw
      );
      registrosProcesados += result.added;
      bucketRowsMerged += result.merged;
      });
    }

    if (shouldProcessHist && !useLegalSheetsOnly) {
      Logger.log("[DASH7] Hoja historica omitida (solo DIAGNOSTICOS activos).");
    } else {
      Logger.log("[DASH7] Hoja historica no procesada (ya viene en DIAGNOSTICOS o no existe).");
    }

    asesoriasBuckets.forEach((bucket) => {
      let remaining = bucket.total || 0;
      if (remaining <= 0) return;

      const subtipoEntries = Array.from(bucket.subtipos.entries()).sort((a, b) => {
        const diff = (b[1] || 0) - (a[1] || 0);
        if (diff !== 0) return diff;
        return a[0].localeCompare(b[0]);
      });

      subtipoEntries.forEach(([subtipo, cnt]) => {
        if (remaining <= 0) return;
        const useCount = Math.min(cnt || 0, remaining);
        if (useCount <= 0) return;
        asesorias.push({
          ruc: bucket.ruc,
          razonSocial: bucket.razonSocial,
          tamano: bucket.tamano,
          sector: bucket.sector,
          tipoDiagnostico: "LEGAL",
          subtipoDiagnostico: subtipo,
          anio: bucket.anio,
          trimestre: bucket.trimestre,
          fuente: bucket.fuente,
          cantidad: useCount,
        });
        remaining -= useCount;
      });

      if (remaining > 0) {
        asesorias.push({
          ruc: bucket.ruc,
          razonSocial: bucket.razonSocial,
          tamano: bucket.tamano,
          sector: bucket.sector,
          tipoDiagnostico: "LEGAL",
          subtipoDiagnostico: "OTROS",
          anio: bucket.anio,
          trimestre: bucket.trimestre,
          fuente: bucket.fuente,
          cantidad: remaining,
        });
      }
    });

    const totalEventos = asesorias.reduce((acc, d) => acc + (d.cantidad || 1), 0);
    Logger.log("[DASH7] ===== RESUMEN =====");
    Logger.log(`Registros procesados: ${registrosProcesados}`);
    Logger.log(`Filtrados 'No': ${filtNo}`);
    Logger.log(`Filtrados 'ninguno': ${filtNinguno}`);
    Logger.log(`No encontrados en base: ${noFound}`);
    Logger.log(`Sin fecha (recuperados): ${sinFechaRecuperados}`);
    Logger.log(`Filas consolidadas en bucket (empresa|anio): ${bucketRowsMerged}`);
    Logger.log(`Total asesorias (eventos): ${totalEventos} | buckets: ${asesoriasBuckets.size} | filas salida: ${asesorias.length}`);
    Logger.log(`[DASH7 DEBUG] Filas utiles DIAGNOSTICOS: ${debugUsefulRowsDjango} | HIST: ${debugUsefulRowsHist}`);
    const tamanosDebug = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];
    tamanosDebug.forEach((t) => {
      const setBase = empresasBaseAsePorTam.get(t);
      const setRaw = empresasRawAsePorTam.get(t);
      Logger.log(`[DASH7 DEBUG] Empresas utiles (baseKey) - ${t}: ${setBase ? setBase.size : 0} | rawKey: ${setRaw ? setRaw.size : 0}`);
    });
    const totalEmpBase = Array.from(empresasBaseAsePorTam.values()).reduce((acc, s) => acc + s.size, 0);
    const totalEmpRaw = Array.from(empresasRawAsePorTam.values()).reduce((acc, s) => acc + s.size, 0);
    Logger.log(`[DASH7 DEBUG] Empresas utiles totales (baseKey): ${totalEmpBase} | (rawKey): ${totalEmpRaw}`);

    const yearsList = Array.from(years).sort((a, b) => yearRank(b) - yearRank(a));
    Logger.log(`Anios detectados: ${yearsList.join(", ")}`);

    generateAsesoriasResumenTable(
      baseMap,
      asesorias,
      yearsList,
      empresasBaseAsePorTam,
      empresasRawAsePorTam,
      useLegalSheetsOnly,
      empresasLegalDenomPorTam
    );
    generateAsesoriasSubtipoTable(asesorias, yearsList);
    generateAsesoriasPorEmpresaTable(asesorias, yearsList);
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

  function generateAsesoriasResumenTable(baseMap, asesorias, years, baseAseMap, rawAseMap, preferRawDenom = false, legalDenomMap = null) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "TOTAL_ASESORIAS", "EMPRESAS_CON_ASE", "EMPRESAS_SIN_ASE", "EMPRESAS_TOTALES"];
    const empresasPorTam = new Map();
    // Si preferRawDenom (modo LEGAL), usar el mapa deduplicado (baseAseMap) y caer a rawAseMap solo si falta.
    let sourceMap = null;
    if (preferRawDenom) {
      if (legalDenomMap && legalDenomMap.size) {
        sourceMap = legalDenomMap;
      } else if (baseAseMap && baseAseMap.size) {
        sourceMap = baseAseMap;
      } else if (rawAseMap && rawAseMap.size) {
        sourceMap = rawAseMap;
      }
    } else {
      sourceMap = baseMap && baseMap.size ? null : rawAseMap && rawAseMap.size ? rawAseMap : baseAseMap && baseAseMap.size ? baseAseMap : null;
    }
    if (!sourceMap) {
      baseMap.forEach((info) => {
        const tam = info.tamano || "SIN_TAMANO";
        if (!empresasPorTam.has(tam)) empresasPorTam.set(tam, new Set());
        empresasPorTam.get(tam).add(info.baseKey || info.ruc || info.razonSocial);
      });
    } else {
      sourceMap.forEach((set, tam) => {
        if (!empresasPorTam.has(tam)) empresasPorTam.set(tam, new Set());
        set.forEach((val) => empresasPorTam.get(tam).add(val));
      });
    }
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = asesorias.filter((d) => d.anio === anio && d.tamano === tam);
        const totalDiag = diags.reduce((acc, d) => acc + (d.cantidad || 1), 0);
        const empresasSet = new Set(diags.map((d) => d.ruc || d.razonSocial));
        const totalEmp = empresasPorTam.get(tam)?.size || 0;
        rows.push([anio, tam, totalDiag, empresasSet.size, Math.max(0, totalEmp - empresasSet.size), totalEmp]);
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

  function generateAsesoriasSubtipoTable(asesorias, years) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "SUBTIPO", "CANTIDAD", "PCT"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = asesorias.filter((d) => d.anio === anio && d.tamano === tam);
        const total = diags.reduce((acc, d) => acc + (d.cantidad || 1), 0);
        const porSubtipo = new Map();
        diags.forEach((d) => {
          const key = d.subtipoDiagnostico || "SIN SUBTIPO";
          porSubtipo.set(key, (porSubtipo.get(key) || 0) + (d.cantidad || 1));
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

  function generateAsesoriasPorEmpresaTable(asesorias, years) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "RUC", "RAZON_SOCIAL", "SECTOR", "TOTAL_ASESORIAS", "SUBTIPOS_TOMADOS", "RANK"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = asesorias.filter((d) => d.anio === anio && d.tamano === tam);
        const porEmp = new Map();
        diags.forEach((d) => {
          const key = d.ruc || d.razonSocial;
          if (!porEmp.has(key)) {
            porEmp.set(key, { ruc: d.ruc, razonSocial: d.razonSocial, sector: d.sector, count: 0, subtipos: new Set() });
          }
          const item = porEmp.get(key);
          item.count += d.cantidad || 1;
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
    if (ss.getSheetByName(preferred)) {
      Logger.log(`[DASH7] Hoja encontrada (exacta): ${preferred}`);
      return preferred;
    }

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
        info.headerIndex["LEGAL_1"] !== undefined ||
        info.headerIndex["LEGAL1"] !== undefined ||
        info.headerIndex["LEGAL_2"] !== undefined ||
        info.headerIndex["ASESORIA_LEGAL"] !== undefined ||
        info.headerIndex["ASESORIAS_LEGALES"] !== undefined ||
        info.headerIndex["SERVICIO_LEGAL"] !== undefined;
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
        Logger.log(`[DASH7] Hoja historica encontrada (preferida): ${name} (${rowCount} filas)`);
        return;
      }

      if (norm.indexOf("DIAGNOST") !== -1 && hasLegalCols && rowCount >= 5) {
        const score = rowCount;
        if (score > bestScore) {
          best = name;
          bestScore = score;
          Logger.log(`[DASH7] Candidato de hoja historica: ${name} (score: ${score})`);
        }
      }
    });

    if (best) {
      Logger.log(`[DASH7] Hoja historica seleccionada: ${best} (score: ${bestScore})`);
      return best;
    }

    Logger.log("[DASH7] Advertencia: No se encontro hoja historica (DIAGNOSTICO)");
    return null;
  }

  function refreshDashboardAsesorias() {
    buildAsesoriasLegalesData();
  }

  function onOpenDash7() {
    SpreadsheetApp.getUi().createMenu("DASH7").addItem("Generar asesorias", "refreshDashboardAsesorias").addToUi();
  }

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
      ss.toast("Actualizando dashboard.");
      try {
        refreshDashboardAsesorias();
        ss.toast("Dashboard listo.");
        sheet.getRange("Q54").setValue("Ultima actualizacion: " + new Date());
      } finally {
        range.setValue(false);
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

function onEditDash7Wrapper(e) {
  return onEditDash7(e);
}

function onOpenDash7Wrapper() {
  return onOpenDash7();
}

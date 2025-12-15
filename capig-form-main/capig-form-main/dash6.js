/**
 * Dashboard de Diagnosticos - DASH6
 * Integra historico (hoja DIAGNOSTICO) + registros Django.
 * Soporta alias multiples, datos sin fecha y fallbacks robustos.
 */
(function (global) {
  const SHEETS = {
    BASE: "SOCIOS",
    BASE_FALLBACK: "BASE DE DATOS",
    DIAG_DJANGO: "DIAGNOSTICOS",
    DIAG_DJANGO_FALLBACK: "DIAGNOSTICO_FINAL",
    DIAG_HIST: "DIAGNOSTICOS_HISTORICOS",
    OUT_RESUMEN: "PIVOT_DIAGNOSTICOS_RESUMEN_ANIO",
    OUT_TIPO: "PIVOT_DIAGNOSTICOS_TIPO_ANIO",
    OUT_EMPRESAS: "PIVOT_DIAGNOSTICOS_POR_EMPRESA",
    OUT_MASTER_SLICER: "DASH6_MAESTRA",
  };

  const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, SIN_TAMANO: 99 };
  const HIST_YEAR = "HISTORICO";
  const NO_DATE_YEAR = "SIN_FECHA";
  const DEFAULT_HIST_YEAR = "2025";
  const TYPE_COLUMNS = ["LEAN", "ESTRATEGIA", "LEGAL", "AMBIENTE", "RRHH"];
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
    if (isMarkerEmpty(clean)) return "SIN TIPO";
    if (compact === "NINGUNO" || compact === "NINGUNA" || compact === "NINGUN") return "NINGUNO";
    if (clean.indexOf("ESTRAT") !== -1) return "ESTRATEGIA";
    if (clean.indexOf("LEAN") !== -1) return "LEAN";
    if (clean.indexOf("AMBI") !== -1) return "AMBIENTE";
    if (clean.indexOf("LEGAL") !== -1) return "LEGAL";
    if (clean.indexOf("RRHH") !== -1 || clean.indexOf("RH") !== -1 || clean.indexOf("RECURSO") !== -1) return "RRHH";
    return clean;
  }

  function normalizeSubtipoDiagnostico(val) {
    const s = (val || "").toString().trim().toUpperCase();
    if (!s || isMarkerEmpty(s)) return "SIN SUBTIPO";
    const clean = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const normalized = clean.replace(/_/g, " ");
    if (normalized.indexOf("PEND") !== -1) return "PENDIENTE";
    if (normalized.indexOf("LABOR") !== -1) return "LABORAL";
    if (normalized.indexOf("PROPIEDAD") !== -1) return "PROPIEDAD INTELECTUAL";
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

  function findSheetByPartialName(partial, exclude) {
    const ss = SpreadsheetApp.getActive();
    const sheets = ss.getSheets();
    const target = normalizeLabel(partial);
    const excl = exclude ? normalizeLabel(exclude) : null;
    for (const s of sheets) {
      const norm = normalizeLabel(s.getName());
      if (norm.includes(target) && (!excl || !norm.includes(excl))) {
        Logger.log(`[DASH6] Hoja encontrada por parcial '${partial}': ${s.getName()}`);
        return s.getName();
      }
    }
    Logger.log(`[DASH6] No se encontro hoja por parcial '${partial}' (excluye '${exclude || ""}')`);
    return null;
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

  // Devuelve la primera tabla disponible segun una lista de nombres candidatos.
  function readTableFlexibleByCandidates(names = []) {
    const ss = SpreadsheetApp.getActive();
    for (const name of names) {
      if (!name) continue;
      const sh = ss.getSheetByName(name);
      if (sh) {
        const data = readTableFlexible(name);
        Logger.log(`[DASH6] Usando hoja ${name}: ${data.rows.length} filas`);
        return data;
      }
    }
    Logger.log(`[DASH6] No se encontraron hojas candidatas: ${names.join(", ")}`);
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

  function parseCountFlag(val) {
    if (val === null || val === undefined) return 0;
    const raw = val.toString().trim();
    if (!raw) return 0;
    const num = Number(raw);
    if (!isNaN(num)) return num > 0 ? 1 : 0; // cualquier numero > 0 vale como 1
    const upper = raw.toUpperCase();
    if (upper === "SI" || upper === "SÍ" || upper === "TRUE" || upper === "X") return 1;
    return 0;
  }

  function collectTiposFromRow(row, headerIndex, tipoRawFromColumn) {
    const tipos = new Map();
    function addTipo(tipo, count) {
      if (!tipo || tipo === "SIN TIPO" || tipo === "NINGUNO") return;
      tipos.set(tipo, (tipos.get(tipo) || 0) + count);
    }

    if (tipoRawFromColumn && tipoRawFromColumn !== "SIN TIPO" && tipoRawFromColumn !== "NINGUNO") {
      addTipo(tipoRawFromColumn, 1);
    }

    TYPE_COLUMNS.forEach((col) => {
      const count = parseCountFlag(getVal(row, headerIndex, [col]));
      if (count > 0) addTipo(normalizeTipoDiagnostico(col), count);
    });

    return Array.from(tipos.entries()).map(([tipo, count]) => ({ tipo, count }));
  }

  function isAdvisoryLegalSubtipo(subtipoRaw) {
    const ADVISORY_LEGAL_SUBTYPES = ["LABORAL", "PROPIEDAD INTELECTUAL", "SOCIETARIO", "CONTACTO", "OTROS"];
    return ADVISORY_LEGAL_SUBTYPES.includes(subtipoRaw);
  }

  // ---------- Core ----------
  function buildDiagnosticosData() {
    Logger.log("[DASH6] Iniciando generacion...");

    const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
    const diagDjango = readTableFlexibleByCandidates([SHEETS.DIAG_DJANGO, SHEETS.DIAG_DJANGO_FALLBACK]);
    const diagHist = readTableFlexibleByCandidates([SHEETS.DIAG_HIST, getSheetNameFlexible("DIAGNOSTICO")]);

    Logger.log(`[DASH6] Procesando hoja de diagnosticos unificada: ${diagDjango.rows.length} filas`);
    Logger.log(`[DASH6] Procesando hoja de diagnosticos historicos: ${diagHist.rows.length} filas`);

    // Si la hoja principal ya trae historico embebido (filas sin fecha), no procesar la hoja historica aparte
    const hasHistInsideDjango = diagDjango.rows.some(
      (row) => !getYear(getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]))
    );
    const shouldProcessHist = !hasHistInsideDjango && diagHist.rows.length > 0;
    Logger.log(`[DASH6] Historico embebido en DIAGNOSTICOS: ${hasHistInsideDjango ? "SI" : "NO"} | Procesar hoja historica aparte: ${shouldProcessHist ? "SI" : "NO"}`);

    const baseMap = new Map(); // key ruc or fallback name -> info
    const aliasToRuc = new Map(); // nombre normalizado -> clave primaria (ruc/altId/razon)
    const altIdToBase = new Map(); // altId normalizado -> data
    const missingEmpresas = new Map(); // normalizado -> placeholder (no impacta totales)
    const empresasBaseDiagPorTam = new Map(); // tamano -> set de empresas provenientes de DIAGNOSTICOS
    const empresasRawDiagPorTam = new Map(); // tamano -> set de rawKeys (sin merge por alias)
    const empresaTamAsignado = new Map(); // baseKey -> tamano asignado unico

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
    Logger.log(`[DASH6] Empresas base: ${baseMap.size}`);

    applyManualAliases(aliasToRuc, baseMap);

    const diagnosticos = [];
    const years = new Set();
    let registrosProcesados = 0;
    let filtNo = 0;
    let filtNinguno = 0;
    let noFound = 0;
    let sinFechaRecuperados = 0;
    let duplicadosAgregados = 0;
    let debugUsefulRowsDjango = 0;
    let debugUsefulRowsHist = 0;

    // Determinar ano de fallback para Django y datos historicos
    // Default a 2025 para historico sin fecha y al maximo ano detectado (o 2025) para Django
    let djangoFallbackYear = "";
    diagDjango.rows.forEach((row) => {
      const y = getYear(getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]));
      if (y && (!djangoFallbackYear || y > djangoFallbackYear)) djangoFallbackYear = y;
    });
    if (!djangoFallbackYear) djangoFallbackYear = DEFAULT_HIST_YEAR;
    const histFallbackYear = DEFAULT_HIST_YEAR; // Historico sin fecha se ancla a 2025 para no mezclar con pruebas nuevas


    Logger.log(`[DASH6] Ano de fallback Django: ${djangoFallbackYear}`);
    Logger.log(`[DASH6] Ano de fallback Historico: ${histFallbackYear}`);


    function yearOrDefault(fechaRaw, fuente, fallbackYear) {
      const y = getYear(fechaRaw);
      if (y) return y;
      if (fuente === "HIST") return fallbackYear || HIST_YEAR;
      if (fuente === "DJANGO") return fallbackYear || NO_DATE_YEAR;
      return fallbackYear || NO_DATE_YEAR;
    }


    function resolveEmpresa(name, rowForAltId, rowHeaderIndex) {
      const norm = normalizeName(name);
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

    // Mapa para agregar diagnosticos por RUC|ANO|TIPO
    const diagMap = new Map();

    function addDiagnosticoAggregated(empresa, tipo, count, fechaRaw, fuente, fallbackYear, subtipoRaw) {
      const anio = yearOrDefault(fechaRaw, fuente, fallbackYear);
      const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
      const key = `${entityKey}|${anio}|${tipo}`;
      let bucket = diagMap.get(key);
      if (!bucket) {
        bucket = {
          ruc: empresa.ruc,
          razonSocial: empresa.razonSocial,
          tamano: empresa.tamano,
          sector: empresa.sector,
          tipoDiagnostico: tipo,
          subtipos: new Set(),
          anio,
          trimestre: getTrimestre(fechaRaw),
          fuente,
          cantidad: 0,
        };
        diagMap.set(key, bucket);
        years.add(anio);
      } else {
        duplicadosAgregados += count;
      }
      bucket.cantidad += count;
      if (subtipoRaw && subtipoRaw !== "SIN SUBTIPO" && subtipoRaw !== "PENDIENTE") {
        bucket.subtipos.add(subtipoRaw);
      }
    }

    // DIAGNOSTICOS - procesar TODO (historico sin fecha + nuevo con fecha)
    diagDjango.rows.forEach((row) => {
      const razonRaw = getVal(row, diagDjango.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "Razon Social"]);
      const empresa = resolveEmpresa(razonRaw, row, diagDjango.headerIndex);
      const seDiagRaw = (
        getVal(row, diagDjango.headerIndex, [
          "SE_DIAGNOSTICO",
          "SE DIAGNOSTICO",
          "SE_DIAGNOSTICO_",
          "SE DIAGNOSTICO_",
          "TOTAL_DIAGNOSTICO",
          "TOTAL DIAGNOSTICO",
          "TOTAL_DIAGNOSTICOS",
          "TOTAL DIAGNOSTICOS",
          "DIAGNOSTICO",
        ]) || ""
      )
        .toString()
        .trim()
        .toUpperCase();
      const tipoRaw = normalizeTipoDiagnostico(getVal(row, diagDjango.headerIndex, ["TIPO_DE_DIAGNOSTICO", "TIPO", "TIPO DIAGNOSTICO", "TIPO DE DIAGNOSTICO"]));
      if (tipoRaw === "NINGUNO") {
        filtNinguno++;
      }
      const subtipoRaw = normalizeSubtipoDiagnostico(
        getVal(row, diagDjango.headerIndex, ["SUBTIPO_DIAGNOSTICO", "SUBTIPO_DE_DIAGNOSTICO", "SUBTIPO", "SUBTIPO DIAGNOSTICO", "SUBTIPO DE DIAGNOSTICO", "OTROS_SUBTIPO", "OTROS SUBTIPO"])
      );
      const tiposDet = collectTiposFromRow(row, diagDjango.headerIndex, tipoRaw);
      const hasUsefulData = !!seDiagRaw || tiposDet.length > 0;

      const fechaRaw = getVal(row, diagDjango.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]);
      const yearFromDate = getYear(fechaRaw);
      const hasDate = !!yearFromDate;

      if (!empresa) {
        Logger.log(`[DASH6] Empresa no encontrada en base: ${razonRaw}`);
        noFound++;
        return;
      }

      if (!hasUsefulData) return;

      // Contabilizar empresa (denominador) incluso si SE_DIAGNOSTICO es "NO"
      if (empresa && !empresa.isPlaceholder) {
        const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
        const tamAsignado = empresaTamAsignado.get(entityKey) || empresa.tamano || "SIN_TAMANO";
        if (!empresaTamAsignado.has(entityKey)) empresaTamAsignado.set(entityKey, tamAsignado);
        const tamKey = tamAsignado || "SIN_TAMANO";
        if (!empresasBaseDiagPorTam.has(tamKey)) empresasBaseDiagPorTam.set(tamKey, new Set());
        empresasBaseDiagPorTam.get(tamKey).add(entityKey);
        // Denominador "raw" sin merge por alias (para coincidir con conteo esperado)
        const rawRuc = padRuc13(getVal(row, diagDjango.headerIndex, ["RUC"]));
        const rawAlt = normalizeKey(getVal(row, diagDjango.headerIndex, ALT_ID_ALIASES));
        const rawName = normalizeName(razonRaw);
        const rawKey = rawRuc || rawAlt || rawName;
        if (rawKey) {
          if (!empresasRawDiagPorTam.has(tamKey)) empresasRawDiagPorTam.set(tamKey, new Set());
          empresasRawDiagPorTam.get(tamKey).add(rawKey);
        }
        debugUsefulRowsDjango++;
      }

      if (seDiagRaw && shouldSkipDiagnosticoFlag(seDiagRaw)) {
        filtNo++;
        return;
      }

      if (!hasDate) sinFechaRecuperados++;

      tiposDet.forEach(({ tipo, count }) => {
        if (tipo === "LEGAL" && isAdvisoryLegalSubtipo(subtipoRaw)) {
          const razonDebug = getVal(row, diagDjango.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "Razon Social"]);
          Logger.log(`[DASH6 FILTRO] EXCLUIDO Asesoria Legal: ${razonDebug} | Subtipo: ${subtipoRaw}`);
          return;
        }
        addDiagnosticoAggregated(empresa, tipo, count, fechaRaw, hasDate ? "DJANGO" : "HIST", hasDate ? djangoFallbackYear : histFallbackYear, subtipoRaw);
        registrosProcesados += count;
      });
    });

    // HISTORICO - procesar hoja separada (sin fecha, usar fallback) solo si no viene embebido en DIAGNOSTICOS
    if (shouldProcessHist) {
      diagHist.rows.forEach((row) => {
        const seDiagRaw = (
          getVal(row, diagHist.headerIndex, [
            "DIAGNOSTICO",
            "TOTAL_DIAGNOSTICO",
            "TOTAL DIAGNOSTICO",
            "TOTAL_DIAGNOSTICOS",
          "TOTAL DIAGNOSTICOS",
        ]) || ""
        )
          .toString()
          .trim()
          .toUpperCase();

        const tiposDet = collectTiposFromRow(row, diagHist.headerIndex, "");
        const hasUsefulData = !!seDiagRaw || tiposDet.length > 0;
        if (!hasUsefulData) return;

        const razonRaw = getVal(row, diagHist.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "NOMBRE", "RAZON_SOCIAL_"]);
        const empresa = resolveEmpresa(razonRaw, row, diagHist.headerIndex);
      if (!empresa) {
        Logger.log(`[DASH6] Empresa no encontrada en base (hist): ${razonRaw}`);
        noFound++;
        return;
      }

      // Contabilizar empresa base proveniente de historico para el denominador (incluye NO)
      if (!empresa.isPlaceholder) {
        const entityKey = empresa.baseKey || empresa.ruc || empresa.razonSocial;
        const tamAsignado = empresaTamAsignado.get(entityKey) || empresa.tamano || "SIN_TAMANO";
        if (!empresaTamAsignado.has(entityKey)) empresaTamAsignado.set(entityKey, tamAsignado);
        const tamKey = tamAsignado || "SIN_TAMANO";
        if (!empresasBaseDiagPorTam.has(tamKey)) empresasBaseDiagPorTam.set(tamKey, new Set());
          empresasBaseDiagPorTam.get(tamKey).add(entityKey);
          const rawRuc = padRuc13(getVal(row, diagHist.headerIndex, ["RUC"]));
          const rawAlt = normalizeKey(getVal(row, diagHist.headerIndex, ALT_ID_ALIASES));
          const rawName = normalizeName(razonRaw);
          const rawKey = rawRuc || rawAlt || rawName;
          if (rawKey) {
            if (!empresasRawDiagPorTam.has(tamKey)) empresasRawDiagPorTam.set(tamKey, new Set());
            empresasRawDiagPorTam.get(tamKey).add(rawKey);
        }
        debugUsefulRowsHist++;
      }

      if (seDiagRaw && shouldSkipDiagnosticoFlag(seDiagRaw)) {
        filtNo++;
        return;
      }

      const fechaRaw = getVal(row, diagHist.headerIndex, ["FECHA", "FECHA_DIAGNOSTICO", "FECHA DE DIAGNOSTICO"]);
      tiposDet.forEach(({ tipo, count }) => {
        addDiagnosticoAggregated(empresa, tipo, count, fechaRaw, "HIST", histFallbackYear, "SIN SUBTIPO");
        registrosProcesados += count;
        sinFechaRecuperados += count;
        });
      });
    } else {
      Logger.log("[DASH6] Hoja historica no procesada (ya viene en DIAGNOSTICOS o no existe).");
    }

    // Convertir mapa agregado a lista de diagnosticos
    diagMap.forEach((val) => {
      const subtiposStr = Array.from(val.subtipos).join(", ");
      diagnosticos.push({
        ruc: val.ruc,
        razonSocial: val.razonSocial,
        tamano: val.tamano,
        sector: val.sector,
        tipoDiagnostico: val.tipoDiagnostico,
        subtipoDiagnostico: subtiposStr || "SIN SUBTIPO", // Si estaba vacio o solo PENDIENTE
        anio: val.anio,
        trimestre: val.trimestre,
        fuente: val.fuente,
        cantidad: val.cantidad,
      });
    });



    const totalEventos = diagnosticos.reduce((acc, d) => acc + (d.cantidad || 1), 0);
    Logger.log("[DASH6] ===== RESUMEN =====");
    Logger.log(`Registros procesados: ${registrosProcesados}`);
    Logger.log(`Filtrados 'No': ${filtNo}`);
    Logger.log(`Filtrados 'ninguno': ${filtNinguno}`);
    Logger.log(`No encontrados en base: ${noFound}`);
    Logger.log(`Sin fecha (recuperados): ${sinFechaRecuperados}`);
    Logger.log(`Duplicados agregados: ${duplicadosAgregados}`);
    Logger.log(`Total diagnosticos (eventos): ${totalEventos} | buckets: ${diagnosticos.length}`);
    Logger.log(`[DASH6 DEBUG] Filas utiles DIAGNOSTICOS: ${debugUsefulRowsDjango} | HIST: ${debugUsefulRowsHist}`);
    const tamanosDebug = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];
    tamanosDebug.forEach((t) => {
      const setBase = empresasBaseDiagPorTam.get(t);
      const setRaw = empresasRawDiagPorTam.get(t);
      Logger.log(
        `[DASH6 DEBUG] Empresas utiles (baseKey) - ${t}: ${setBase ? setBase.size : 0} | rawKey: ${setRaw ? setRaw.size : 0}`
      );
    });
    const totalEmpBase = Array.from(empresasBaseDiagPorTam.values()).reduce((acc, s) => acc + s.size, 0);
    const totalEmpRaw = Array.from(empresasRawDiagPorTam.values()).reduce((acc, s) => acc + s.size, 0);
    Logger.log(`[DASH6 DEBUG] Empresas utiles totales (baseKey): ${totalEmpBase} | (rawKey): ${totalEmpRaw}`);

    const yearsList = Array.from(years).sort((a, b) => yearRank(b) - yearRank(a));
    Logger.log(`Anos detectados: ${yearsList.join(", ")}`);

    generateDiagnosticosResumenTable(baseMap, diagnosticos, yearsList, empresasBaseDiagPorTam, empresasRawDiagPorTam);
    generateDiagnosticosTipoTable(diagnosticos, yearsList);
    generateDiagnosticosPorEmpresaTable(diagnosticos, yearsList);
    generateSlicerMasterSheet();

    Logger.log("[DASH6] Completado");
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

  function yearRank(val) {
    const num = parseInt(val, 10);
    if (!isNaN(num)) return num;
    if (val === HIST_YEAR) return -1;
    if (val === NO_DATE_YEAR) return -2;
    return -3;
  }

  function generateDiagnosticosResumenTable(baseMap, diagnosticos, years, baseDiagMap, rawDiagMap) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "TOTAL_DIAGNOSTICOS", "EMPRESAS_CON_DIAG", "EMPRESAS_SIN_DIAG", "EMPRESAS_TOTALES"];
    const empresasPorTam = new Map();
    const sourceMap = rawDiagMap && rawDiagMap.size ? rawDiagMap : baseDiagMap && baseDiagMap.size ? baseDiagMap : null;
    if (sourceMap) {
      sourceMap.forEach((set, tam) => {
        if (!empresasPorTam.has(tam)) empresasPorTam.set(tam, new Set());
        set.forEach((val) => empresasPorTam.get(tam).add(val));
      });
    } else {
      baseMap.forEach((info) => {
        const tam = info.tamano || "SIN_TAMANO";
        if (!empresasPorTam.has(tam)) empresasPorTam.set(tam, new Set());
        empresasPorTam.get(tam).add(info.baseKey || info.ruc || info.razonSocial);
      });
    }
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
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
    Logger.log(`[DASH6] ${SHEETS.OUT_RESUMEN}: ${rows.length} filas`);
  }

  function generateDiagnosticosTipoTable(diagnosticos, years) {
    const rows = [];
    // Se ignoran los subtipos: se agrupa solo por tipo de diagnostico.
    const header = ["ANIO", "TAMANO", "TIPO_DIAGNOSTICO", "CANTIDAD", "PCT"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
        const total = diags.reduce((acc, d) => acc + (d.cantidad || 1), 0);
        const porTipo = new Map();
        diags.forEach((d) => {
          const key = d.tipoDiagnostico || "SIN TIPO";
          porTipo.set(key, (porTipo.get(key) || 0) + (d.cantidad || 1));
        });
        porTipo.forEach((count, key) => {
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

    const sheet = getSheetOrCreate(SHEETS.OUT_TIPO);
    sheet.clear();
    const out = rows.length ? [header, ...rows] : [header];
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.autoResizeColumns(1, header.length);
    Logger.log(`[DASH6] ${SHEETS.OUT_TIPO}: ${rows.length} filas`);
  }

  function generateDiagnosticosPorEmpresaTable(diagnosticos, years) {
    const rows = [];
    const header = ["ANIO", "TAMANO", "RUC", "RAZON_SOCIAL", "SECTOR", "TOTAL_DIAGNOSTICOS", "TIPOS_TOMADOS", "RANK"];
    const tamanos = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE", "SIN_TAMANO"];

    years.forEach((anio) => {
      tamanos.forEach((tam) => {
        const diags = diagnosticos.filter((d) => d.anio === anio && d.tamano === tam);
        const porEmp = new Map();
        diags.forEach((d) => {
          const key = d.ruc || d.razonSocial;
          if (!porEmp.has(key)) {
            porEmp.set(key, { ruc: d.ruc, razonSocial: d.razonSocial, sector: d.sector, count: 0, tipos: new Set() });
          }
          const item = porEmp.get(key);
          item.count += d.cantidad || 1;
          item.tipos.add(d.tipoDiagnostico);
        });
        const lista = Array.from(porEmp.values()).sort((a, b) => b.count - a.count);
        lista.forEach((item, idx) => {
          rows.push([anio, tam, item.ruc, item.razonSocial, item.sector, item.count, Array.from(item.tipos).join(", "), idx + 1]);
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
    Logger.log(`[DASH6] ${SHEETS.OUT_EMPRESAS}: ${rows.length} filas`);
  }

  function generateSlicerMasterSheet() {
    const slicerSources = [
      { name: SHEETS.OUT_RESUMEN, cols: 6 },
      { name: SHEETS.OUT_TIPO, cols: 5 },
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

      const header = values[0].slice(0, cols).map((h, idx) => {
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
      });
      moneyCols.forEach((col) => {
        sheet.getRange(2, col, mergedRows.length, 1).setNumberFormat('"$"#,##0.00');
      });
    }

    Logger.log(`[DASH6] ${SHEETS.OUT_MASTER_SLICER}: ${mergedRows.length} filas (merge para slicers)`);
  }

  function getSheetNameFlexible(preferred) {
    const ss = SpreadsheetApp.getActive();
    if (ss.getSheetByName(preferred)) {
      Logger.log(`[DASH6] Hoja historica encontrada (exacta): ${preferred}`);
      return preferred;
    }

    const sheets = ss.getSheets();
    const preferredNorm = normalizeLabel(preferred);
    let best = null;
    let bestScore = -1;

    // EXCLUSION PATTERNS - Explicitly exclude advisory/asesorias related sheets
    const EXCLUDED_PATTERNS = [
      "ASESORIA",
      "ASESORIAS",
      "SERVICIO_LEGAL",
      "SERVICIOS_LEGALES",
      "LEGAL_FINAL",
      "ASESORIAS_LEGALES",
      "CONSULTORIA",
      "ADVISORY"
    ];

    function scoreSheet(sheetName) {
      const info = readTableFlexible(sheetName);
      const rowCount = info.rows.length;
      if (rowCount === 0) return null;
      const hasDiagCols =
        info.headerIndex["LEAN"] !== undefined ||
        info.headerIndex["ESTRATEGIA"] !== undefined ||
        info.headerIndex["LEGAL"] !== undefined ||
        info.headerIndex["AMBIENTE"] !== undefined ||
        info.headerIndex["RRHH"] !== undefined ||
        info.headerIndex["DIAGNOSTICO"] !== undefined ||
        info.headerIndex["DIAGNOSTICO_"] !== undefined;

      return { rowCount, hasDiagCols };
    }

    sheets.forEach((s) => {
      const name = s.getName();
      const norm = normalizeLabel(name);

      // Skip FINAL and PIVOT sheets
      if (norm.indexOf("FINAL") !== -1) return;
      if (norm.indexOf("PIVOT") !== -1) return;

      // CRITICAL FIX: Explicitly exclude any sheets containing advisory-related keywords
      const isExcluded = EXCLUDED_PATTERNS.some(pattern => norm.indexOf(pattern) !== -1);
      if (isExcluded) {
        Logger.log(`[DASH6] Hoja EXCLUIDA (asesoria detectada): ${name}`);
        return;
      }

      const meta = scoreSheet(name);
      if (!meta) return;
      const { rowCount, hasDiagCols } = meta;

      // Prioridad 1: nombre exacto preferido "DIAGNOSTICO" 
      if (norm === preferredNorm && rowCount > 0) {
        Logger.log(`[DASH6] Hoja historica encontrada (preferida exacta): ${name} (${rowCount} filas)`);
        best = name;
        bestScore = rowCount + 10000; // muy alto sesgo para exactos
        return;
      }

      // Prioridad 2: nombre "DIAGNOSTICOS" (plural)
      if (norm === "DIAGNOSTICOS" && rowCount > 0) {
        Logger.log(`[DASH6] Hoja historica encontrada (diagnosticos plural): ${name} (${rowCount} filas)`);
        best = name;
        bestScore = rowCount + 5000; // alto sesgo
        return;
      }

      // Prioridad 3: contiene DIAGNOST, tiene columnas diagnostico, suficientes filas
      if (norm.indexOf("DIAGNOST") !== -1 && hasDiagCols && rowCount >= 5) {
        const score = rowCount;
        if (score > bestScore) {
          best = name;
          bestScore = score;
          Logger.log(`[DASH6] Candidato de hoja historica: ${name} (score: ${score})`);
        }
      }
    });

    if (best) {
      Logger.log(`[DASH6] Hoja historica SELECCIONADA FINAL: ${best} (score: ${bestScore})`);
      return best;
    }

    Logger.log("[DASH6] ADVERTENCIA: No se encontro hoja historica (DIAGNOSTICO)");
    return null;
  }

  function refreshDashboardDiagnosticos() {
    buildDiagnosticosData();
  }

  function onOpenDash6() {
    SpreadsheetApp.getUi().createMenu("DASH6").addItem("Generar diagnosticos", "refreshDashboardDiagnosticos").addToUi();
  }

  // Trigger manual via checkbox (hoja REPORTE_2, celda Q1)
  const DASH6_TRIGGER_SHEET = "REPORTE_2";
  const DASH6_TRIGGER_CELL = "Q1";

  function onEditDash6(e) {
    const range = e.range;
    if (!range) return;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== DASH6_TRIGGER_SHEET) return;
    if (range.getA1Notation() !== DASH6_TRIGGER_CELL) return;

    const val = range.getValue();
    if (val === true) {
      const ss = SpreadsheetApp.getActive();
      ss.toast("Actualizando dashboard…");
      try {
        refreshDashboardDiagnosticos();
        ss.toast("Dashboard listo.");
        sheet.getRange("Q2").setValue("Ultima actualizacion: " + new Date());
      } finally {
        range.setValue(false); // deja la casilla lista para el siguiente clic
      }
    }
  }

  global.refreshDashboardDiagnosticos = refreshDashboardDiagnosticos;
  global.onOpenDash6 = onOpenDash6;
  global.onEditDash6 = onEditDash6;
})(this);

function refreshDashboardDiagnosticosWrapper() {
  return refreshDashboardDiagnosticos();
}

// Wrapper para que el trigger instalable detecte la función en el selector
function onEditDash6Wrapper(e) {
  return onEditDash6(e);
}

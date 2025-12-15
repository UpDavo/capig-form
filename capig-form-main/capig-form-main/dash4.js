
/**
 * Dashboard de Desempeno de Empresas Socias - DASH4
 * Genera tablas pre-agregadas para analisis de ventas, sectores, estados y tendencias.
 * Disenado para actualizacion automatica con datos historicos + Django.
 * Ejecutar refreshDashboardDesempeno() para regenerar los pre-agregados.
 */
(function (global) {
    const SHEETS = {
        BASE: "SOCIOS",
        BASE_FALLBACK: "BASE DE DATOS",
        VENTAS: "VENTAS_SOCIO",
        VENTAS_FALLBACK: "VENTAS_AFILIADOS",
        ESTADO: "ESTADO_SOCIO",
        ESTADO_FALLBACK: "ESTADO_AFILIADOS",
        SECTOR: "SECTOR",
        OUT_RESUMEN: "PIVOT_VENTAS_RESUMEN_ANIO",
        OUT_SECTOR: "PIVOT_VENTAS_SECTOR_ANIO",
        OUT_TOP: "PIVOT_TOP_EMPRESAS_VENTAS",
        OUT_ESTADO: "PIVOT_ESTADO_EMPRESAS",
        OUT_SEMAFORO: "PIVOT_SEMAFORO_VENTAS",
        OUT_TOP_SECTORES_REL: "PIVOT_TOP_SECTORES_RELEVANTES",
        OUT_TOP_EMPRESAS_REL: "PIVOT_TOP_EMPRESAS_RELEVANTES",
        OUT_TOP_EMPRESAS_ANIO: "PIVOT_TOP_EMPRESAS_ANIO",
        OUT_MASTER: "DASH4_MASTER",
        OUT_MASTER_ALL: "DASH4_MASTER_ALL",
        OUT_MASTER_SLICER: "DASH4_MAESTRA"
    };

    const TAMANO_BY_CODE = { 1: "MICRO", 2: "PEQUENA", 3: "MEDIANA", 4: "GRANDE" };
    const TAMANO_ORDER = { MICRO: 1, PEQUENA: 2, MEDIANA: 3, GRANDE: 4, GLOBAL: 0 };

    // -------------- Utils (Reutilizados de dash3) --------------
    function normalizeLabel(label) {
        let txt = (label || "").toString().trim().toUpperCase();
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

    function detectYearColumns(headerIndex) {
        const yearColumns = [];
        const yearMap = new Map();
        const EXCLUDED_PATTERNS = [
            /COLABORA/,
            /EMPLEA/,
            /TRABAJADOR/,
            /RUC/,
            /RAZON/,
            /NOMBRE/,
            /CIUDAD/,
            /DIRECCION/,
            /TELEFONO/,
            /EMAIL/,
            /REPRESENT/,
            /CARGO/,
            /GENERO/,
            /SECTOR/,
            /TAMANO/,
            /ESTADO/,
            /AFILIACION/,
            /FECHA/
        ];

        for (const colName in headerIndex) {
            if (EXCLUDED_PATTERNS.some((pattern) => pattern.test(colName))) continue;

            const matchT = colName.match(/^T_?(\d{4})$/);
            if (matchT) {
                const year = parseInt(matchT[1], 10);
                const colIndex = headerIndex[colName];
                if (!yearMap.has(year)) {
                    yearMap.set(year, { year, colName, colIndex, priority: 1 });
                }
                continue;
            }

            const matchYear = colName.match(/^(\d{4})$/);
            if (matchYear) {
                const year = parseInt(matchYear[1], 10);
                if (year >= 1900 && year <= 2100) {
                    const colIndex = headerIndex[colName];
                    if (!yearMap.has(year)) {
                        yearMap.set(year, { year, colName, colIndex, priority: 2 });
                    }
                }
            }
        }

        yearMap.forEach((col) => yearColumns.push(col));
        yearColumns.sort((a, b) => b.year - a.year);
        return yearColumns;
    }
    // -------------- Utils Especificos DASH4 --------------
    function normalizeSector(val) {
        let s = (val || "").toString().trim().toUpperCase();
        s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        if (!s) return "SIN CLASIFICAR";
        if (s.indexOf("QUIM") !== -1) return "QUIMICO";
        if (s.indexOf("METAL") !== -1) return "METALMECANICO";
        if (s.indexOf("ALIMENT") !== -1) return "ALIMENTOS";
        if (s.indexOf("AGRIC") !== -1 || s.indexOf("AGROP") !== -1) return "AGRICOLA";
        if (s.indexOf("MAQUIN") !== -1) return "MAQUINARIAS";
        if (s.indexOf("CONST") !== -1) return "CONSTRUCCION";
        if (s.indexOf("TEXT") !== -1) return "TEXTIL";
        if (s.indexOf("COME") !== -1 || s.indexOf("RETAIL") !== -1) return "COMERCIO";
        return s;
    }

    function normalizeEstado(val) {
        const e = (val || "").toString().trim().toUpperCase();
        if (!e) return "DESCONOCIDO";
        const clean = e.replace(/\s+/g, " ");

        const esPagado =
            clean === "ACTIVO" ||
            clean === "ACTIVA" ||
            clean === "PAGADO" ||
            clean === "PAGADA" ||
            (clean.indexOf("PAG") !== -1 && clean.indexOf("NO") === -1 && clean.indexOf("PEND") === -1);
        if (esPagado) return "PAGADO";

        if (
            clean.indexOf("NO") !== -1 ||
            clean.indexOf("PEND") !== -1 ||
            clean.indexOf("INACTIV") !== -1 ||
            clean.indexOf("INACTIVO") !== -1 ||
            clean === "INACTIVO" ||
            clean.indexOf("SUSPEND") !== -1 ||
            clean.indexOf("RETIR") !== -1 ||
            clean.indexOf("BAJA") !== -1 ||
            clean.indexOf("MORA") !== -1
        ) {
            return "NO PAGADO";
        }

        return "DESCONOCIDO";
    }

    function parseMonto(val) {
        if (val === null || val === undefined || val === "") return 0;
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

        str = str.replace(/[^0-9\.-]/g, "");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    function parseEmpleados(val) {
        if (!val) return 0;
        const e = parseInt(val, 10);
        return isNaN(e) ? 0 : e;
    }

    function formatNumberColumn(sheet, startRow, startCol, numRows, formatPattern) {
        if (numRows <= 0) return;
        sheet.getRange(startRow, startCol, numRows, 1).setNumberFormat(formatPattern);
    }

    const MILLION_DIVISOR = 1e6;
    const MONEY_FMT = '"$"#,##0.00';
    const MONEY_FMT_MILL = '"$"#,##0.00" M"';

    function getIdEmpresa(ruc, razonSocial) {
        const r = ruc || "";
        if (r) return r;
        const razon = normalizeName(razonSocial || "");
        if (!razon) return "";
        return `ID_${razon.replace(/[^A-Z0-9]+/g, "_")}`;
    }

    // -------------- Builder Principal --------------
    function buildVentasDesempeno() {
        Logger.log("[DASH4] Iniciando generacion de tablas de desempeno...");

        const empresas = new Map();
        const rucToId = new Map();
        const aniosSet = new Set();
        const trimestresSet = new Set();
        let ventasProcesadas = 0;
        let ventasOmitidas = 0;
        let estadosBaseDetectados = 0;

        const base = readTableFlexibleByCandidates([SHEETS.BASE, SHEETS.BASE_FALLBACK]);
        const ventas = readTableFlexibleByCandidates([SHEETS.VENTAS, SHEETS.VENTAS_FALLBACK]);
        const estados = readTableFlexibleByCandidates([SHEETS.ESTADO, SHEETS.ESTADO_FALLBACK]);
        const sectores = readTableFlexibleByCandidates([SHEETS.SECTOR]);

        Logger.log(
            `[DASH4] Hojas leidas: BASE=${base.rows.length}, VENTAS=${ventas.rows.length}, ESTADO=${estados.rows.length}, SECTOR=${sectores.rows.length}`
        );

        const rucDuplicados = { base: 0, ventas: 0, estado: 0, sector: 0 };
        const seenRucBase = new Set();
        const seenRucVentas = new Set();
        const seenRucEstado = new Set();
        const seenRucSector = new Set();

        const yearColumns = detectYearColumns(base.headerIndex);
        Logger.log(
            `[DASH4] Columnas historicas detectadas: ${yearColumns
                .map((y) => `${y.colName}(idx=${y.colIndex})->${y.year}`)
                .join(", ")}`
        );
        if (yearColumns.length === 0) {
            Logger.log("[DASH4 WARNING] No se detectaron columnas de anos. Verifica columnas T2023, T2022, etc. en BASE DE DATOS");
        }

        const sectorMap = new Map();
        sectores.rows.forEach((row, idx) => {
            const ruc = padRuc13(getVal(row, sectores.headerIndex, ["RUC"]));
            if (ruc) {
                if (seenRucSector.has(ruc)) {
                    rucDuplicados.sector++;
                    if (rucDuplicados.sector <= 3) Logger.log(`[DASH4 WARNING] RUC duplicado en SECTOR fila ${idx + 1}: ${ruc}`);
                }
                seenRucSector.add(ruc);
                const sector = normalizeSector(getVal(row, sectores.headerIndex, ["SECTOR", "SECTOR_ECONOMICO", "ACTIVIDAD"]));
                if (sector) sectorMap.set(ruc, sector);
            }
        });
        Logger.log(`[DASH4] Sectores cargados desde hoja SECTOR: ${sectorMap.size}`);

        base.rows.forEach((row, idx) => {
            const rucRaw = getVal(row, base.headerIndex, ["RUC"]);
            const ruc = padRuc13(rucRaw);
            if (ruc) {
                if (seenRucBase.has(ruc)) {
                    rucDuplicados.base++;
                    if (rucDuplicados.base <= 3) Logger.log(`[DASH4 WARNING] RUC duplicado en BASE fila ${idx + 1}: ${ruc}`);
                }
                seenRucBase.add(ruc);
            }

            let tamano = normalizeTamano(getVal(row, base.headerIndex, ["TAMANO", "TAMANO_EMPRESA", "TAMANIO", "TAMANO_EMP"]));
            if (!tamano) tamano = "DESCONOCIDO";

            const razonSocial = normalizeName(getVal(row, base.headerIndex, ["RAZON_SOCIAL", "RAZON SOCIAL", "EMPRESA", "NOMBRE"]));
            const idEmpresa = getIdEmpresa(ruc, razonSocial);
            if (!idEmpresa) return;

            const empleados = parseEmpleados(
                getVal(row, base.headerIndex, [
                    "EMPLEADOS",
                    "NUM_EMPLEADOS",
                    "NUMERO_EMPLEADOS",
                    "COLABORADORES",
                    "NUM_COLABORADORES",
                    "NUMERO_COLABORADORES",
                    "TRABAJADORES",
                    "NO_COLABORADORES",
                    "NO._COLABORADORES"
                ])
            );

            const sectorBase = normalizeSector(getVal(row, base.headerIndex, ["SECTOR", "SECTOR_ECONOMICO", "ACTIVIDAD"]));
            const sector = sectorMap.get(ruc) || sectorBase || "SIN CLASIFICAR";
            const estadoBaseRaw = getVal(row, base.headerIndex, [
                "ESTADO",
                "ESTADO_PAGO",
                "ESTADO_DE_PAGO",
                "ESTADO_AFILIACION",
                "ESTADO_DE_AFILIACION",
                "ESTADO_ACTUAL"
            ]);
            const estadoBaseNorm = normalizeEstado(estadoBaseRaw);
            const estadoInicial = estadoBaseNorm && estadoBaseNorm !== "DESCONOCIDO" ? estadoBaseNorm : "DESCONOCIDO";
            if (estadoInicial !== "DESCONOCIDO") {
                estadosBaseDetectados++;
            }

            if (empresas.size < 3) {
                Logger.log(
                    `[DASH4 DEBUG] Empresa ${empresas.size + 1}: RUC=${ruc}, tamano='${tamano}', empleados=${empleados}, razon='${razonSocial.substring(
                        0,
                        30
                    )}'`
                );
            }

            let empresaObj = empresas.get(idEmpresa);
            if (!empresaObj) {
                empresaObj = {
                    id: idEmpresa,
                    ruc,
                    razonSocial: razonSocial || "SIN NOMBRE",
                    tamano,
                    empleados,
                    sector,
                    estado: estadoInicial,
                    ventas: {},
                    ventasTrimestre: {}
                };
                empresas.set(idEmpresa, empresaObj);
                if (ruc) rucToId.set(ruc, idEmpresa);
            } else {
                // Refuerza datos faltantes sin sobrescribir ventas ya acumuladas
                if (!empresaObj.razonSocial && razonSocial) empresaObj.razonSocial = razonSocial;
                if (empresaObj.tamano === "DESCONOCIDO" && tamano) empresaObj.tamano = tamano;
                if (!empresaObj.sector && sector) empresaObj.sector = sector;
                if (empresaObj.estado === "DESCONOCIDO" && estadoInicial !== "DESCONOCIDO") empresaObj.estado = estadoInicial;
                if (empresaObj.empleados === 0 && empleados > 0) empresaObj.empleados = empleados;
            }

            yearColumns.forEach(({ year, colName, colIndex }) => {
                const rawVentas = row[colIndex] || "";
                const monto = parseMonto(rawVentas);
                if (monto > 0) {
                    aniosSet.add(year.toString());
                    const emp = empresas.get(idEmpresa);
                    emp.ventas[year.toString()] = (emp.ventas[year.toString()] || 0) + monto;
                }
            });
        });

        Logger.log(`[DASH4] Empresas procesadas desde BASE DE DATOS: ${empresas.size}`);
        Logger.log(`[DASH4] Estados detectados desde BASE DE DATOS: ${estadosBaseDetectados}`);
        Logger.log(`[DASH4] Anos detectados en columnas historicas: ${Array.from(aniosSet).sort().join(", ")}`);

        let estadosActualizados = 0;
        estados.rows.forEach((row, idx) => {
            const ruc = padRuc13(getVal(row, estados.headerIndex, ["RUC"]));
            if (ruc) {
                if (seenRucEstado.has(ruc)) {
                    rucDuplicados.estado++;
                    if (rucDuplicados.estado <= 3) Logger.log(`[DASH4 WARNING] RUC duplicado en ESTADO fila ${idx + 1}: ${ruc}`);
                }
                seenRucEstado.add(ruc);
            }
            const idEmpresa = ruc ? rucToId.get(ruc) : null;
            if (idEmpresa && empresas.has(idEmpresa)) {
                const estadoRaw = getVal(row, estados.headerIndex, ["ESTADO", "ESTADO_PAGO", "PAGADO"]);
                const estadoNorm = normalizeEstado(estadoRaw);
                if (estadoNorm && estadoNorm !== "DESCONOCIDO") {
                    empresas.get(idEmpresa).estado = estadoNorm;
                    estadosActualizados++;
                }
            }
        });
        Logger.log(`[DASH4] Estados actualizados desde ESTADO_AFILIADOS: ${estadosActualizados} de ${empresas.size} empresas`);

        ventas.rows.forEach((row, idx) => {
            const ruc = padRuc13(getVal(row, ventas.headerIndex, ["RUC"]));
            if (ruc) {
                if (seenRucVentas.has(ruc)) {
                    rucDuplicados.ventas++;
                    if (rucDuplicados.ventas <= 3) Logger.log(`[DASH4 WARNING] RUC duplicado en VENTAS fila ${idx + 1}: ${ruc}`);
                }
                seenRucVentas.add(ruc);
            }
            let anio = getVal(row, ventas.headerIndex, ["ANO", "ANIO", "AÑO"]);
            const fechaRegistro = getVal(row, ventas.headerIndex, ["FECHA_REGISTRO", "FECHA", "FECHA_VENTA"]);

            if (!anio) {
                const parsedYear = getYear(fechaRegistro);
                anio = parsedYear;
                if (!parsedYear && fechaRegistro) {
                    Logger.log(`[DASH4 WARNING] Fecha invalida en VENTAS fila ${idx + 1}: '${fechaRegistro}'`);
                }
            }
            const trimestre = getTrimestre(fechaRegistro);

            const rawMonto = getVal(row, ventas.headerIndex, ["MONTO_ESTIMADO", "MONTO", "VALOR", "VENTAS", "PRECIO"]);
            const monto = parseMonto(rawMonto);

            const razonSocialVenta = normalizeName(getVal(row, ventas.headerIndex, ["RAZON_SOCIAL", "EMPRESA", "NOMBRE"]));
            const idEmpresa = getIdEmpresa(ruc, razonSocialVenta);

            if (idEmpresa && anio && monto > 0) {
                aniosSet.add(anio);
                if (trimestre) trimestresSet.add(`${anio}-${trimestre}`);

                if (!empresas.has(idEmpresa)) {
                    const sector = sectorMap.get(ruc) || normalizeSector(getVal(row, ventas.headerIndex, ["SECTOR"]));

                    empresas.set(idEmpresa, {
                        id: idEmpresa,
                        ruc,
                        razonSocial: razonSocialVenta || "SIN NOMBRE",
                        tamano: "DESCONOCIDO",
                        empleados: 0,
                        sector: sector || "SIN CLASIFICAR",
                        estado: "DESCONOCIDO",
                        ventas: {},
                        ventasTrimestre: {}
                    });
                    if (ruc) rucToId.set(ruc, idEmpresa);
                }

                const emp = empresas.get(idEmpresa);
                emp.ventas[anio] = (emp.ventas[anio] || 0) + monto;

                if (trimestre) {
                    const keyTrimestre = `${anio}-${trimestre}`;
                    emp.ventasTrimestre[keyTrimestre] = (emp.ventasTrimestre[keyTrimestre] || 0) + monto;
                }

                ventasProcesadas++;
            } else {
                ventasOmitidas++;
                if (ventasOmitidas <= 3) {
                    Logger.log(
                        `[DASH4 DEBUG] Venta omitida fila ${idx + 1}: RUC=${ruc}, ANIO=${anio}, MONTO_RAW='${rawMonto}', MONTO_PARSED=${monto}`
                    );
                }
            }
        });

        const aniosOrdenados = Array.from(aniosSet).sort((a, b) => b.localeCompare(a));
        const trimestresOrdenados = Array.from(trimestresSet).sort((a, b) => b.localeCompare(a));
        const totalColaboradores = Array.from(empresas.values()).reduce((sum, emp) => sum + emp.empleados, 0);

        Logger.log("[DASH4] ========== RESUMEN FINAL ==========");
        Logger.log(`[DASH4] Total empresas: ${empresas.size}`);
        Logger.log(`[DASH4] Total colaboradores: ${totalColaboradores}`);
        Logger.log(`[DASH4] Ventas procesadas: ${ventasProcesadas}, Omitidas: ${ventasOmitidas}`);
        Logger.log(`[DASH4] Anos detectados: ${aniosOrdenados.join(", ")}`);
        Logger.log(`[DASH4] Trimestres detectados: ${trimestresOrdenados.join(", ")}`);
        if (rucDuplicados.base || rucDuplicados.ventas || rucDuplicados.estado || rucDuplicados.sector) {
            Logger.log(
                `[DASH4 WARNING] RUC duplicados detectados -> BASE:${rucDuplicados.base}, VENTAS:${rucDuplicados.ventas}, ESTADO:${rucDuplicados.estado}, SECTOR:${rucDuplicados.sector}`
            );
        }

        generateResumenTable(empresas, aniosOrdenados);
        generateSectorTable(empresas, aniosOrdenados);
        generateTopEmpresasTable(empresas, aniosOrdenados);
        generateTopEmpresasPorAnio(empresas, aniosOrdenados);
        generateEstadoTable(empresas, aniosOrdenados);
        generateSemaforoTable(empresas, aniosOrdenados);
        generateTrimestreTable(empresas, trimestresOrdenados);
        generateTopSectoresRelevantes(empresas, aniosOrdenados);
        generateTopEmpresasRelevantes(empresas, aniosOrdenados);
        generateMasterTable(empresas, aniosOrdenados, trimestresOrdenados);
        generateMergedOutputs();
        generateSlicerMasterSheet();

        Logger.log("[DASH4] Tablas de desempeno generadas exitosamente.");
    }
    function generateResumenTable(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "VENTAS_TOTALES", "VENTAS_TOTALES_M", "EMPRESAS", "COLABORADORES"];

        anios.forEach((anio) => {
            const grupos = new Map();
            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const key = `${emp.tamano}|${emp.estado}`;
                    const current = grupos.get(key) || { ventas: 0, empresas: 0, colaboradores: 0 };
                    current.ventas += emp.ventas[anio];
                    current.empresas += 1;
                    current.colaboradores += emp.empleados;
                    grupos.set(key, current);
                }
            });
            grupos.forEach((data, key) => {
                const [tamano, estado] = key.split("|");
                rows.push([anio, tamano, estado, data.ventas, data.ventas / MILLION_DIVISOR, data.empresas, data.colaboradores]);
            });
        });

        if (rows.length > 0) {
            Logger.log("[DASH4 DEBUG] Sample rows from RESUMEN table (first 3):");
            rows.slice(0, 3).forEach((row, idx) => {
                Logger.log(`  Row ${idx}: anio=${row[0]}, tamano='${row[1]}', estado='${row[2]}', ventas=${row[3]}, empresas=${row[4]}, colab=${row[5]}`);
            });
        }

        rows.sort((a, b) => {
            if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
            const ordenA = TAMANO_ORDER[a[1]] || 99;
            const ordenB = TAMANO_ORDER[b[1]] || 99;
            return ordenA - ordenB;
        });

        const sheet = getSheetOrCreate(SHEETS.OUT_RESUMEN);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        // Ventas totales formato moneda (bruto y en millones)
        formatNumberColumn(sheet, 2, 4, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_RESUMEN}: ${rows.length} filas`);
    }

    function generateSectorTable(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "SECTOR", "VENTAS_MONTO", "VENTAS_MONTO_M", "EMPRESAS"];

        anios.forEach((anio) => {
            const grupos = new Map();
            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const key = `${emp.tamano}|${emp.estado}|${emp.sector}`;
                    const current = grupos.get(key) || { ventas: 0, empresas: 0 };
                    current.ventas += emp.ventas[anio];
                    current.empresas += 1;
                    grupos.set(key, current);
                }
            });
            grupos.forEach((data, key) => {
                const [tamano, estado, sector] = key.split("|");
                rows.push([anio, tamano, estado, sector, data.ventas, data.ventas / MILLION_DIVISOR, data.empresas]);
            });
        });

        rows.sort((a, b) => {
            if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
            return b[4] - a[4];
        });

        const sheet = getSheetOrCreate(SHEETS.OUT_SECTOR);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        // Ventas por sector formato moneda
        formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_SECTOR}: ${rows.length} filas`);
    }

    function generateTopEmpresasTable(empresas, anios, topN = 5) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "RUC", "RAZON_SOCIAL", "SECTOR", "VENTAS_MONTO", "VENTAS_MONTO_M", "RANK"];

        anios.forEach((anio) => {
            const grupos = new Map();
            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const key = `${emp.tamano}|${emp.estado}`;
                    if (!grupos.has(key)) grupos.set(key, []);
                    grupos.get(key).push({ emp, ventas: emp.ventas[anio] });
                }
            });
            grupos.forEach((lista, key) => {
                const [tamano, estado] = key.split("|");
                lista.sort((a, b) => b.ventas - a.ventas);
                lista.slice(0, topN).forEach((item, idx) => {
                    rows.push([anio, tamano, estado, item.emp.ruc || item.emp.id, item.emp.razonSocial, item.emp.sector, item.ventas, item.ventas / MILLION_DIVISOR, idx + 1]);
                });
            });
        });

        const sheet = getSheetOrCreate(SHEETS.OUT_TOP);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 7, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 8, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_TOP}: ${rows.length} filas`);
    }

    function generateTopEmpresasPorAnio(empresas, anios, topN = 5) {
        const rows = [];
        const headerRow = ["ANIO", "RUC", "RAZON_SOCIAL", "SECTOR", "TAMANO", "ESTADO", "VENTAS_MONTO", "VENTAS_MONTO_M", "RANK"];

        anios.forEach((anio) => {
            const lista = [];
            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    lista.push({ emp, ventas: emp.ventas[anio] });
                }
            });
            lista.sort((a, b) => b.ventas - a.ventas);
            lista.slice(0, topN).forEach((item, idx) => {
                rows.push([anio, item.emp.ruc || item.emp.id, item.emp.razonSocial, item.emp.sector, item.emp.tamano, item.emp.estado, item.ventas, item.ventas / MILLION_DIVISOR, idx + 1]);
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate(SHEETS.OUT_TOP_EMPRESAS_ANIO);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 7, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 8, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_TOP_EMPRESAS_ANIO}: ${rows.length} filas`);
    }
    function generateEstadoTable(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "EMPRESAS", "PCT"];

        anios.forEach((anio) => {
            const grupos = new Map();
            const totalPorTamano = new Map();

            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const key = `${emp.tamano}|${emp.estado}`;
                    grupos.set(key, (grupos.get(key) || 0) + 1);
                    totalPorTamano.set(emp.tamano, (totalPorTamano.get(emp.tamano) || 0) + 1);
                }
            });
            grupos.forEach((count, key) => {
                const [tamano, estado] = key.split("|");
                const total = totalPorTamano.get(tamano) || 1;
                const pct = (count / total) * 100;
                rows.push([anio, tamano, estado, count, pct]);
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate(SHEETS.OUT_ESTADO);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        Logger.log(`[DASH4] ${SHEETS.OUT_ESTADO}: ${rows.length} filas`);
    }

    function generateSemaforoTable(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "TENDENCIA", "EMPRESAS"];

        anios.forEach((anio) => {
            const anioAnterior = (parseInt(anio, 10) - 1).toString();
            const grupos = new Map();

            empresas.forEach((emp) => {
                const ventasActual = emp.ventas[anio] || 0;
                const ventasAnterior = emp.ventas[anioAnterior] || 0;

                if (ventasActual > 0 && ventasAnterior > 0) {
                    let tendencia = "IGUAL";
                    if (ventasActual > ventasAnterior) tendencia = "AUMENTO";
                    else if (ventasActual < ventasAnterior) tendencia = "DISMINUCION";

                    const key = `${emp.tamano}|${emp.estado}|${tendencia}`;
                    grupos.set(key, (grupos.get(key) || 0) + 1);
                } else if (ventasActual > 0 && ventasAnterior === 0) {
                    const key = `${emp.tamano}|${emp.estado}|NUEVO`;
                    grupos.set(key, (grupos.get(key) || 0) + 1);
                }
            });
            grupos.forEach((count, key) => {
                const [tamano, estado, tendencia] = key.split("|");
                rows.push([anio, tamano, estado, tendencia, count]);
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate(SHEETS.OUT_SEMAFORO);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        Logger.log(`[DASH4] ${SHEETS.OUT_SEMAFORO}: ${rows.length} filas`);
    }

    function generateTrimestreTable(empresas, trimestres) {
        const rows = [];
        const headerRow = ["ANIO_TRIMESTRE", "ANIO", "TRIMESTRE", "TAMANO", "ESTADO", "VENTAS_MONTO", "VENTAS_MONTO_M", "EMPRESAS"];

        trimestres.forEach((anioTrimestre) => {
            const [anio, trimestre] = anioTrimestre.split("-");
            const grupos = new Map();

            empresas.forEach((emp) => {
                if (emp.ventasTrimestre && emp.ventasTrimestre[anioTrimestre]) {
                    const key = `${emp.tamano}|${emp.estado}`;
                    const current = grupos.get(key) || { ventas: 0, empresas: 0 };
                    current.ventas += emp.ventasTrimestre[anioTrimestre];
                    current.empresas += 1;
                    grupos.set(key, current);
                }
            });

            grupos.forEach((data, key) => {
                const [tamano, estado] = key.split("|");
                rows.push([anioTrimestre, anio, trimestre, tamano, estado, data.ventas, data.ventas / MILLION_DIVISOR, data.empresas]);
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate("PIVOT_VENTAS_TRIMESTRE");
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 7, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] PIVOT_VENTAS_TRIMESTRE: ${rows.length} filas`);
    }

    function generateTopSectoresRelevantes(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "SECTOR", "VENTAS_MONTO", "VENTAS_MONTO_M"];

        anios.forEach((anio) => {
            const grupos = new Map();
            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const key = `${emp.tamano}|${emp.estado}|${emp.sector}`;
                    const current = grupos.get(key) || 0;
                    grupos.set(key, current + emp.ventas[anio]);
                }
            });

            const porGrupo = new Map();
            grupos.forEach((ventas, key) => {
                const [tamano, estado, sector] = key.split("|");
                const groupKey = `${anio}|${tamano}|${estado}`;
                if (!porGrupo.has(groupKey)) porGrupo.set(groupKey, []);
                porGrupo.get(groupKey).push({ sector, ventas });
            });

            porGrupo.forEach((lista, groupKey) => {
                const [gAnio, tamano, estado] = groupKey.split("|");
                lista.sort((a, b) => b.ventas - a.ventas);
                lista.slice(0, 5).forEach((item) => {
                    rows.push([gAnio, tamano, estado, item.sector, item.ventas, item.ventas / MILLION_DIVISOR]);
                });
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate(SHEETS.OUT_TOP_SECTORES_REL);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_TOP_SECTORES_REL}: ${rows.length} filas`);
    }

    function generateTopEmpresasRelevantes(empresas, anios) {
        const rows = [];
        const headerRow = ["ANIO", "TAMANO", "ESTADO", "RAZON_SOCIAL", "VENTAS_MONTO", "VENTAS_MONTO_M"];

        anios.forEach((anio) => {
            const porGrupo = new Map();

            empresas.forEach((emp) => {
                if (emp.ventas[anio] && emp.ventas[anio] > 0) {
                    const groupKey = `${anio}|${emp.tamano}|${emp.estado}`;
                    if (!porGrupo.has(groupKey)) porGrupo.set(groupKey, []);
                    porGrupo.get(groupKey).push({
                        razon: emp.razonSocial,
                        ventas: emp.ventas[anio]
                    });
                }
            });

            porGrupo.forEach((lista, groupKey) => {
                const [gAnio, tamano, estado] = groupKey.split("|");
                lista.sort((a, b) => b.ventas - a.ventas);
                lista.slice(0, 5).forEach((item) => {
                    rows.push([gAnio, tamano, estado, item.razon, item.ventas, item.ventas / MILLION_DIVISOR]);
                });
            });
        });

        rows.sort((a, b) => b[0].localeCompare(a[0]));

        const sheet = getSheetOrCreate(SHEETS.OUT_TOP_EMPRESAS_REL);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 5, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 6, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_TOP_EMPRESAS_REL}: ${rows.length} filas`);
    }

    function generateMasterTable(empresas, anios, trimestres) {
        const rows = [];
        const headerRow = [
            "PERIODO_TIPO",
            "PERIODO_ID",
            "ANIO",
            "TRIMESTRE",
            "ID_EMPRESA",
            "RUC",
            "RAZON_SOCIAL",
            "SECTOR",
            "TAMANO",
            "ESTADO",
            "VENTAS_MONTO",
            "VENTAS_MONTO_M",
            "VENTAS_PREV",
            "VENTAS_PREV_M",
            "TENDENCIA",
            "COLABORADORES"
        ];

        anios.forEach((anio) => {
            const anioAnterior = (parseInt(anio, 10) - 1).toString();
            empresas.forEach((emp) => {
                const ventas = emp.ventas[anio] || 0;
                const ventasPrev = emp.ventas[anioAnterior] || 0;
                if (ventas > 0 || ventasPrev > 0) {
                    let tendencia = "";
                    if (ventas > 0 && ventasPrev > 0) {
                        if (ventas > ventasPrev) tendencia = "AUMENTO";
                        else if (ventas < ventasPrev) tendencia = "DISMINUCION";
                        else tendencia = "IGUAL";
                    } else if (ventas > 0 && ventasPrev === 0) {
                        tendencia = "NUEVO";
                    } else if (ventas === 0 && ventasPrev > 0) {
                        tendencia = "SIN_VENTAS_ACTUAL";
                    } else {
                        tendencia = "SIN_DATOS";
                    }

                    rows.push([
                        "ANUAL",
                        anio,
                        anio,
                        "",
                        emp.id,
                        emp.ruc || emp.id,
                        emp.razonSocial,
                        emp.sector,
                        emp.tamano,
                        emp.estado,
                        ventas,
                        ventas / MILLION_DIVISOR,
                        ventasPrev,
                        ventasPrev / MILLION_DIVISOR,
                        tendencia,
                        emp.empleados
                    ]);
                }
            });
        });

        trimestres.forEach((anioTrimestre) => {
            const [anio, trimestre] = anioTrimestre.split("-");
            empresas.forEach((emp) => {
                const ventasTrim = emp.ventasTrimestre ? emp.ventasTrimestre[anioTrimestre] || 0 : 0;
                if (ventasTrim > 0) {
                    rows.push([
                        "TRIMESTRE",
                        anioTrimestre,
                        anio,
                        trimestre,
                        emp.id,
                        emp.ruc || emp.id,
                        emp.razonSocial,
                        emp.sector,
                        emp.tamano,
                        emp.estado,
                        ventasTrim,
                        ventasTrim / MILLION_DIVISOR,
                        "",
                        "",
                        "",
                        emp.empleados
                    ]);
                }
            });
        });

        rows.sort((a, b) => {
            if (a[2] !== b[2]) return b[2].localeCompare(a[2]);
            if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
            if (a[3] !== b[3]) return b[3].localeCompare(a[3]);
            return (a[5] || "").localeCompare(b[5] || "");
        });

        const sheet = getSheetOrCreate(SHEETS.OUT_MASTER);
        sheet.clear();
        const output = rows.length ? [headerRow, ...rows] : [headerRow];
        sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
        sheet.autoResizeColumns(1, headerRow.length);
        formatNumberColumn(sheet, 2, 11, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 12, rows.length, MONEY_FMT_MILL);
        formatNumberColumn(sheet, 2, 13, rows.length, MONEY_FMT);
        formatNumberColumn(sheet, 2, 14, rows.length, MONEY_FMT_MILL);
        Logger.log(`[DASH4] ${SHEETS.OUT_MASTER}: ${rows.length} filas`);
    }

    function generateMergedOutputs() {
        const outputSheets = [
            SHEETS.OUT_RESUMEN,
            SHEETS.OUT_SECTOR,
            SHEETS.OUT_TOP,
            SHEETS.OUT_TOP_EMPRESAS_ANIO,
            SHEETS.OUT_ESTADO,
            SHEETS.OUT_SEMAFORO,
            "PIVOT_VENTAS_TRIMESTRE",
            SHEETS.OUT_TOP_SECTORES_REL,
            SHEETS.OUT_TOP_EMPRESAS_REL,
            SHEETS.OUT_MASTER
        ];

        const ss = SpreadsheetApp.getActive();
        const headerSet = new Set(["SOURCE_SHEET"]);
        const sheetsData = [];

        outputSheets.forEach((name) => {
            const sh = ss.getSheetByName(name);
            if (!sh) return;
            const values = sh.getDataRange().getDisplayValues();
            if (!values || values.length === 0) return;
            const header = values[0].map((h, idx) => {
                const clean = (h || "").toString().trim();
                return clean || `COL_${idx + 1}`;
            });
            header.forEach((h) => headerSet.add(h));
            const rows = values.slice(1);
            sheetsData.push({ name, header, rows });
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

        const sheet = getSheetOrCreate(SHEETS.OUT_MASTER_ALL);
        sheet.clear();
        const output = mergedRows.length ? [masterHeaders, ...mergedRows] : [masterHeaders];
        sheet.getRange(1, 1, output.length, masterHeaders.length).setValues(output);
        sheet.autoResizeColumns(1, masterHeaders.length);
        Logger.log(`[DASH4] ${SHEETS.OUT_MASTER_ALL}: ${mergedRows.length} filas (merge de pivots)`);
    }

    function generateSlicerMasterSheet() {
        // Une solo las tablas usadas en el dashboard para que un slicer controle todo.
        const slicerSources = [
            { name: SHEETS.OUT_RESUMEN, cols: 7 }, // PIVOT_VENTAS_RESUMEN_ANIO!A:G
            { name: SHEETS.OUT_SEMAFORO, cols: 5 }, // PIVOT_SEMAFORO_VENTAS!A:E
            { name: SHEETS.OUT_TOP, cols: 8 }, // PIVOT_TOP_EMPRESAS_VENTAS!A:H
            { name: SHEETS.OUT_ESTADO, cols: 5 }, // PIVOT_ESTADO_EMPRESAS!A:E
            { name: SHEETS.OUT_TOP_SECTORES_REL, cols: 6 } // PIVOT_TOP_SECTORES_RELEVANTES!A:F
        ];

        const ss = SpreadsheetApp.getActive();
        const headerSet = new Set(["SOURCE_SHEET"]);
        const sheetsData = [];

        slicerSources.forEach(({ name, cols }) => {
            const sh = ss.getSheetByName(name);
            if (!sh) return;
            const numRows = sh.getLastRow();
            if (numRows === 0) return;

            // Usar valores crudos (no display) para conservar números y que las tablas dinámicas sumen correctamente.
            const values = sh.getRange(1, 1, numRows, cols).getValues();
            if (!values || values.length === 0) return;

            const header = values[0]
                .slice(0, cols)
                .map((h, idx) => {
                    const clean = (h || "").toString().trim();
                    return clean || `COL_${idx + 1}`;
                });
            header.forEach((h) => headerSet.add(h));

            // Guarda filas no vacias recortadas al numero de columnas esperado.
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

        // Formato numérico para columnas de montos (bruto y millones) manteniendo datos como número.
        if (mergedRows.length) {
            const headersWritten = sheet.getRange(1, 1, 1, masterHeaders.length).getValues()[0].map((h) => (h || "").toString().trim().toUpperCase());
            const moneyCols = [];
            const moneyMCols = [];
            headersWritten.forEach((h, idx) => {
                if (h === "VENTAS_TOTALES" || h === "VENTAS_MONTO" || h === "VENTAS_PREV") moneyCols.push(idx + 1);
                if (h === "VENTAS_TOTALES_M" || h === "VENTAS_MONTO_M" || h === "VENTAS_PREV_M") moneyMCols.push(idx + 1);
            });
            moneyCols.forEach((col) => formatNumberColumn(sheet, 2, col, mergedRows.length, MONEY_FMT));
            moneyMCols.forEach((col) => formatNumberColumn(sheet, 2, col, mergedRows.length, MONEY_FMT_MILL));
        }

        Logger.log(`[DASH4] ${SHEETS.OUT_MASTER_SLICER}: ${mergedRows.length} filas (merge para slicers)`);
    }

    function refreshDashboardDesempeno() {
        buildVentasDesempeno();
        Logger.log("Tablas de desempeno regeneradas.");
    }

    function onOpenDash4() {
        SpreadsheetApp.getUi().createMenu("DASH4").addItem("Generar tablas desempeno", "refreshDashboardDesempeno").addToUi();
    }

    // Trigger manual via checkbox (hoja REPORTE_1, celda Q157)
    const DASH4_TRIGGER_SHEET = "REPORTE_1";
    const DASH4_TRIGGER_CELL = "Q157";

    function onEditDash4(e) {
        const range = e.range;
        if (!range) return;
        const sheet = range.getSheet();
        if (!sheet || sheet.getName() !== DASH4_TRIGGER_SHEET) return;
        if (range.getA1Notation() !== DASH4_TRIGGER_CELL) return;

        const val = range.getValue();
        if (val === true) {
            const ss = SpreadsheetApp.getActive();
            ss.toast("Actualizando dashboard…");
            try {
                refreshDashboardDesempeno();
                ss.toast("Dashboard listo.");
                sheet.getRange("Q158").setValue("Ultima actualizacion: " + new Date());
            } finally {
                range.setValue(false); // deja la casilla lista para el siguiente clic
            }
        }
    }

    global.refreshDashboardDesempeno = refreshDashboardDesempeno;
    global.onOpenDash4 = onOpenDash4;
    global.onEditDash4 = onEditDash4;
})(this);

function refreshDashboardDesempenoWrapper() {
    return refreshDashboardDesempeno();
}

function onOpenDash4Wrapper() {
    return onOpenDash4();
}

// Wrapper para que el trigger instalable detecte la función en el selector
function onEditDash4Wrapper(e) {
    return onEditDash4(e);
}

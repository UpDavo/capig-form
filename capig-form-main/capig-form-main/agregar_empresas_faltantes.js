/**
 * Script para agregar empresas faltantes a SOCIOS.
 * Ejecutar una sola vez: agregarEmpresasFaltantes()
 */

function agregarEmpresasFaltantes() {
  const ss = SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName("SOCIOS") || ss.getSheetByName("BASE DE DATOS");

  if (!baseSheet) {
    Logger.log("ERROR: No se encontro la hoja 'SOCIOS' ni 'BASE DE DATOS'");
    return;
  }

  // Encontrar la fila de encabezados
  const allData = baseSheet.getDataRange().getDisplayValues();
  let headerRow = 0;
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i].join("").toUpperCase();
    if (row.indexOf("RUC") !== -1 && row.indexOf("RAZON") !== -1) {
      headerRow = i + 1;
      break;
    }
  }

  if (headerRow === 0) {
    Logger.log("ERROR: No se encontro fila de encabezados");
    return;
  }

  const headers = allData[headerRow - 1];
  const rucCol = headers.findIndex((h) => h.toUpperCase().indexOf("RUC") !== -1) + 1;
  const razonCol = headers.findIndex((h) => h.toUpperCase().indexOf("RAZON") !== -1) + 1;
  const tamanoCol = headers.findIndex((h) => h.toUpperCase().indexOf("TAMANO") !== -1) + 1;
  const sectorCol = headers.findIndex((h) => h.toUpperCase().indexOf("SECTOR") !== -1) + 1;

  Logger.log(`Columnas detectadas: RUC=${rucCol}, RAZON=${razonCol}, TAMANO=${tamanoCol}, SECTOR=${sectorCol}`);

  // Empresas a agregar
  const nuevasEmpresas = [
    {
      ruc: "0991300333001",
      razonSocial: "TONISA S.A.",
      tamano: "MEDIANA",
      sector: "COMERCIO",
    },
    {
      ruc: "0992257946001",
      razonSocial: "ECUASERVIGLOBAL S.A.",
      tamano: "MEDIANA",
      sector: "SERVICIOS",
    },
    {
      ruc: "0991318380001",
      razonSocial: "CONSTRUCCIONES CIVILES Y METALICAS CONSTRUME S.A.",
      tamano: "GRANDE",
      sector: "CONSTRUCCION",
    },
  ];

  // Verificar si ya existen
  const rucExistentes = baseSheet.getRange(headerRow + 1, rucCol, baseSheet.getLastRow() - headerRow, 1).getValues().flat();

  let agregadas = 0;
  nuevasEmpresas.forEach((empresa) => {
    if (rucExistentes.includes(empresa.ruc)) {
      Logger.log(`SKIP: ${empresa.razonSocial} (RUC ${empresa.ruc}) ya existe`);
    } else {
      const nuevaFila = baseSheet.getLastRow() + 1;

      if (rucCol) baseSheet.getRange(nuevaFila, rucCol).setValue(empresa.ruc);
      if (razonCol) baseSheet.getRange(nuevaFila, razonCol).setValue(empresa.razonSocial);
      if (tamanoCol) baseSheet.getRange(nuevaFila, tamanoCol).setValue(empresa.tamano);
      if (sectorCol) baseSheet.getRange(nuevaFila, sectorCol).setValue(empresa.sector);

      Logger.log(`AGREGADA: ${empresa.razonSocial} (RUC ${empresa.ruc})`);
      agregadas++;
    }
  });

  Logger.log("\n========== RESUMEN ==========");
  Logger.log(`Empresas agregadas: ${agregadas}`);
  Logger.log(`Total en hoja base: ${baseSheet.getLastRow() - headerRow}`);
  Logger.log("\nListo! Ahora ejecuta refreshDashboardDiagnosticos()");
}

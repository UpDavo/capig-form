"""
Job local para generar la tabla de porcentaje de gerentes por género y tamaño
desde el Excel consolidado (`data/datos_completos.xlsx`).

No toca Google Sheets ni otras lógicas del sistema.
"""
import os
from typing import Dict

import pandas as pd


BASE_FILENAME = "datos_completos.xlsx"
OUTPUT_FILENAME = "gerentes_genero_por_tamano.xlsx"

# Ruta base y salida (ambas relativas al repo, sin tocar hojas remotas)
BASE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "data", BASE_FILENAME)
)
OUTPUT_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "data", OUTPUT_FILENAME)
)

TAMANO_ORDER = ["MICRO", "PEQUENA", "MEDIANA", "GRANDE"]


def _normalize_ruc(value) -> str:
    return str(value or "").replace("'", "").replace('"', "").strip().lstrip("0")


def _normalize_genero(value: str) -> str:
    txt = str(value or "").strip().upper()
    for a, b in [("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")]:
        txt = txt.replace(a, b)
    mapping: Dict[str, str] = {
        "F": "FEMENINO",
        "M": "MASCULINO",
        "FEMENINA": "FEMENINO",
        "MASCULINA": "MASCULINO",
        "": "",
        "NAN": "",
        "NONE": "",
        "NA": "",
        "0": "",
    }
    return mapping.get(txt, txt)


def build_dataframe(base_path: str) -> pd.DataFrame:
    """
    Lee BASE DE DATOS y TAMANO_EMPRESA_GLOBAL, normaliza RUC/género
    y devuelve un dataframe con columnas: RUC_norm, TAM_G, genero_norm.
    """
    base_raw = pd.read_excel(base_path, sheet_name="BASE DE DATOS", header=None).fillna("")
    tam_raw = pd.read_excel(base_path, sheet_name="TAMANO_EMPRESA_GLOBAL").fillna("")

    # Filtrar filas de empresas (primera col numérica)
    base_df = base_raw[pd.to_numeric(base_raw.iloc[:, 0], errors="coerce").notna()].copy()
    base_df["RUC_norm"] = base_df.iloc[:, 2].apply(_normalize_ruc)
    base_df["genero_norm"] = base_df.iloc[:, 15].apply(_normalize_genero)

    tam_df = tam_raw.copy()
    tam_df["RUC_norm"] = tam_df["RUC"].apply(_normalize_ruc)
    tam_df["TAM_G"] = tam_df.iloc[:, 1].astype(str).str.strip().str.upper()

    merged = base_df.merge(tam_df[["RUC_norm", "TAM_G"]], on="RUC_norm", how="inner")
    # Mantener solo filas con género y tamaño válidos
    merged = merged[(merged["genero_norm"] != "") & merged["TAM_G"].notna()]
    merged["TAM_G"] = merged["TAM_G"].str.strip().str.upper()
    return merged[["RUC_norm", "TAM_G", "genero_norm"]]


def summarize(df: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve un resumen con conteos y porcentajes por tamaño.
    """
    ct = df.groupby(["TAM_G", "genero_norm"]).size().unstack(fill_value=0)

    rows = []
    for tam in TAMANO_ORDER:
        if tam not in ct.index:
            continue
        fem = int(ct.loc[tam].get("FEMENINO", 0))
        mas = int(ct.loc[tam].get("MASCULINO", 0))
        total = fem + mas
        if total == 0:
            pct_f = pct_m = 0.0
        else:
            pct_f = round(fem * 100.0 / total, 2)
            pct_m = round(mas * 100.0 / total, 2)
        rows.append(
            {
                "tamano": tam,
                "femenino": fem,
                "masculino": mas,
                "total": total,
                "pct_femenino": pct_f,
                "pct_masculino": pct_m,
            }
        )

    # Global
    fem_global = int(df["genero_norm"].eq("FEMENINO").sum())
    mas_global = int(df["genero_norm"].eq("MASCULINO").sum())
    total_global = fem_global + mas_global
    pct_f_global = round(fem_global * 100.0 / total_global, 2) if total_global else 0.0
    pct_m_global = round(mas_global * 100.0 / total_global, 2) if total_global else 0.0
    rows.append(
        {
            "tamano": "GLOBAL",
            "femenino": fem_global,
            "masculino": mas_global,
            "total": total_global,
            "pct_femenino": pct_f_global,
            "pct_masculino": pct_m_global,
        }
    )
    return pd.DataFrame(rows)


def run():
    if not os.path.exists(BASE_PATH):
        raise FileNotFoundError(f"No se encontró el archivo base: {BASE_PATH}")

    df = build_dataframe(BASE_PATH)
    resumen = summarize(df)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH) as writer:
        resumen.to_excel(writer, sheet_name="resumen", index=False)
        # Tabla de conteos crudos
        ct = df.groupby(["TAM_G", "genero_norm"]).size().unstack(fill_value=0)
        ct.to_excel(writer, sheet_name="crosstab")
        # Detalle (solo columnas necesarias)
        df.to_excel(writer, sheet_name="detalle", index=False)

    print(f"Archivo generado: {OUTPUT_PATH}")


if __name__ == "__main__":
    run()

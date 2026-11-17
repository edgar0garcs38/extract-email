#logic: get number of poliza from the excel given from ACS and search for the number of poliza in the base of 

import pandas as pd

# ---------- Archivos ----------
archivo_periodo_gracia = "Nov-01-2025 en Gracia/POLIZAS EN PERIODO DE GRACIA 1 DE DICIEMBRE 2025.xlsx"
archivo_base_clientes = "Sep-29-2025/REPORTE PERIODO DE GRACIA CON EMAIL 2 SEPTIEMBRE.xlsx"


# ---------- Leer archivos ----------
df_periodo_gracia = pd.read_excel(archivo_periodo_gracia, sheet_name="Hoja1", engine="openpyxl")
df_periodo_gracia = df_periodo_gracia.rename(columns=lambda x: str(x).strip())

df_base_clientes = pd.read_excel(archivo_base_clientes, sheet_name="Sheet1", engine="openpyxl")
df_base_clientes = df_base_clientes.rename(columns=lambda x: str(x).strip())

# ---------- Funcion eliminar duplicados ----------
def drop_duplicate_contracts(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates(subset=["Póliza"], keep="first").reset_index(drop=True)

if "Póliza" in df_base_clientes.columns:
    df_periodo_gracia = df_periodo_gracia[~df_periodo_gracia["Póliza"].isin(df_base_clientes["Póliza"])]
    


df_periodo_gracia.to_excel("Clientes Sin Duplicar.xlsx", index=False)
print("✅ Finished filling commissions.")


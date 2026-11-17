import pandas as pd

# ---------- Archivos ----------
archivo_periodo_gracia = "CLIENTES CANCELADOS OCTUBRE 13 FN LN.xlsx"
archivo_base_clientes = "BMI_REPORTE_A_SU_MEDIDA_525_20250923_15_17_15_5555_20250923.xlsx"

# ---------- Leer archivos ----------
df_periodo_gracia = pd.read_excel(archivo_periodo_gracia, sheet_name="Sheet1", engine="openpyxl")
df_periodo_gracia = df_periodo_gracia.rename(columns=lambda x: str(x).strip())

df_base_clientes = pd.read_excel(archivo_base_clientes, engine="openpyxl")
df_base_clientes = df_base_clientes.rename(columns=lambda x: str(x).strip())

# ---------- Crear diccionario: póliza → email ----------
map_poliza_email = df_base_clientes.set_index("Póliza/Contrato No.")["Correo Electrónic Personal"].to_dict()

# ---------- Insertar columna de email en base a póliza ----------
df_periodo_gracia.insert(
    0,
    "Email Cliente",
    df_periodo_gracia["Póliza"].map(map_poliza_email)
)

# ---------- Exportar resultado ----------
df_periodo_gracia.to_excel("Excel Final.xlsx", index=False)
print("✅ Archivo generado: Excel Final.xlsx")

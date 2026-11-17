import pandas as pd
import re

# ============================================================
# Helpers
# ============================================================

def clean_poliza(value: str) -> str:
    """Normalize contract numbers (remove spaces, dashes, leading zeros)."""
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)  # keep only alphanumerics
    return s.lstrip("0")  # remove leading zeros

def is_valid_email(email: str) -> bool:
    """Validate email addresses, including multi-level domains, and reject empty values."""
    if not email or pd.isna(email):
        return False
    email = str(email).strip()
    if email == "":
        return False

    pattern = (
        r'^(?!.*\.\.)'                                # no consecutive dots
        r'[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]{0,62}[a-zA-Z0-9])?'  # local part
        r'@'
        r'(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+'  # domain labels
        r'[a-zA-Z]{2,}$'                              # final TLD (letters only, ≥2 chars)
    )
    return re.match(pattern, email) is not None

import pandas as pd
import re

def split_names(
    df: pd.DataFrame,
    source_col: str = "Asegurado",
    first_col: str = "First Name",
    last_col: str = "Last Name"
) -> pd.DataFrame:
    """
    Split a name column into first and last name safely.
    Logic:
    - If the value contains a comma, assume "Last, First" and split on first comma.
    - If whitespace:
        * Two words -> split on first whitespace.
        * More than two words -> split on second whitespace.
    - Multi-word last names are preserved.
    - Handles NaN or empty values gracefully.
    """

    def split_name(full_name):
        s = "" if pd.isna(full_name) else str(full_name).strip()
        if s == "":
            return "", ""
        # If comma present, treat as "Last, First"
        if "," in s:
            parts = [p.strip() for p in s.split(",", 1)]
            last = parts[0].upper()
            first = parts[1].upper() if len(parts) > 1 else ""
            return first, last

        # Normalize spaces
        tokens = re.sub(r"\s+", " ", s).split(" ")

        if len(tokens) == 1:
            return tokens[0].upper(), ""
        elif len(tokens) == 2:
            # Two words -> split on first whitespace
            return tokens[0].upper(), tokens[1].upper()
        else:
            # More than two words -> split on second whitespace
            first = " ".join(tokens[:2]).upper()
            last = " ".join(tokens[2:]).upper()
            return first, last

    if source_col not in df.columns:
        df[first_col] = ""
        df[last_col] = ""
        return df

    df[[first_col, last_col]] = df[source_col].apply(lambda x: pd.Series(split_name(x)))
    return df

# ============================================================
# Load input files
# ============================================================

archivo_periodo_gracia = "Mar - 04 - 2026 En Gracia/Vida Y Ahorro/Vida Y Ahorro.xlsx"
archivo_base_clientes = "Feb - 04 - 2026 En Gracia/POLIZAS EN PERIODO DE GRACIA AL 2 DE FEBRERO 2026 MAIL DE AGENTES.xlsx"
archivo_base_clientes_full = "BMI_REPORTE_A_SU_MEDIDA_525_20250923_15_17_15_5555_20250923.xlsx"

df_periodo_gracia = pd.read_excel(archivo_periodo_gracia, sheet_name="Sheet1", engine="openpyxl")
df_base_clientes = pd.read_excel(archivo_base_clientes, sheet_name="Hoja1", engine="openpyxl")
df_base_clientes_full = pd.read_excel(archivo_base_clientes_full, engine="openpyxl")

# Normalize column names (strip whitespace)
df_periodo_gracia = df_periodo_gracia.rename(columns=lambda x: str(x).strip())
df_base_clientes = df_base_clientes.rename(columns=lambda x: str(x).strip())
df_base_clientes_full = df_base_clientes_full.rename(columns=lambda x: str(x).strip())

# ============================================================
# Normalize contract numbers (if columns exist)
# ============================================================

if "Póliza" in df_periodo_gracia.columns:
    df_periodo_gracia["Póliza"] = df_periodo_gracia["Póliza"].apply(clean_poliza)
if "Póliza" in df_base_clientes.columns:
    df_base_clientes["Póliza"] = df_base_clientes["Póliza"].apply(clean_poliza)
if "Póliza/Contrato No." in df_base_clientes_full.columns:
    df_base_clientes_full["Póliza/Contrato No."] = df_base_clientes_full["Póliza/Contrato No."].apply(clean_poliza)

# ============================================================
# Deduplicate by Póliza where applicable
# ============================================================

if "Póliza" in df_periodo_gracia.columns:
    df_periodo_gracia = df_periodo_gracia.drop_duplicates(subset=["Póliza"], keep="first")
if "Póliza" in df_base_clientes.columns:
    df_base_clientes = df_base_clientes.drop_duplicates(subset=["Póliza"], keep="first")

# ============================================================
# Split into unique vs duplicated (based on df_base_clientes)
# ============================================================

if "Póliza" in df_base_clientes.columns and "Póliza" in df_periodo_gracia.columns:
    df_no_duplicados = df_periodo_gracia[~df_periodo_gracia["Póliza"].isin(df_base_clientes["Póliza"])].copy()
    df_duplicados = df_periodo_gracia[df_periodo_gracia["Póliza"].isin(df_base_clientes["Póliza"])].copy()
else:
    df_no_duplicados = df_periodo_gracia.copy()
    df_duplicados = pd.DataFrame(columns=df_periodo_gracia.columns)

print(f"✅ Step 1: {len(df_no_duplicados)} unique, {len(df_duplicados)} duplicated contracts.")

# ============================================================
# Step 2: Split names (Asegurado) safely
# ============================================================

if not df_no_duplicados.empty:
    df_no_duplicados = split_names(df_no_duplicados, source_col="Asegurado", first_col="First Name", last_col="Last Name")
if not df_duplicados.empty:
    df_duplicados = split_names(df_duplicados, source_col="Asegurado", first_col="First Name", last_col="Last Name")

print("✅ Step 2: names extracted.")

# ============================================================
# Step 3: Map emails from full client file and validate
# ============================================================

# Build mapping only if expected columns exist
map_poliza_email = {}
if "Póliza/Contrato No." in df_base_clientes_full.columns and "Correo Electrónic Personal" in df_base_clientes_full.columns:
    map_poliza_email = df_base_clientes_full.set_index("Póliza/Contrato No.")["Correo Electrónic Personal"].to_dict()

def add_email_and_validation(df: pd.DataFrame):
    if df.empty:
        return df
    # Map email (safe: use .map on Póliza if exists)
    if "Póliza" in df.columns:
        # If "Email Cliente" already exists, overwrite with mapped value
        df["Email Cliente"] = df["Póliza"].map(map_poliza_email).where(df["Póliza"].notna(), None)
    else:
        df["Email Cliente"] = None
    df["Valid Email"] = df["Email Cliente"].apply(is_valid_email)
    return df

df_no_duplicados = add_email_and_validation(df_no_duplicados)
df_duplicados = add_email_and_validation(df_duplicados)

# ============================================================
# Step 4: Extract Asesor email and Alias (from df_base_clientes)
# ============================================================

# ============================================================
# Step 4: Extract Asesor email and Alias (robust splitting)
# ============================================================

def normalize_whitespace(s: pd.Series) -> pd.Series:
    """Normalize whitespace: convert non-breaking spaces, tabs; collapse multiple spaces; strip."""
    s = s.fillna("").astype(str)
    s = s.str.replace("\u00A0", " ", regex=False)   # non-breaking space
    s = s.str.replace(r"\s+", " ", regex=True)      # collapse any whitespace sequence to single space
    return s.str.strip()

# Ensure expected columns exist
for col in ["Póliza", "Dirección de correo Asesor", "Alias"]:
    if col not in df_base_clientes.columns:
        df_base_clientes[col] = ""

# Normalize Póliza and dedupe by Póliza (preferred)
# df_base_clientes["Póliza"] = df_base_clientes["Póliza"].apply(clean_poliza)
# if df_base_clientes["Póliza"].replace("", pd.NA).notna().any():
#     df_asesores = df_base_clientes.drop_duplicates(subset=["Póliza"], keep="first").copy()
# else:
#     # fallback: dedupe by asesor email
#     if df_base_clientes["Dirección de correo Asesor"].replace("", pd.NA).notna().any():
#         df_asesores = df_base_clientes.drop_duplicates(subset=["Dirección de correo Asesor"], keep="first").copy()
#     else:
#         df_asesores = df_base_clientes.drop_duplicates(keep="first").copy()

df_asesores = df_base_clientes.drop_duplicates(subset=["Dirección de correo Asesor"], keep="first").copy()
# Normalize Alias whitespace and force string
df_asesores["Alias_clean"] = normalize_whitespace(df_asesores["Alias"])

# Split on first whitespace (keeps multi-word last names)
alias_split = df_asesores["Alias_clean"].str.split(" ", n=2, expand=True)
df_asesores["Alias First Name"] = alias_split[0].str.upper().fillna("")
df_asesores["Alias Last Name"] = alias_split[2].str.upper().fillna("")

# Optional debug: print a sample to verify splitting (remove or comment out in production)
print("Sample Alias splits (first 10 rows):")
print(df_asesores[["Alias", "Alias_clean", "Alias First Name", "Alias Last Name"]].head(10).to_string(index=False))

# Prepare export columns and export
cols_asesores = ["Póliza", "Dirección de correo Asesor", "Alias First Name", "Alias Last Name"]
for c in cols_asesores:
    if c not in df_asesores.columns:
        df_asesores[c] = ""
df_asesores[cols_asesores].to_csv("Final_Asesores.csv", index=False)
print("✅ Final_Asesores.csv exported with Alias First Name and Alias Last Name.")

# ============================================================
# Final exports (only essentials)
# ============================================================

cols_finales = ["Email Cliente", "First Name", "Last Name"]

# Unique clients
if not df_no_duplicados.empty:
    df_no_duplicados[cols_finales].drop_duplicates(subset=["Email Cliente"]).to_csv("Final_Unique.csv", index=False)
else:
    pd.DataFrame(columns=cols_finales).to_csv("Final_Unique.csv", index=False)

# Duplicated clients
if not df_duplicados.empty:
    df_duplicados[cols_finales].drop_duplicates(subset=["Email Cliente"]).to_csv("Final_Duplicated.csv", index=False)
else:
    pd.DataFrame(columns=cols_finales).to_csv("Final_Duplicated.csv", index=False)

# Valid emails only
if not df_no_duplicados.empty:
    df_no_duplicados[df_no_duplicados["Valid Email"]][cols_finales].drop_duplicates(subset=["Email Cliente"]).to_csv("Final_Unique_ValidEmails.csv", index=False)
else:
    pd.DataFrame(columns=cols_finales).to_csv("Final_Unique_ValidEmails.csv", index=False)

if not df_duplicados.empty:
    df_duplicados[df_duplicados["Valid Email"]][cols_finales].drop_duplicates(subset=["Email Cliente"]).to_csv("Final_Duplicated_ValidEmails.csv", index=False)
else:
    pd.DataFrame(columns=cols_finales).to_csv("Final_Duplicated_ValidEmails.csv", index=False)

# Missing emails log (from unique set)
if not df_no_duplicados.empty:
    missing_emails = df_no_duplicados[df_no_duplicados["Email Cliente"].isna()]
    missing_emails.to_excel("Missing_Emails.xlsx", index=False)
else:
    pd.DataFrame().to_excel("Missing_Emails.xlsx", index=False)

print("✅ Step 5: final CSVs exported with deduplication, validation, and missing email log.")

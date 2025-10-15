"""
------------------------------------------------------------
ESCENARIO 2 - MERGE CONSOLIDADO (HV / PUNTOS / PREPAGO)
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
"""

import pandas as pd
from pathlib import Path
import re
import numpy as np
from tkinter import Tk, messagebox

# ------------------------------------------------------------
# 1️⃣ CONFIGURACIÓN INICIAL
# ------------------------------------------------------------
base_path = Path("data_clean")

archivos = {
    "HV": base_path / "HV_limpio.xlsx",
    "PUNTOS": base_path / "PUNTOS_limpio.xlsx",
    "PREPAGO": base_path / "PREPAGO_limpio.xlsx"
}

datasets = []

# ------------------------------------------------------------
# AVISO INFORMATIVO (no elimina filas, solo notifica)
# ------------------------------------------------------------
root = Tk()
root.withdraw()  # Oculta la ventana principal
messagebox.showinfo(
    "Control ANS – MERGE CONSOLIDADO",
    "El sistema conservará todos los registros.\n\n"
    "Los valores vacíos serán reemplazados por indicadores visibles:\n"
    "• 'SIN_PEDIDO' en pedidos vacíos\n"
    "• 'SIN DATO' en texto\n"
    "• '1900-01-01' en fechas\n"
    "• 0 en números\n\n"
    "Recuerda cerrar Excel antes de continuar."
)
root.destroy()


# ------------------------------------------------------------
# 2️⃣ CARGA Y LIMPIEZA DE CADA ARCHIVO
# ------------------------------------------------------------
# ------------------------------------------------------------
# 2️⃣ CARGA Y LIMPIEZA DE CADA ARCHIVO (CON VALIDACIÓN ESTRUCTURAL)
# ------------------------------------------------------------
diagnostico_path = base_path / "Diagnostico_Estructura.txt"
with open(diagnostico_path, "w", encoding="utf-8") as log:
    log.write("------------------------------------------------------------\n")
    log.write("📋 DIAGNÓSTICO DE ESTRUCTURA DE ARCHIVOS LIMPIOS - MERGE ANS\n")
    log.write("------------------------------------------------------------\n")

columnas_requeridas = [
    "PEDIDO", "MUNICIPIO", "FECHA_INGRESO", "FECHA_INICIO_ANS",
    "ZONA", "SECTOR", "DIAS_CUMP", "FECHA_LIMITE",
    "DIAS_TRANSCURRIDOS", "DIAS_RESTANTES", "ESTADO"
]

for nombre, ruta in archivos.items():
    if ruta.exists():
        df = pd.read_excel(ruta)

        # 🔹 Registrar estructura detectada
        with open(diagnostico_path, "a", encoding="utf-8") as log:
            log.write(f"\n🗂️ Archivo: {ruta.name}\n")
            log.write(f"Total columnas detectadas: {len(df.columns)}\n")
            for col in df.columns:
                log.write(f"   - {col}\n")

        # 🔹 Normalizar encabezados
        df.columns = df.columns.str.strip().str.upper()

        # 🔹 Validar columnas requeridas
        faltantes = [col for col in columnas_requeridas if col not in df.columns]
        if faltantes:
            mensaje = (
                f"❌ ERROR: Faltan columnas obligatorias en el archivo {ruta.name}.\n\n"
                f"Columnas faltantes: {', '.join(faltantes)}\n\n"
                f"Revisa el formato antes de continuar.\n"
                f"El proceso se detuvo para evitar errores en Power BI."
            )
            print(mensaje)
            root = Tk()
            root.withdraw()
            messagebox.showerror("Error de estructura", mensaje)
            root.destroy()
            raise SystemExit("❌ Estructura inválida. Proceso detenido.")

        # 🔹 Eliminar columna duplicada "PEDIDO" si existe más de una
        columnas_pedido = [c for c in df.columns if "PEDIDO" in c]
        if len(columnas_pedido) > 1:
            print(f"⚠️ {nombre}: se detectó columna PEDIDO duplicada, se eliminará la segunda.")
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.loc[:, ~df.columns.str.contains("PEDIDO.1", regex=True)]

        # 🔹 Asegurar columnas consistentes
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = pd.NA

        # 🔹 Reordenar y agregar tipo
        df = df[columnas_requeridas]
        df["TIPO_DATASET"] = nombre

        datasets.append(df)
        print(f"✅ {nombre} cargado correctamente ({len(df)} registros)")
    else:
        print(f"⚠️ No se encontró el archivo: {ruta}")

with open(diagnostico_path, "a", encoding="utf-8") as log:
    log.write("\n✅ Validación completada. No se detectaron errores estructurales.\n")
# ------------------------------------------------------------
# 3️⃣ MERGE FINAL ORIGINAL
# ------------------------------------------------------------
if not datasets:
    raise SystemExit("❌ No se encontró ningún archivo limpio para consolidar.")

consolidado = pd.concat(datasets, ignore_index=True)
print(f"\n🔹 Registros combinados totales: {len(consolidado)}")

# ⚠️ NO eliminar vacíos ni duplicados aquí (archivo original)
salida_total = base_path / "MERGE_ANS.xlsx"
consolidado.to_excel(salida_total, index=False)
print(f"\n📦 Archivo consolidado original generado (intacto): {salida_total}")

# ------------------------------------------------------------
# 4️⃣ LIMPIEZA Y NORMALIZACIÓN DE TEXTO (FINAL)
# ------------------------------------------------------------
for col in ["MUNICIPIO", "SECTOR", "ESTADO"]:
    consolidado[col] = (
        consolidado[col]
        .astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"–", "-", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"Á", "A", regex=True)
        .str.replace(r"É", "E", regex=True)
        .str.replace(r"Í", "I", regex=True)
        .str.replace(r"Ó", "O", regex=True)
        .str.replace(r"Ú", "U", regex=True)
    )

# 🔹 LIMPIEZA DE SECTOR
reemplazos_sector = {
    "OCCIDENETE": "OCCIDENTE",
    "OCCIDENTE - OLAYA": "OCCIDENTE",
    "OCCIDENTE - SAN CRISTOBAL": "OCCIDENTE",
    "OCCIDENTE – AGUAS FRIAS": "OCCIDENTE",
    "OCCIDENTE-AGUAS FRIAS": "OCCIDENTE",
    "SUR - SABANETA": "SUR",
    "SUR - AGIZAL": "SUR",
    "SUR-S.PRADO": "SUR",
    "SUR-SABANETA": "SUR",
    "SUR ITAGUI AGIZAL": "SUR",
    "SUR-LIMONAR": "SUR",
    "ENV": "SUR",
    "REPL": np.nan,
    "NORTE": np.nan
}
consolidado["SECTOR"] = consolidado["SECTOR"].replace(reemplazos_sector)

def limpiar_sector(valor):
    if pd.isna(valor):
        return np.nan
    if re.match(r"^OCCIDENTE", valor):
        return "OCCIDENTE"
    elif re.match(r"^SUR", valor):
        return "SUR"
    elif re.match(r"^ORIENTE", valor):
        return "ORIENTE"
    elif re.match(r"^LA ESTRELLA", valor):
        return "SUR"
    else:
        return valor

consolidado["SECTOR"] = consolidado["SECTOR"].apply(limpiar_sector)
consolidado["SECTOR"] = consolidado["SECTOR"].replace(["NAN", "NONE", "NULL", "PD.NA", "SIN DATO"], np.nan)

# 🔹 LIMPIEZA DE MUNICIPIO
reemplazos_municipio = {
    "MED": "MEDELLIN",
    "MEDLLIN": "MEDELLIN",
    "ENVIG": "ENVIGADO",
    "ENV": "ENVIGADO",
    "SAB": "SABANETA",
    "LA ESTRELLA": "LA ESTRELLA",
    "CAL": "CALDAS",
    "GUAR": "GUARNE",
    "ESTR": "LA ESTRELLA",
    "EL RET": "EL RETIRO",
    "ENVI": "ENVIGADO"
}
consolidado["MUNICIPIO"] = consolidado["MUNICIPIO"].replace(reemplazos_municipio)

# 🔹 LIMPIEZA DE ESTADO
reemplazos_estado = {
    "VENC": "VENCIDO",
    "VENCID": "VENCIDO",
    "CUMP": "CUMPLIDO",
    "ALER": "ALERTA",
    "SINFECHA": "SIN FECHA",
    "SIN FECHAS": "SIN FECHA"
}
consolidado["ESTADO"] = consolidado["ESTADO"].replace(reemplazos_estado)

# ------------------------------------------------------------
# 5️⃣ DIAGNÓSTICO DE CALIDAD
# ------------------------------------------------------------
print("\n📊 Diagnóstico de valores únicos después de limpieza:")
print("SECTOR     →", consolidado["SECTOR"].dropna().unique())
print("MUNICIPIO  →", consolidado["MUNICIPIO"].dropna().unique()[:10])
print("ESTADO     →", consolidado["ESTADO"].dropna().unique())

print("\n📈 Conteo por SECTOR y ESTADO:")
print(consolidado.groupby(["SECTOR", "ESTADO"]).size())

# ------------------------------------------------------------
# 6️⃣ MERGE PARA POWER BI (CONSERVA TODO Y RELLENA VACÍOS)
# ------------------------------------------------------------
print("\n🔍 Preparando MERGE_ANS_FINAL con valores visibles para análisis en Power BI...")

merge_para_powerbi = consolidado.copy()

# 🔹 Reemplazar valores vacíos por indicadores visibles
# 🔹 Reemplazar valores vacíos por indicadores visibles
for col in merge_para_powerbi.columns:
    # 🔸 Primero, limpiar espacios y valores tipo "nan", "None", etc.
    merge_para_powerbi[col] = merge_para_powerbi[col].replace(
        ["", " ", "NaN", "NAN", "None", None, np.nan], pd.NA
    )

    # 🔸 Ahora aplicar los reemplazos según tipo de dato
    if col == "PEDIDO":
        merge_para_powerbi[col] = merge_para_powerbi[col].fillna("SIN_PEDIDO")
    elif merge_para_powerbi[col].dtype in ["float64", "int64"]:
        merge_para_powerbi[col] = merge_para_powerbi[col].fillna(0)
    elif "FECHA" in col:
        merge_para_powerbi[col] = merge_para_powerbi[col].fillna("1900-01-01")
    else:
        merge_para_powerbi[col] = merge_para_powerbi[col].fillna("SIN DATO")

# 🔹 Crear hoja resumen de vacíos originales (antes del reemplazo)
vacios_por_columna = consolidado.isna().sum()
porcentaje_vacios = (vacios_por_columna / len(consolidado) * 100).round(2)

df_alerta = pd.DataFrame({
    "Columna": vacios_por_columna.index,
    "Valores_Vacíos": vacios_por_columna.values,
    "Porcentaje (%)": porcentaje_vacios.values
}).sort_values(by="Valores_Vacíos", ascending=False)

# 🔹 Generar hoja con resumen general
df_resumen = pd.DataFrame({
    "Métrica": [
        "Total de registros",
        "Filas duplicadas",
        "Filas completamente vacías"
    ],
    "Valor": [
        len(consolidado),
        consolidado.duplicated().sum(),
        consolidado.isna().all(axis=1).sum()
    ]
})
# 🔹 Guardar en un solo archivo Excel con múltiples hojas
salida_powerbi = base_path / "MERGE_ANS_FINAL.xlsx"

# 🟢 Aviso preventivo antes de exportar
from tkinter import messagebox
messagebox.showinfo(
    "Control ANS - MERGE CONSOLIDADO",
    "Antes de continuar, asegúrate de que el archivo MERGE_ANS_FINAL.xlsx esté cerrado.\n\n"
    "Esto evitará errores de permiso durante la exportación."
)

try:
    with pd.ExcelWriter(salida_powerbi, engine="openpyxl") as writer:
        merge_para_powerbi.to_excel(writer, sheet_name="MERGE_ANS_FINAL", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen_General", index=False)
        df_alerta.to_excel(writer, sheet_name="ALERTA_DATOS_VACIOS", index=False)

    print(f"📁 Archivo listo para Power BI: {salida_powerbi}")
    print(f"📊 Incluye hojas 'MERGE_ANS_FINAL', 'Resumen_General' y 'ALERTA_DATOS_VACIOS'.")
    print("✅ Todos los vacíos fueron reemplazados con valores visibles (SIN DATO, 0, 1900-01-01).")

except PermissionError:
    messagebox.showerror(
        "Error de Permiso",
        "No se pudo generar el archivo MERGE_ANS_FINAL.xlsx.\n\n"
        "Causa: El archivo está abierto en Excel o Power BI.\n\n"
        "Cierra el archivo y vuelve a ejecutar el proceso."
    )
    print("❌ Error: El archivo MERGE_ANS_FINAL.xlsx está abierto. Cierra Excel y vuelve a ejecutar.")

# ------------------------------------------------------------
# 7️⃣ (OPCIONAL) ARCHIVOS SEPARADOS POR TIPO
# ------------------------------------------------------------
for tipo in consolidado["TIPO_DATASET"].unique():
    df_tipo = consolidado[consolidado["TIPO_DATASET"] == tipo]
    salida_tipo = base_path / f"{tipo}_filtrado.xlsx"
    df_tipo.to_excel(salida_tipo, index=False)
    print(f"🗂️ Archivo separado generado: {salida_tipo}")

print("\n✅ Consolidación completada con éxito y datos estandarizados.")

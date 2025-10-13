"""
------------------------------------------------------------
ESCENARIO 2 - MERGE CONSOLIDADO (HV / PUNTOS / PREPAGO)
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
"""

import pandas as pd
from pathlib import Path

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
# 2️⃣ CARGA Y LIMPIEZA DE CADA ARCHIVO
# ------------------------------------------------------------
for nombre, ruta in archivos.items():
    if ruta.exists():
        df = pd.read_excel(ruta)

        # 🔹 Normalizar encabezados
        df.columns = df.columns.str.strip().str.upper()

        # 🔹 Eliminar columna duplicada "PEDIDO" si existe más de una
        columnas_pedido = [c for c in df.columns if "PEDIDO" in c]
        if len(columnas_pedido) > 1:
            print(f"⚠️ {nombre}: se detectó columna PEDIDO duplicada, se eliminará la segunda.")
            df = df.loc[:, ~df.columns.duplicated()]  # elimina duplicados
            df = df.loc[:, ~df.columns.str.contains("PEDIDO.1", regex=True)]  # por si Excel la nombra "PEDIDO.1"

        # 🔹 Asegurar columnas consistentes con las otras bases
        columnas_base = [
            "PEDIDO", "MUNICIPIO", "FECHA_INGRESO", "FECHA_INICIO_ANS",
            "ZONA", "SECTOR", "DIAS_CUMP", "FECHA_LIMITE",
            "DIAS_TRANSCURRIDOS", "DIAS_RESTANTES", "ESTADO"
        ]
        for col in columnas_base:
            if col not in df.columns:
                df[col] = pd.NA

        # 🔹 Reordenar columnas y agregar tipo
        df = df[columnas_base]
        df["TIPO_DATASET"] = nombre

        datasets.append(df)
        print(f"✅ {nombre} cargado ({len(df)} registros)")
    else:
        print(f"⚠️ No se encontró el archivo: {ruta}")

# ------------------------------------------------------------
# 3️⃣ MERGE FINAL
# ------------------------------------------------------------
if not datasets:
    raise SystemExit("❌ No se encontró ningún archivo limpio para consolidar.")

consolidado = pd.concat(datasets, ignore_index=True)

# ------------------------------------------------------------
# 4️⃣ EXPORTAR RESULTADOS
# ------------------------------------------------------------
salida_total = base_path / "MERGE_ANS.xlsx"
consolidado.to_excel(salida_total, index=False)
print(f"\n📦 Archivo consolidado generado: {salida_total}")

# ------------------------------------------------------------
# 5️⃣ (OPCIONAL) ARCHIVOS SEPARADOS POR TIPO
# ------------------------------------------------------------
for tipo in consolidado["TIPO_DATASET"].unique():
    df_tipo = consolidado[consolidado["TIPO_DATASET"] == tipo]
    salida_tipo = base_path / f"{tipo}_filtrado.xlsx"
    df_tipo.to_excel(salida_tipo, index=False)
    print(f"🗂️ Archivo separado generado: {salida_tipo}")

print("\n✅ Consolidación completada con éxito sin columnas duplicadas.")

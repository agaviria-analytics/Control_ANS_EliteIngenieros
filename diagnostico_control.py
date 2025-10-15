"""
------------------------------------------------------------
CONTROL DE VACÍOS – Proyecto ANS
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

# ------------------------------------------------------------
# CONFIGURACIÓN INICIAL
# ------------------------------------------------------------
base_path = Path("data_clean")
ruta_merge = base_path / "MERGE_ANS_FINAL.xlsx"   # ✅ ahora lee el archivo final, no el original
salida_informe = base_path / "Pedidos_incompletos.xlsx"

# ------------------------------------------------------------
# CREAR VENTANA OCULTA
# ------------------------------------------------------------
root = tk.Tk()
root.withdraw()

# ------------------------------------------------------------
# CARGA DE DATOS
# ------------------------------------------------------------
if not ruta_merge.exists():
    messagebox.showerror("Error", "No se encontró el archivo MERGE_ANS_FINAL.xlsx.\nEjecuta primero el merge consolidado.")
    raise SystemExit("❌ No se encontró el archivo MERGE_ANS_FINAL.xlsx.")

df = pd.read_excel(ruta_merge)
print(f"📂 Archivo cargado correctamente: {len(df)} registros")

# ------------------------------------------------------------
# DETECCIÓN DE CELDAS VACÍAS
# ------------------------------------------------------------
df_incompletos = df[df.isin(["SIN DATO", "NAN", "", None]).any(axis=1)]
resumen = df.isin(["SIN DATO", "NAN", "", None]).sum().reset_index()
resumen.columns = ["Columna", "Valores_Vacíos"]

# ------------------------------------------------------------
# EXPORTACIÓN DEL INFORME
# ------------------------------------------------------------
with pd.ExcelWriter(salida_informe, engine="openpyxl") as writer:
    df_incompletos.to_excel(writer, sheet_name="Pedidos Incompletos", index=False)
    resumen.to_excel(writer, sheet_name="Resumen Vacíos", index=False)

print(f"\n📊 Informe generado: {salida_informe}")
print(f"🚨 Pedidos con información incompleta: {len(df_incompletos)}")

# ------------------------------------------------------------
# MENSAJE FINAL
# ------------------------------------------------------------
messagebox.showinfo(
    "Control de Vacíos",
    f"Se detectaron {len(df_incompletos)} pedidos con datos incompletos.\n\n"
    f"✅ MERGE_ANS_FINAL.xlsx se mantuvo intacto.\n"
    f"📄 Informe exportado: {salida_informe.name}\n\n"
    f"Envía este archivo al usuario responsable para revisión de datos."
)

print("\n✅ CONTROL DE VACÍOS completado sin modificar el MERGE_ANS_FINAL.")
root.destroy()

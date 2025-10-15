"""
------------------------------------------------------------
ESCENARIO 1 - Limpieza individual (HV / PUNTOS / PREPAGO)
------------------------------------------------------------
Autor: Héctor + IA (2025)
------------------------------------------------------------
"""

import argparse
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path

# ------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ------------------------------------------------------------
WEEKMASK = "1111100"  # Lunes a viernes hábiles

# Lista de días festivos nacionales (puedes ampliarla con los que maneja tu empresa o EPM)
# ------------------------------------------------------------
# LISTA DE FESTIVOS NACIONALES (2025–2026)
# ------------------------------------------------------------
FESTIVOS = [
    # 2025
    "2025-01-01",  # Año Nuevo
    "2025-01-06",  # Reyes Magos
    "2025-03-24",  # San José (trasladado)
    "2025-04-17",  # Jueves Santo
    "2025-04-18",  # Viernes Santo
    "2025-05-01",  # Día del Trabajo
    "2025-05-26",  # Ascensión del Señor (trasladado)
    "2025-06-16",  # Corpus Christi (trasladado)
    "2025-06-23",  # Sagrado Corazón (trasladado)
    "2025-07-07",  # San Pedro y San Pablo (trasladado)
    "2025-08-07",  # Batalla de Boyacá
    "2025-08-18",  # Asunción de la Virgen (trasladado)
    "2025-10-13",  # Día de la Raza (trasladado)
    "2025-11-03",  # Todos los Santos (trasladado)
    "2025-11-17",  # Independencia de Cartagena (trasladado)
    "2025-12-08",  # Inmaculada Concepción
    "2025-12-25",  # Navidad

    # 2026
    "2026-01-01",  # Año Nuevo
    "2026-01-12",  # Reyes Magos (trasladado)
    "2026-03-23",  # San José (trasladado)
    "2026-04-02",  # Jueves Santo
    "2026-04-03",  # Viernes Santo
    "2026-05-01",  # Día del Trabajo
    "2026-05-18",  # Ascensión del Señor (trasladado)
    "2026-06-08",  # Corpus Christi (trasladado)
    "2026-06-15",  # Sagrado Corazón (trasladado)
    "2026-06-29",  # San Pedro y San Pablo (trasladado)
    "2026-07-20",  # Independencia de Colombia
    "2026-08-07",  # Batalla de Boyacá
    "2026-08-17",  # Asunción de la Virgen (trasladado)
    "2026-10-12",  # Día de la Raza
    "2026-11-02",  # Todos los Santos (trasladado)
    "2026-11-16",  # Independencia de Cartagena (trasladado)
    "2026-12-08",  # Inmaculada Concepción
    "2026-12-25",  # Navidad
]

# Convertimos a tipo fecha (datetime64) para que Numpy las use correctamente
FESTIVOS = np.array(FESTIVOS, dtype='datetime64[D]')

# Convertimos a tipo fecha (datetime64) para que Numpy las use correctamente
FESTIVOS = np.array(FESTIVOS, dtype='datetime64[D]')

ALERTA_UMBRAL_DIAS = 2  # Alerta si faltan 2 días o menos

DIAS_PACTADOS_CONFIG = {
    "HV": {"URBANA": 5, "URBANOS": 5, "URBANO": 5, "URBAN": 5, "RURAL": 8, "RURALES": 8},
    "PUNTOS": {"URBANA": 4, "RURAL": 4},
    "PREPAGO": {"URBANA": 5, "RURAL": 8},
}

# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------
def to_datetime(x):
    """Convierte texto o números a datetime (mantiene hora:minuto)."""
    if pd.isna(x):
        return pd.NaT
    try:
        return pd.to_datetime(str(x), dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def detectar_columna_ru(df):
    """Detecta automáticamente la columna R/U o ZONA."""
    for col in df.columns:
        col_clean = col.replace(" ", "").replace("/", "").replace("\\", "").upper()
        if col_clean in ["RU", "R/U", "R∕U", "R_U"]:
            return col
    return None


def add_business_days_keep_time(dt, n_days):
    """Suma n días hábiles (lunes-viernes) sin contar festivos."""
    if pd.isna(dt):
        return pd.NaT
    date_part = np.datetime64(dt.date())
    time_part = dt.time()
    new_date = np.busday_offset(
        date_part,
        n_days,
        roll='forward',
        weekmask=WEEKMASK,
        holidays=FESTIVOS
    )
    return datetime.combine(pd.to_datetime(str(new_date)).date(), time_part)


def business_days_between(start_dt, end_dt):
    """Cuenta días hábiles (lunes-viernes) sin incluir festivos."""
    if pd.isna(start_dt) or pd.isna(end_dt):
        return np.nan
    start_date = np.datetime64(start_dt.date() + timedelta(days=1))
    end_date = np.datetime64(end_dt.date())
    return int(np.busday_count(start_date, end_date, weekmask=WEEKMASK, holidays=FESTIVOS))


# ------------------------------------------------------------
# LIMPIEZA PRINCIPAL
# ------------------------------------------------------------
def limpiar_individual(df, dataset):
    # ------------------------------------------------------------
    # LIMPIEZA DE NOMBRES DE COLUMNAS
    # ------------------------------------------------------------
    df.columns = (
        df.columns.str.strip()
        .str.upper()
        .str.replace(r"[^A-Z0-9_/ ]", "", regex=True)
        .str.replace("  ", " ")
    )

    # 🔍 Elimina columnas duplicadas invisibles
    clean_cols = []
    final_cols = []
    for col in df.columns:
        base = col.strip().upper()
        if base not in clean_cols:
            clean_cols.append(base)
            final_cols.append(col)
    df = df.loc[:, final_cols]

    # ------------------------------------------------------------
    # DETECTAR Y RENOMBRAR COLUMNAS CLAVE
    # ------------------------------------------------------------
    renombres = {}
    for c in df.columns:
        col_clean = c.strip().upper().replace("_", "").replace("-", "").replace(" ", "")
        if "PEDIDO" in col_clean:
            renombres[c] = "PEDIDO"
        elif "MPIO" in col_clean or "MUNICIPIO" in col_clean:
            renombres[c] = "MUNICIPIO"
        elif "FECHAINGRESO" in col_clean:
            renombres[c] = "FECHA_INGRESO"
        elif "FECHAINICIO" in col_clean and "ANS" in col_clean:
            renombres[c] = "FECHA_INICIO_ANS"
        elif "ACTIVIDAD" in col_clean:
            renombres[c] = "ACTIVIDAD"
        elif "SECTOR" in col_clean:
            renombres[c] = "SECTOR"
        elif col_clean in ["EST", "ESTADO", "STATUS"]:
            renombres[c] = "ESTADO_DIGITADO"    

    df.rename(columns=renombres, inplace=True)

    # ------------------------------------------------------------
    # ASEGURAR COLUMNAS CLAVE
    # ------------------------------------------------------------
    for col in ["PEDIDO", "MUNICIPIO", "FECHA_INGRESO", "FECHA_INICIO_ANS"]:
        if col not in df.columns:
            df[col] = pd.NA

    # Detectar columna ZONA o R/U
    col_zona = detectar_columna_ru(df)
    if col_zona:
        df.rename(columns={col_zona: "ZONA"}, inplace=True)
    elif "ZONA" not in df.columns:
        df["ZONA"] = np.nan

    if "SECTOR" not in df.columns:
        df["SECTOR"] = np.nan

    # ------------------------------------------------------------
    # NORMALIZAR MUNICIPIO Y ZONA
    # ------------------------------------------------------------
    MUNICIPIOS_MAP = {
        "MED": "MEDELLÍN",
        "ITA": "ITAGÜÍ",
        "LA EST": "LA ESTRELLA",
        "SAB": "SABANETA",
        "ENV": "ENVIGADO",
    }

    df["MUNICIPIO"] = (
        df["MUNICIPIO"]
        .astype(str)
        .str.upper()
        .str.strip()
        .replace(MUNICIPIOS_MAP)
    )

    df["ZONA"] = (
        df["ZONA"]
        .astype(str)
        .str.upper()
        .str.strip()
        .str.replace("-", "")
        .str.replace("_", "")
    )
    # ------------------------------------------------------------
# 🔹 LIMPIEZA DE ZONA (URBANA / RURAL / SIN DATO)
# ------------------------------------------------------------
    reemplazos_zona = {
        "URBAN": "URBANA",
        "URBANO": "URBANA",
        "URBANA": "URBANA",
        "RURAL": "RURAL",
        "RURALES": "RURAL",
        "SIN DATO": "SIN DATO",
        "NAN": "SIN DATO",
        "": "SIN DATO",
        "NONE": "SIN DATO"
    }

    df["ZONA"] = df["ZONA"].replace(reemplazos_zona)
    print(f"📍 Limpieza aplicada a columna ZONA en dataset {dataset.upper()}.")
    print("Valores únicos después de limpieza:", df["ZONA"].dropna().unique())


    # Convertir fechas
    df["FECHA_INGRESO"] = df["FECHA_INGRESO"].apply(to_datetime)
    df["FECHA_INICIO_ANS"] = df["FECHA_INICIO_ANS"].apply(to_datetime)

    # ------------------------------------------------------------
    # DÍAS PACTADOS POR ZONA O ACTIVIDAD
    # ------------------------------------------------------------
    if dataset.upper() == "PREPAGO":
        def dias_cumplimiento(row):
            act = str(row.get("ACTIVIDAD", "")).upper()
            zona = str(row.get("ZONA", "")).upper()
            if any(x in act for x in ["DESINSTALAR", "INSTALAR", "TRABAJO"]):
                return 11
            elif "REPLANTEO" in act and "URBANA" in zona:
                return 5
            elif "REPLANTEO" in act and "RURAL" in zona:
                return 8
            else:
                return 0

        df["DIAS_CUMP"] = df.apply(dias_cumplimiento, axis=1)

    else:
        dias_cfg = DIAS_PACTADOS_CONFIG.get(dataset.upper(), DIAS_PACTADOS_CONFIG["HV"])
        df["DIAS_CUMP"] = df["ZONA"].map(dias_cfg).fillna(0).astype(int)

    # ------------------------------------------------------------
    # CÁLCULOS EXACTOS
    # ------------------------------------------------------------
    hoy = datetime.now()

    # Fecha límite considerando festivos
    df["FECHA_LIMITE"] = df.apply(
        lambda r: add_business_days_keep_time(r["FECHA_INICIO_ANS"], r["DIAS_CUMP"]),
        axis=1
    )

    # Días transcurridos hábiles (sin festivos)
    df["DIAS_TRANSCURRIDOS"] = df.apply(
        lambda r: business_days_between(r["FECHA_INICIO_ANS"], hoy)
        if pd.notna(r["FECHA_INICIO_ANS"]) else np.nan,
        axis=1,
    )

    # Tiempo restante y estado
    def tiempo_restante(row):
        if pd.isna(row["FECHA_LIMITE"]):
            return ""
        if row["FECHA_LIMITE"] < hoy:
            return "VENCIDO"
        return row["FECHA_LIMITE"] - hoy

    df["TIEMPO_RESTANTE"] = df.apply(tiempo_restante, axis=1)

    def formatear_dias_restantes(row):
        valor = row["TIEMPO_RESTANTE"]
        fecha_inicio = row["FECHA_INICIO_ANS"]
        if isinstance(valor, str) and valor == "VENCIDO":
            return "VENCIDO"
        if isinstance(valor, timedelta):
            if valor.total_seconds() <= 0:
                return "VENCIDO"
            dias = int(valor.days)
            hora_inicio = fecha_inicio.time()
            return f"{dias} días {hora_inicio.hour:02d}:{hora_inicio.minute:02d}"
        return ""

    df["DIAS_RESTANTES"] = df.apply(formatear_dias_restantes, axis=1)

    def calcular_estado(row):
        valor = row["TIEMPO_RESTANTE"]
        if isinstance(valor, str) and valor == "VENCIDO":
            return "VENCIDO"
        if isinstance(valor, timedelta):
            dias_rest = valor.days
            if dias_rest <= ALERTA_UMBRAL_DIAS:
                return "ALERTA"
            return "A TIEMPO"
        return "SIN FECHA"

    df["ESTADO"] = df.apply(calcular_estado, axis=1)

    # ------------------------------------------------------------
    # SALIDA FINAL
    # ------------------------------------------------------------
    columnas_finales = [
        "PEDIDO", "MUNICIPIO", "FECHA_INGRESO", "FECHA_INICIO_ANS",
        "ZONA", "SECTOR", "DIAS_CUMP", "FECHA_LIMITE",
        "DIAS_TRANSCURRIDOS", "DIAS_RESTANTES", "ESTADO", "ESTADO_DIGITADO"
    ]

    df_final = df[columnas_finales].copy()

    print("\n🧩 Vista previa:")
    print(df_final.head(10))

    return df_final


# ------------------------------------------------------------
# EJECUCIÓN Y EXPORTACIÓN
# ------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--dataset", required=True, choices=["HV", "PUNTOS", "PREPAGO"])
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    src = Path(args.input)
    if not src.exists():
        raise SystemExit(f"No se encontró el archivo: {src}")

    print(f"\n📥 Leyendo archivo: {src}")
    df = pd.read_excel(src)

    print("\n🧹 Ejecutando limpieza y cálculos...")
    limpio = limpiar_individual(df, args.dataset)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------
    # ⚙️ Diagnóstico opcional de duplicados
    # ------------------------------------------------------------
    duplicados = limpio[limpio.duplicated(subset=["PEDIDO"], keep=False)]
    if not duplicados.empty:
        print(f"\n⚠️ Se detectaron {len(duplicados)} pedidos duplicados. Se marcarán en una hoja aparte.")
        duplicados["OBSERVACION"] = "DUPLICADO DETECTADO"
        out_duplicados = out.parent / f"Duplicados_{out.stem}.xlsx"
        duplicados.to_excel(out_duplicados, index=False)
        print(f"📄 Archivo de duplicados generado: {out_duplicados}")
    else:
        print("\n✅ No se encontraron duplicados en la columna PEDIDO.")

    # ------------------------------------------------------------
    # 📤 EXPORTACIÓN FINAL
    # ------------------------------------------------------------
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        limpio.to_excel(writer, index=False, sheet_name="DATA_CLEAN")
        resumen = limpio["ESTADO"].value_counts().reset_index()
        resumen.columns = ["ESTADO", "CANTIDAD"]
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")

    print(f"\n✅ Limpieza completada con éxito: {out}")
    print("📊 Hoja 'RESUMEN' generada con totales por estado.")


if __name__ == "__main__":
    main()

"""
------------------------------------------------------------
CONTROL ANS – Panel Empresarial ELITE Ingenieros S.A.S.
------------------------------------------------------------
Autor: Héctor A. Gaviria + IA (2025)
------------------------------------------------------------
"""

import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext
from PIL import Image, ImageTk
import sys
import io

# ------------------------------------------------------------
# CONFIGURACIÓN UTF-8 GLOBAL
# ------------------------------------------------------------
if sys.stdout.encoding is None or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
if sys.stderr.encoding is None or sys.stderr.encoding.lower() != "utf-8":
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

# ------------------------------------------------------------
# RUTAS DE ARCHIVOS Y LOGO
# ------------------------------------------------------------
RUTA_LOGO = r"data_raw/elite.png"
RUTA_HV = r"data_raw/PROGRAMACION HV SUR 2025.xlsx"
RUTA_PUNTOS = r"data_raw/PROGRAMACION PUNTOS DE CONEXION.xlsx"
RUTA_PREPAGO = r"data_raw/PROGRAMACION PREPAGO SUR 2025.xlsx"

# ------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# ------------------------------------------------------------
def ejecutar_comando(nombre, comando):
    """Ejecuta un script y muestra salida en el log."""
    def tarea():
        log_text.insert(tk.END, f"\n🚀 Iniciando {nombre}...\n", "info")
        log_text.see(tk.END)
        boton_estado(False)
        barra_progreso.start()

        try:
            proceso = subprocess.Popen(
                comando,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True,
                cwd=os.path.dirname(os.path.abspath(__file__)),
                encoding="utf-8"
            )

            for linea in proceso.stdout:
                log_text.insert(tk.END, linea)
                log_text.see(tk.END)

            error_salida = proceso.stderr.read()
            proceso.wait()

            if proceso.returncode == 0:
                log_text.insert(tk.END, f"\n✅ {nombre} completado con éxito.\n", "success")
            else:
                log_text.insert(tk.END, f"\n❌ Error al ejecutar {nombre}:\n{error_salida}\n", "error")

        except Exception as e:
            log_text.insert(tk.END, f"\n⚠️ Error inesperado: {e}\n", "error")
        finally:
            barra_progreso.stop()
            boton_estado(True)
            log_text.insert(tk.END, "-" * 60 + "\n", "separador")
            log_text.see(tk.END)

    threading.Thread(target=tarea).start()

# ------------------------------------------------------------
# COMANDOS DE LOS BOTONES
# ------------------------------------------------------------
def ejecutar_hv():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_HV}" --dataset HV --output "data_clean/HV_limpio.xlsx"'
    ejecutar_comando("HV", comando)

def ejecutar_puntos():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_PUNTOS}" --dataset PUNTOS --output "data_clean/PUNTOS_limpio.xlsx"'
    ejecutar_comando("PUNTOS", comando)

def ejecutar_prepago():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_PREPAGO}" --dataset PREPAGO --output "data_clean/PREPAGO_limpio.xlsx"'
    ejecutar_comando("PREPAGO", comando)

def ejecutar_merge():
    comando = 'python -X utf8 merge_escenario2.py'
    ejecutar_comando("MERGE", comando)

# ------------------------------------------------------------
# INTERFAZ GRÁFICA
# ------------------------------------------------------------
ventana = tk.Tk()
ventana.title("Control ANS - ELITE Ingenieros S.A.S.")
ventana.config(bg="#EAEDED")

# 🔹 Tamaño fijo y centrado
ancho, alto = 780, 640
x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
y = (ventana.winfo_screenheight() // 2) - (alto // 2)
ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
ventana.resizable(False, False)

# ------------------------------------------------------------
# ENCABEZADO PROFESIONAL CON DOS LÍNEAS
# ------------------------------------------------------------
frame_banner = tk.Frame(ventana, bg="#EAEDED", height=120)
frame_banner.pack(fill="x")

# 🔸 Fila superior: logo + nombre empresa
frame_superior = tk.Frame(frame_banner, bg="#EAEDED")
frame_superior.pack(pady=(10, 0))

try:
    logo_img = Image.open(RUTA_LOGO)
    logo_img = logo_img.resize((70, 70), Image.Resampling.LANCZOS)
    logo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(frame_superior, image=logo, bg="#EAEDED")
    logo_label.pack(side="left", padx=15)
except Exception:
    logo_label = tk.Label(frame_superior, text="[Logo no encontrado]", fg="black", bg="#EAEDED", font=("Segoe UI", 10))
    logo_label.pack(side="left", padx=15)

# 🔹 Texto bicolor: ELITE (negro) + Ingenieros S.A.S. (verde)
elite_label = tk.Label(
    frame_superior,
    text="ELITE ",
    font=("Segoe UI", 18, "bold"),
    fg="black",
    bg="#EAEDED"
)
elite_label.pack(side="left")

ingenieros_label = tk.Label(
    frame_superior,
    text="Ingenieros S.A.S.",
    font=("Segoe UI", 18, "bold"),
    fg="#1E8449",
    bg="#EAEDED"
)
ingenieros_label.pack(side="left")

# 🔸 Fila inferior centrada: Control ANS
titulo_control = tk.Label(
    frame_banner,
    text="Control ANS",
    font=("Segoe UI", 14, "bold"),
    fg="#1B263B",
    bg="#EAEDED"
)
titulo_control.pack(pady=(0, 10))

# ------------------------------------------------------------
# FRAME DE BOTONES
# ------------------------------------------------------------
frame_botones = tk.Frame(ventana, bg="#EAEDED")
frame_botones.pack(pady=10)

botones = [
    ("HV", ejecutar_hv),
    ("PUNTOS DE CONEXIÓN", ejecutar_puntos),
    ("PREPAGO", ejecutar_prepago),
    ("GENERAR MERGE", ejecutar_merge)
]

botones_widgets = []
for texto, comando in botones:
    b = tk.Button(
        frame_botones,
        text=texto,
        command=comando,
        width=20,
        height=2,
        bg="#1E8449",
        fg="white",
        font=("Segoe UI", 10, "bold"),
        relief="ridge",
        borderwidth=3,
        cursor="hand2",
        activebackground="#229954",
        activeforeground="white"
    )
    b.pack(side="left", padx=8)
    botones_widgets.append(b)

# ------------------------------------------------------------
# BARRA DE PROGRESO
# ------------------------------------------------------------
barra_progreso = ttk.Progressbar(
    ventana,
    orient="horizontal",
    mode="indeterminate",
    length=450
)
barra_progreso.pack(pady=10)

# ------------------------------------------------------------
# CUADRO DE LOG
# ------------------------------------------------------------
log_text = scrolledtext.ScrolledText(
    ventana,
    width=95,
    height=20,
    bg="white",
    font=("Consolas", 9)
)
log_text.pack(pady=5)

log_text.tag_config("info", foreground="#2980B9")
log_text.tag_config("success", foreground="#27AE60")
log_text.tag_config("error", foreground="#C0392B")
log_text.tag_config("separador", foreground="#95A5A6")

# ------------------------------------------------------------
# CONTROL DE BOTONES
# ------------------------------------------------------------
def boton_estado(habilitar=True):
    for b in botones_widgets:
        if habilitar:
            b.config(state="normal", bg="#1E8449", fg="white")
        else:
            b.config(state="disabled", disabledforeground="white", bg="#1E8449", fg="white")

# ------------------------------------------------------------
# BOTÓN SALIR
# ------------------------------------------------------------
tk.Button(
    ventana,
    text="Salir del Panel",
    command=ventana.quit,
    width=25,
    height=2,
    bg="#B3B6B7",
    fg="black",
    font=("Segoe UI", 10, "bold"),
    relief="ridge",
    borderwidth=3,
    cursor="hand2",
    activebackground="#909497",
    activeforeground="white"
).pack(pady=10)

# ------------------------------------------------------------
# PIE DE PÁGINA (FIJO Y SIEMPRE VISIBLE)
# ------------------------------------------------------------
frame_footer = tk.Frame(ventana, bg="#EAEDED")
frame_footer.pack(side="bottom", fill="x", pady=(5, 5))

pie = tk.Label(
    frame_footer,
    text="© 2025 ELITE Ingenieros S.A.S. | Pasión por lo que hacemos",
    font=("Segoe UI", 9, "italic"),
    fg="#515A5A",
    bg="#EAEDED"
)
pie.pack(pady=2)

ventana.mainloop()



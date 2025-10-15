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
from tkinter import ttk, scrolledtext, messagebox
from PIL import Image, ImageTk
import sys
import io
from datetime import datetime  # 🕒 Para mostrar hora en el pie

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
# FUNCIÓN DE EFECTO VISUAL EN BOTONES
# ------------------------------------------------------------
def resaltar_boton(boton):
    """Cambia color del botón mientras se ejecuta y luego lo restaura."""
    color_original = boton.cget("bg")
    boton.config(bg="#27AE60")  # Verde brillante
    ventana.update_idletasks()
    return color_original

def restaurar_boton(boton, color_original):
    """Restaura color original del botón."""
    boton.config(bg=color_original)
    ventana.update_idletasks()

# ------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EJECUCIÓN
# ------------------------------------------------------------
def ejecutar_comando(nombre, comando, boton=None):
    """Ejecuta un script y muestra salida en el log."""
    def tarea():
        log_text.insert(tk.END, f"\n🚀 Iniciando {nombre}...\n", "info")
        log_text.see(tk.END)
        barra_progreso.start()

        # 🟢 Mostrar estado en pie
        hora = datetime.now().strftime("%I:%M %p")
        pie_estado.config(text=f"🔄 Procesando {nombre}...  |  {hora}", fg="#1A5276")
        ventana.update_idletasks()

        color_original = resaltar_boton(boton) if boton else None

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
                hora = datetime.now().strftime("%I:%M %p")
                pie_estado.config(text=f"✅ {nombre} completado con éxito.  |  {hora}", fg="#27AE60")
            else:
                log_text.insert(tk.END, f"\n❌ Error al ejecutar {nombre}:\n{error_salida}\n", "error")
                pie_estado.config(text=f"⚠️ Error en {nombre}. Revisa el log.", fg="#C0392B")

        except Exception as e:
            log_text.insert(tk.END, f"\n⚠️ Error inesperado: {e}\n", "error")
            pie_estado.config(text=f"⚠️ Error en {nombre}. Revisa el log.", fg="#C0392B")

        finally:
            barra_progreso.stop()
            if boton and color_original:
                restaurar_boton(boton, color_original)
            log_text.insert(tk.END, "-" * 60 + "\n", "separador")
            log_text.see(tk.END)
            pie_estado.config(text="⚙️ Esperando acción del usuario...", fg="#1B263B")
            ventana.update_idletasks()

    threading.Thread(target=tarea).start()

# ------------------------------------------------------------
# COMANDOS DE LOS BOTONES
# ------------------------------------------------------------
def ejecutar_hv():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_HV}" --dataset HV --output "data_clean/HV_limpio.xlsx"'
    ejecutar_comando("HABITACIÓN VIVIENDAS", comando, btn_hv)

def ejecutar_puntos():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_PUNTOS}" --dataset PUNTOS --output "data_clean/PUNTOS_limpio.xlsx"'
    ejecutar_comando("PUNTOS DE CONEXIÓN", comando, btn_puntos)

def ejecutar_prepago():
    comando = f'python -X utf8 escenario1_individual.py --input "{RUTA_PREPAGO}" --dataset PREPAGO --output "data_clean/PREPAGO_limpio.xlsx"'
    ejecutar_comando("PREPAGO", comando, btn_prepago)

def ejecutar_merge():
    comando = 'python -X utf8 merge_escenario2.py'
    ejecutar_comando("MERGE", comando, btn_merge)

def ejecutar_control_vacios():
    comando = 'python -X utf8 diagnostico_control.py'
    ejecutar_comando("CONTROL DE VACÍOS", comando, btn_vacios)

# ------------------------------------------------------------
# INTERFAZ GRÁFICA
# ------------------------------------------------------------
ventana = tk.Tk()
ventana.title("Control ANS - ELITE Ingenieros S.A.S.")
ventana.config(bg="#EAEDED")

# ------------------------------------------------------------
# BARRA SUPERIOR CON RELOJ VERDE
# ------------------------------------------------------------
from datetime import datetime

frame_topbar = tk.Frame(ventana, bg="#1E8449", height=22)
frame_topbar.pack(fill="x")

reloj_top = tk.Label(
    frame_topbar,
    font=("Segoe UI", 9, "bold"),
    fg="white",
    bg="#1E8449",
    anchor="e"
)
reloj_top.pack(side="right", padx=15)

def actualizar_hora_top():
    hora_actual = datetime.now().strftime("%I:%M:%S %p")
    reloj_top.config(text=f"{hora_actual}")
    ventana.after(1000, actualizar_hora_top)

actualizar_hora_top()


# 🔹 Ajuste dinámico del tamaño según pantalla
screen_w = ventana.winfo_screenwidth()
screen_h = ventana.winfo_screenheight()

ancho = int(screen_w * 0.55)
alto = int(screen_h * 0.78)
x = (screen_w // 2) - (ancho // 2)
y = (screen_h // 2) - (alto // 2)
ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
ventana.resizable(False, False)

# ------------------------------------------------------------
# ENCABEZADO PROFESIONAL
# ------------------------------------------------------------
frame_banner = tk.Frame(ventana, bg="#EAEDED", height=120)
frame_banner.pack(fill="x")

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

elite_label = tk.Label(frame_superior, text="ELITE ", font=("Segoe UI", 18, "bold"), fg="black", bg="#EAEDED")
elite_label.pack(side="left")

ingenieros_label = tk.Label(frame_superior, text="Ingenieros S.A.S.", font=("Segoe UI", 18, "bold"), fg="#1E8449", bg="#EAEDED")
ingenieros_label.pack(side="left")

titulo_control = tk.Label(frame_banner, text="Control ANS", font=("Segoe UI", 14, "bold"), fg="#1B263B", bg="#EAEDED")
titulo_control.pack(pady=(0, 10))

# ------------------------------------------------------------
# BOTONES PRINCIPALES
# ------------------------------------------------------------
frame_botones = tk.Frame(ventana, bg="#EAEDED")
frame_botones.pack(pady=5)

btn_hv = tk.Button(frame_botones, text="HABITACIÓN VIVIENDA", command=ejecutar_hv, width=20, height=2,
                   bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"), relief="ridge", borderwidth=3, cursor="hand2",
                   activebackground="#229954", activeforeground="white")
btn_hv.pack(side="left", padx=8)

btn_puntos = tk.Button(frame_botones, text="PUNTOS DE CONEXIÓN", command=ejecutar_puntos, width=20, height=2,
                       bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"), relief="ridge", borderwidth=3, cursor="hand2",
                       activebackground="#229954", activeforeground="white")
btn_puntos.pack(side="left", padx=8)

btn_prepago = tk.Button(frame_botones, text="PREPAGO", command=ejecutar_prepago, width=20, height=2,
                        bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"), relief="ridge", borderwidth=3, cursor="hand2",
                        activebackground="#229954", activeforeground="white")
btn_prepago.pack(side="left", padx=8)

btn_merge = tk.Button(frame_botones, text="GENERAR MERGE", command=ejecutar_merge, width=20, height=2,
                      bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"), relief="ridge", borderwidth=3, cursor="hand2",
                      activebackground="#229954", activeforeground="white")
btn_merge.pack(side="left", padx=8)

# ------------------------------------------------------------
# BOTÓN ADICIONAL (CONTROL DE VACÍOS)
# ------------------------------------------------------------
frame_boton_extra = tk.Frame(ventana, bg="#EAEDED")
frame_boton_extra.pack(pady=(8, 5))

btn_vacios = tk.Button(frame_boton_extra, text="CONTROL DE VACÍOS", command=ejecutar_control_vacios,
                       width=25, height=2, bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"),
                       relief="ridge", borderwidth=3, cursor="hand2",
                       activebackground="#229954", activeforeground="white")
btn_vacios.pack()

# ------------------------------------------------------------
# BARRA DE PROGRESO
# ------------------------------------------------------------
barra_progreso = ttk.Progressbar(ventana, orient="horizontal", mode="indeterminate", length=450)
barra_progreso.pack(pady=(5, 5))

# ------------------------------------------------------------
# ÁREA DE LOG
# ------------------------------------------------------------
frame_log = tk.Frame(ventana, bg="#EAEDED")
frame_log.pack(fill="both", expand=False, pady=(5, 0))

log_text = scrolledtext.ScrolledText(frame_log, width=90, height=12, bg="white", font=("Consolas", 9))
log_text.pack(padx=15, pady=(5, 10), expand=True, fill="both")

log_text.tag_config("info", foreground="#2980B9")
log_text.tag_config("success", foreground="#27AE60")
log_text.tag_config("error", foreground="#C0392B")
log_text.tag_config("separador", foreground="#95A5A6")

# ------------------------------------------------------------
# BOTÓN SALIR
# ------------------------------------------------------------
frame_salida = tk.Frame(ventana, bg="#EAEDED")
frame_salida.pack(pady=(5, 5))

tk.Button(frame_salida, text="Salir del Panel", command=ventana.quit,
          width=25, height=2, bg="#1E8449", fg="white", font=("Segoe UI", 10, "bold"),
          relief="ridge", borderwidth=3, cursor="hand2",
          activebackground="#229954", activeforeground="white").pack()

# ------------------------------------------------------------
# PIE DE PÁGINA (corporativo + dinámico)
# ------------------------------------------------------------
frame_footer = tk.Frame(ventana, bg="#EAEDED")
frame_footer.pack(side="bottom", fill="x", pady=(0, 5), ipady=4)

# Línea divisoria superior
tk.Frame(frame_footer, bg="#B3B6B7", height=2).pack(fill="x", pady=(2, 3))

# Marco interno con dos etiquetas: estado y texto corporativo
frame_pie = tk.Frame(frame_footer, bg="#EAEDED")
frame_pie.pack(fill="x", pady=(0, 3))

# Estado dinámico (lado izquierdo)
pie_estado = tk.Label(
    frame_pie,
    text="⚙️ Esperando acción del usuario...",
    font=("Segoe UI", 9, "italic"),
    fg="#1B263B",
    bg="#EAEDED",
    anchor="w"
)
pie_estado.pack(side="left", padx=(15, 0))

# Texto corporativo (lado derecho)
pie_corporativo = tk.Label(
    frame_pie,
    text="© 2025 ELITE Ingenieros S.A.S.  |  Pasión por lo que hacemos.",
    font=("Segoe UI", 9, "italic"),
    fg="#1B263B",
    bg="#EAEDED",
    anchor="e"
)
pie_corporativo.pack(side="right", padx=(0, 15))


# ------------------------------------------------------------
# INICIAR INTERFAZ
# ------------------------------------------------------------
ventana.mainloop()

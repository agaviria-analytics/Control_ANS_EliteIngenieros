# 🧠 Control ANS – Elite Ingenieros S.A.S.

**Autor:** Héctor A. Gaviria  
**Año:** 2025  
**Tipo de proyecto:** Panel Empresarial en Python + Tkinter  
**Estado:** ✅ En ejecución

---

## 🚀 Descripción

**Control ANS** es un panel automatizado desarrollado en **Python (Tkinter)** para la empresa **Elite Ingenieros S.A.S.**, diseñado con el propósito de:

- Estandarizar y limpiar los reportes provenientes de Fénix / ENTER.  
- Consolidar tres escenarios: **HV**, **PUNTOS DE CONEXIÓN** y **PREPAGO**.  
- Generar automáticamente un archivo `MERGE_ANS.xlsx` limpio y listo para análisis en Power BI.  
- Reducir tiempos operativos mediante ejecución de scripts desde una interfaz visual.

---

## 🧩 Estructura del proyecto

Proyecto_ANS/
│
├── data_raw/ # Archivos originales de entrada (Excel)
├── data_clean/ # Archivos procesados y consolidados
├── escenario1_individual.py # Limpieza de escenarios individuales
├── merge_escenario2.py # Consolidación MERGE
├── menu_proyecto_ans.py # Interfaz gráfica (Tkinter)
├── requirements.txt # Librerías necesarias
└── iniciar_panel.bat # Script rápido de ejecución


---

## 💻 Tecnologías utilizadas

| Herramienta | Uso principal |
|--------------|----------------|
| **Python 3.12** | Lenguaje base del proyecto |
| **Tkinter** | Interfaz gráfica de usuario |
| **Pandas / Numpy** | Limpieza y transformación de datos |
| **Pillow (PIL)** | Manejo del logo corporativo |
| **OpenPyXL** | Exportación a Excel |

---

## ⚙️ Instalación y ejecución

1️⃣ Clona este repositorio  
```bash
git clone https://github.com/tu_usuario/Control_ANS_EliteIngenieros.git

2️⃣ Crea el entorno virtual
python -m venv venv

3️⃣ Actívalo
En Windows:
venv\Scripts\activate

4️⃣ Instala las dependencias
pip install -r requirements.txt

5️⃣ Ejecuta el panel
python menu_proyecto_ans.py

🧠 Características destacadas
Interfaz profesional con colores corporativos (verde y gris).
Botones activos para cada escenario con barra de progreso y log de ejecución.
Pie de página fijo con créditos y lema empresarial.
Exportación automática a carpeta data_clean.

📸 Capturas
| Vista                        | Imagen                                 |
| ---------------------------- | -------------------------------------- |
| **Panel principal**          | ![Panel principal](data_raw/elite.png) |
| **Proceso MERGE completado** | *(Próxima captura)*                    |

🏢 Créditos
Desarrollado por Héctor A. Gaviria
© 2025 Elite Ingenieros S.A.S. | Pasión por lo que hacemos.

🌟 Próximos pasos

Integración directa con Power BI para actualización en tiempo real.
Automatización por tareas programadas (Scheduler).
Publicación en entorno corporativo con logs centralizados.
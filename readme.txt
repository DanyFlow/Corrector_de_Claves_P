# 📋 Control de Asistencia y Corrector de Claves

Aplicación web desarrollada en **Python** y **Streamlit** diseñada para automatizar la transformación de claves de pago, el cruce de datos entre reportes de asistencia (Q1, Q2 y ACTIVOS) y la generación de archivos de actualización de códigos en formato CSV.

---

## 🚀 Características Principales

* **Compatibilidad Multiformato:** Lee archivos de Excel en formatos `.xlsx`, `.xlsm`, `.xls` (legacy) y `.xlsb` (binarios) de forma transparente sin alterar los archivos originales.
* **Transformación Automática de Claves:** Replica la lógica de negocio para reestructurar formatos de claves de pago específicas (claves con prefijo `R014`).
* **Cruce Intuitivo de Datos (Lookup):** Compara los registros de empleados en la lista de `ACTIVOS` contra las tablas `Q1` y `Q2` (priorizando Q1) para asociar la clave de pago correspondiente.
* **Generación de CSV Estandarizado:** Exporta el archivo `CODIGOS.csv` formateado en `UTF-8 con BOM` para garantizar la compatibilidad inmediata de tildes y caracteres especiales al abrirlo en Microsoft Excel.
* **Interfaz Modular e Interactiva:** Permite ejecutar cada módulo por separado o correr el flujo completo con un solo clic.

---

## 🛠️ Tecnologías y Librerías Utilizadas

* **Python** (Lógica principal de procesamiento y análisis de datos)
* **Streamlit** (Interfaz web interactiva)
* **Pandas** (Indexación, cruce de información y limpieza de datos)
* **Openpyxl, xlrd, pyxlsb** (Motores de lectura para las distintas extensiones de Excel)

---

## 📂 Estructura del Repositorio

```text
├── CorrectorClaves.py      # Archivo principal de la aplicación en Streamlit
├── requirements.txt        # Dependencias y librerías del proyecto
└── README.md               # Documentación general del proyecto
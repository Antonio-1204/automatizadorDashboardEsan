# Automatizador de Descarga y Procesamiento de Tickets - ESANLabs

Este proyecto automatiza la descarga de reportes de tickets desde el ServiceDesk de ESAN, realiza la limpieza y transformación de los datos, y sincroniza la información procesada con Google Sheets y Microsoft Power Automate.

## 🚀 Características principales
* **Extracción Automatizada:** Inicia sesión en ServiceDesk usando Selenium en modo incógnito y descarga el reporte más reciente de manera silenciosa (sin ventanas emergentes).
* **Manejo Dinámico de Archivos:** Detecta automáticamente si el sistema genera el archivo como `Requestexport.XLSX`
* **Transformación de Datos:** Limpia fechas, reasigna grupos de trabajo según reglas de negocio y categoriza servicios usando expresiones regulares y Pandas.
* **Integración de APIs:** Envía la información segmentada a Google Apps Script (Sheets) y gatilla flujos en Power Automate.

## 🛠️ Requisitos Previos

Asegúrate de tener instalado **Python 3.9+** y las siguientes librerías. Puedes instalarlas ejecutando:

```bash
pip install pandas requests selenium webdriver-manager openpyxl unidecode pygetwindow

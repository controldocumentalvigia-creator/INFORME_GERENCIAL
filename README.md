# Informe Gerencial Ejecutivo - Streamlit Cloud

Aplicación ejecutiva para análisis operativo, financiero, rentabilidad, cierre operativo y comparativo de periodos.

## Archivos requeridos

- app.py
- requirements.txt
- runtime.txt
- .streamlit/config.toml

## Despliegue en Streamlit Cloud

1. Subir estos archivos al repositorio en GitHub.
2. Crear nueva app en Streamlit Cloud.
3. Main file path: app.py
4. Python: definido por runtime.txt como python-3.11.
5. Clear cache and reboot si se actualizan paquetes.

## Formatos soportados

- .xlsx
- .xlsm
- .xls (requiere xlrd, incluido)
- .csv

## KPIs clave

- OTIF Operativo = CUMPLIDO / total de servicios válidos no anulados.
- Cumplimiento de Cierre = CUMPLIDO OPERATIVO / total de servicios válidos no anulados.
- Pendientes Operativos = estados diferentes de CUMPLIDO y CUMPLIDO OPERATIVO.
- Margen = V.CLIENTE - V.CONDUCT.
- Rentabilidad = Margen / V.CLIENTE.

## Versión

Versión final Power BI style, optimizada para bases de 20.000+ registros.

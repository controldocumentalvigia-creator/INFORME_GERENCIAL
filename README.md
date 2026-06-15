# Informe Gerencial Ejecutivo - Streamlit Cloud

Archivos requeridos en la raíz del repositorio:
- app.py
- requirements.txt
- runtime.txt
- .streamlit/config.toml

Importante:
- Esta versión NO instala streamlit-aggrid para evitar errores de instalación en Streamlit Cloud.
- La app usa tablas nativas de Streamlit si AgGrid no está disponible.
- Permite cargar bases Excel/CSV desde la barra lateral.

Después de subir cambios a GitHub:
1. Streamlit Cloud > Manage app.
2. Clear cache.
3. Reboot app.

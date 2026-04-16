# Plan de Gestión de Viaje (ALS Ecuador)

App **Streamlit** para:

- Completar el **Plan de Gestión de Viaje** como formulario.
- Personas y vehículos desde **Google Sheets** (gspread).
- Exportar el plan a **PDF**.

## Requisitos

- Python 3.10+

## Instalación (local)

```bash
python -m venv .venv
# Windows: .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Ejecutar (local)

```bash
streamlit run app.py
```

Credenciales Google: archivo `credenciales.json` en la carpeta del proyecto **o** variables en `.streamlit/secrets.toml` (no subir al repositorio).

## Despliegue (Streamlit Cloud)

- Repositorio en GitHub.
- **Main file:** `app.py`.
- **Secrets** en el panel de la app: bloque `[google_service_account]` con la cuenta de servicio.
- Compartir las hojas de cálculo con el `client_email` de esa cuenta.

# Plan de Gestión de Viaje (ALS Ecuador) - Formulario automático

Esta carpeta contiene una app local para:

- Llenar el **Plan de Gestión de Viaje** como formulario (con opciones).
- Autocompletar personas desde `BD.csv`.
- Exportar el plan final a **PDF**.

## Requisitos

- Python 3.10+ instalado en Windows

## Instalación

En PowerShell, dentro de esta carpeta:

```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Ejecutar

```bash
streamlit run app.py
```

Al final usa el botón **Generar PDF** para descargarlo.

## Datos

- `BD.csv`: base de personas (nombre, celular, cédula).


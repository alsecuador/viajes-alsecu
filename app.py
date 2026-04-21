from __future__ import annotations

import io
import json
import os
from collections.abc import Mapping
import re
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime, time
from typing import Any, Literal

import gspread
import pandas as pd
import streamlit as st
from dateutil import parser as date_parser
from google.oauth2.service_account import Credentials
from PIL import Image as PILImage

try:
    from streamlit_paste_button import paste_image_button
except Exception:
    paste_image_button = None

from pdf_builder import build_plan_pdf


# Google Sheets (BD personal y vehículos). Por defecto estos IDs; en local: variables de entorno;
# en Streamlit Cloud: mismas claves en st.secrets (ver _env_or_secret).
_DEFAULT_GSHEET_BD_ID = "18QdZarvALNmq0NPG1XNaxE-Gwv9iHnNcrGsELuWUsfA"
_DEFAULT_GSHEET_VEHICLES_ID = "1ZnUoIG--aGTtUC54pp_UJqQOm0d2ERfAc64XqhPy_yY"


def _env_or_secret(key: str, default: str = "") -> str:
    """Variable de entorno o clave en st.secrets (primer nivel o dentro de un bloque tipo [seccion])."""
    v = (os.environ.get(key) or "").strip()
    if v:
        return v
    try:
        sec = st.secrets
        if key in sec:
            return str(sec[key]).strip()
        for _block, val in sec.items():
            # Streamlit puede exponer bloques [x] como Mapping que no es dict puro.
            if isinstance(val, Mapping) and key in val:
                inner = val.get(key)
                if inner is not None and str(inner).strip():
                    return str(inner).strip()
    except Exception:
        pass
    return (default or "").strip()


def _gsheet_bd_spreadsheet_id() -> str:
    return _env_or_secret("PLAN_BD_SPREADSHEET_ID", _DEFAULT_GSHEET_BD_ID) or _DEFAULT_GSHEET_BD_ID


def _gsheet_vehicles_spreadsheet_id() -> str:
    return (
        _env_or_secret("PLAN_VEHICLES_SPREADSHEET_ID", _DEFAULT_GSHEET_VEHICLES_ID)
        or _DEFAULT_GSHEET_VEHICLES_ID
    )


def _gsheet_ubicaciones_spreadsheet_id() -> str:
    return _env_or_secret("PLAN_UBICACIONES_SPREADSHEET_ID", "")

# Hoja / índice de worksheet (0 = primera hoja, o nombre exacto del tab).
PLAN_BD_WORKSHEET_ENV = "PLAN_BD_WORKSHEET"
PLAN_VEHICLES_WORKSHEET_ENV = "PLAN_VEHICLES_WORKSHEET"
PLAN_UBICACIONES_WORKSHEET_ENV = "PLAN_UBICACIONES_WORKSHEET"
# Libro CODIFICACIÓN con pestañas separadas (INEC): provincias / cantones / parroquias
PLAN_UBICACIONES_SHEET_PROVINCIAS_ENV = "PLAN_UBICACIONES_SHEET_PROVINCIAS"
PLAN_UBICACIONES_SHEET_CANTONES_ENV = "PLAN_UBICACIONES_SHEET_CANTONES"
PLAN_UBICACIONES_SHEET_PARROQUIAS_ENV = "PLAN_UBICACIONES_SHEET_PARROQUIAS"

# Credenciales: en producción usar st.secrets (ver _load_service_account_dict). Local: credenciales.json
CREDENTIALS_JSON_PATH = (os.environ.get("PLAN_CREDENTIALS_JSON") or "credenciales.json").strip()

# Máximo de técnicos por viaje (1–20). Por defecto 10. Ej.: PLAN_MAX_TECNICOS=12
MAX_TECNICOS_ENV = "PLAN_MAX_TECNICOS"
DEFAULT_COMPANY_NAME = "ALS ECUADOR"
DEFAULT_LOGO_PATH = os.path.join("assets", "als_logo.png")
DEFAULT_DOC_CODE = "RU-40"
DEFAULT_REV = "Rev. 01"
DEFAULT_DOC_DATE = "08-09-2025"

# Carpeta por defecto para guardar PDFs (compartida en red). Sobrescribible en la app o con variable de entorno.
PDF_OUTPUT_DIR_ENV = "PLAN_PDF_OUTPUT_DIR"

# Último PDF generado en memoria (para st.download_button tras reruns; Cloud no tiene disco compartido con el usuario).
SESSION_PLAN_PDF_BYTES = "plan_export_pdf_bytes"
SESSION_PLAN_PDF_FILENAME = "plan_export_pdf_filename"

# Ubicaciones Ecuador: provincia → cantón → parroquia
# Prioridad: Excel oficial CODIFICACIÓN_2025.xlsx (misma carpeta que la app), luego ubicaciones.csv
UBICACIONES_XLSX_PRIMARY = "CODIFICACIÓN_2025.xlsx"
UBICACIONES_XLSX_ALT = "CODIFICACION_2025.xlsx"
UBICACIONES_CSV_PATH = "ubicaciones.csv"
UBICACIONES_SHEET_ENV = "PLAN_CODIFICACION_SHEET"  # opcional: nombre de hoja Excel

# Contactos de emergencia: atajos con nombre en BD; o cualquier fila de BD; o texto libre
EMERGENCY_PRESET_TO_BD_NAME: dict[str, str] = {
    "Santiago Montalvan": "Montalvan Samaniego Santiago Javier",
    "David Solano": "Solano Bazurto David Roberto",
}
EMERGENCY_FROM_BD = "Otra persona de la lista (BD)…"
EMERGENCY_OTRO_TEXTO = "Otro (escribir nombre y teléfono)"
EMERGENCY_SELECT_OPTIONS = [
    "—",
    "Santiago Montalvan",
    "David Solano",
    EMERGENCY_FROM_BD,
    EMERGENCY_OTRO_TEXTO,
]

# Persona no catalogada: técnicos, firmas, etc.
PERSON_SEL_OTRO_MANUAL = "Otro (no está en BD — escribir datos)"

# Técnicos / conductores en el viaje (máximo por defecto; ver _max_tecnicos_viaje())
MAX_TECNICOS_VIAJE_DEFAULT = 10
STOP_MOTIVO_OPTIONS = [
    "Logistica",
    "Combustible",
    "Descanso",
    "Monitoreo",
    "Retiro de companero",
    "Envio encomiendas",
    "Descanso y combustible",
    "Alimentacion (desayuno)",
    "Alimentacion (almuerzo)",
    "Alimentacion (merienda)",
]
STOP_TIEMPO_MIN_OPTIONS = [5, 10, 15, 20, 30, 45, 60, 120, 240, 300, 360, 420]


def _coerce_tiempo_min(val: Any) -> int:
    if val in STOP_TIEMPO_MIN_OPTIONS:
        return int(val)
    s = str(val or "").strip()
    if not s:
        return STOP_TIEMPO_MIN_OPTIONS[0]
    try:
        n = int(float(s))
    except ValueError:
        return STOP_TIEMPO_MIN_OPTIONS[0]
    if n in STOP_TIEMPO_MIN_OPTIONS:
        return n
    return STOP_TIEMPO_MIN_OPTIONS[0]


def _resolve_emergency_contact(
    choice: str,
    otro_nombre: str,
    otro_tel: str,
    people_df: pd.DataFrame,
    *,
    bd_pick_name: str = "",
) -> tuple[str, str, str]:
    """Devuelve (nombre completo para PDF, teléfono, cédula)."""
    if choice == "—":
        return "", "", ""
    if choice in EMERGENCY_PRESET_TO_BD_NAME:
        bd_name = EMERGENCY_PRESET_TO_BD_NAME[choice]
        d = _person_lookup(people_df, bd_name)
        return d["name"], d["cel"], d["id"]
    if choice == EMERGENCY_FROM_BD:
        if not bd_pick_name or bd_pick_name == "—":
            return "", "", ""
        d = _person_lookup(people_df, bd_pick_name)
        return d["name"], d["cel"], d["id"]
    # Otro (texto libre)
    return (otro_nombre or "").strip(), (otro_tel or "").strip(), ""


def _parse_extra_names(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []
    parts = re.split(r"[\n,;]+", str(text))
    return [p.strip() for p in parts if p.strip()]


def _merge_passenger_lists(from_multiselect: list[str], extra_text: str) -> list[str]:
    order: list[str] = []
    seen: set[str] = set()
    for p in from_multiselect + _parse_extra_names(extra_text):
        p = (p or "").strip()
        if p and p not in seen:
            seen.add(p)
            order.append(p)
    return order


def _opciones_con_manual(people_opts: list[str]) -> list[str]:
    return people_opts + [PERSON_SEL_OTRO_MANUAL]


def _valor_firma(sel: str, manual_nombre: str) -> str:
    if sel == "—":
        return ""
    if sel == PERSON_SEL_OTRO_MANUAL:
        return (manual_nombre or "").strip()
    return sel


def _parse_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str) and value.strip():
        try:
            return date_parser.parse(value, dayfirst=True).date()
        except Exception:
            return None
    return None


def _format_hora_plan(t: time) -> str:
    """Hora para el PDF / datos (ej. 17h00, 09h05)."""
    return f"{t.hour}h{t.minute:02d}"


def _pil_to_png_bytes(img: PILImage.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _image_bytes_input(
    *,
    upload_label: str,
    upload_key: str,
    paste_label: str,
    allowed_types: list[str],
) -> bytes | None:
    """
    Entrada de imagen con dos opciones:
    - Pegar screenshot desde portapapeles (Ctrl+V en navegador compatible).
    - Subir archivo (fallback universal).
    Prioridad: imagen pegada > archivo subido.
    """
    pasted_bytes: bytes | None = None
    if paste_image_button is not None:
        paste_res = paste_image_button(
            paste_label,
            key=f"{upload_key}_paste_btn",
            errors="ignore",
        )
        if paste_res is not None and getattr(paste_res, "image_data", None) is not None:
            try:
                pasted_bytes = _pil_to_png_bytes(paste_res.image_data)
                st.caption("Imagen pegada desde portapapeles.")
            except Exception:
                pasted_bytes = None

    uploaded = st.file_uploader(upload_label, type=allowed_types, key=upload_key)
    uploaded_bytes = uploaded.read() if uploaded is not None else None
    return pasted_bytes or uploaded_bytes


def _max_tecnicos_viaje() -> int:
    raw = _env_or_secret(MAX_TECNICOS_ENV).strip()
    if raw.isdigit():
        return max(1, min(20, int(raw)))
    return MAX_TECNICOS_VIAJE_DEFAULT


def _norm_header(h: str) -> str:
    t = str(h).strip().upper()
    for a, b in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("\u00d1", "N")):
        t = t.replace(a, b)
    return " ".join(t.split())


def _pick_csv_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    nh = {_norm_header(c): c for c in df.columns}
    for cand in candidates:
        k = _norm_header(cand)
        if k in nh:
            return nh[k]
    return None


def _parse_worksheet_spec(raw: str | None) -> str | int:
    """Índice 0-based (string numérico) o título de pestaña."""
    s = (raw or "").strip()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
    return s


def _bd_worksheet_spec() -> str | int:
    return _parse_worksheet_spec(_env_or_secret(PLAN_BD_WORKSHEET_ENV))


def _vehicles_worksheet_spec() -> str | int:
    return _parse_worksheet_spec(_env_or_secret(PLAN_VEHICLES_WORKSHEET_ENV))


def _ubicaciones_gsheet_worksheet_spec() -> str | int:
    return _parse_worksheet_spec(_env_or_secret(PLAN_UBICACIONES_WORKSHEET_ENV))


def _load_service_account_dict() -> dict[str, Any]:
    """
    Credenciales de cuenta de servicio para gspread.
    Orden: bloque en st.secrets → JSON en raíz de secrets → archivo credenciales.json → GOOGLE_APPLICATION_CREDENTIALS.
    """
    try:
        sec = st.secrets
        for block_name in ("google_service_account", "gspread", "credenciales", "service_account"):
            if block_name in sec:
                block = sec[block_name]
                if not isinstance(block, Mapping):
                    continue
                # Claves tipo PLAN_* en el mismo bloque [google_service_account] no son del JSON de Google.
                d = {k: block[k] for k in block.keys() if not str(k).startswith("PLAN_")}
                if d.get("type") == "service_account":
                    return d
        if sec.get("type") == "service_account":
            return {k: sec[k] for k in sec.keys()}
    except Exception:
        pass
    for path in (
        CREDENTIALS_JSON_PATH,
        (os.environ.get("GOOGLE_APPLICATION_CREDENTIALS") or "").strip(),
    ):
        if path and os.path.isfile(path):
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get("type") == "service_account":
                return data
    raise RuntimeError(
        "No hay credenciales de Google: defínelas en st.secrets (cuenta de servicio) o en "
        f"`{CREDENTIALS_JSON_PATH}` / GOOGLE_APPLICATION_CREDENTIALS."
    )


@st.cache_resource(show_spinner=False)
def _gspread_client() -> gspread.Client:
    scopes = ("https://www.googleapis.com/auth/spreadsheets",)
    info = _load_service_account_dict()
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def _open_worksheet(spreadsheet_id: str, worksheet: str | int) -> gspread.Worksheet:
    gc = _gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    if isinstance(worksheet, int):
        ws = sh.get_worksheet(worksheet)
        if ws is None:
            raise ValueError(f"Worksheet index {worksheet} no existe en {spreadsheet_id!r}")
        return ws
    if isinstance(worksheet, str) and worksheet.isdigit():
        ws = sh.get_worksheet(int(worksheet))
        if ws is None:
            raise ValueError(f"Worksheet index {worksheet} no existe en {spreadsheet_id!r}")
        return ws
    return sh.worksheet(worksheet)


def _worksheet_to_dataframe(ws: gspread.Worksheet) -> pd.DataFrame:
    vals = ws.get_all_values()
    if not vals:
        return pd.DataFrame()
    header = [str(h).strip() for h in vals[0]]
    data = vals[1:]
    if not header:
        return pd.DataFrame()
    if not data:
        return pd.DataFrame(columns=header)
    ncols = len(header)
    padded: list[list[str]] = []
    for row in data:
        cells = [str(c) if c is not None else "" for c in row]
        if len(cells) < ncols:
            cells = cells + [""] * (ncols - len(cells))
        else:
            cells = cells[:ncols]
        padded.append(cells)
    df = pd.DataFrame(padded, columns=header)
    return df.fillna("").astype(str)


def _ubicaciones_sheet_title_provinces() -> str:
    return (_env_or_secret(PLAN_UBICACIONES_SHEET_PROVINCIAS_ENV, "PROVINCIAS") or "PROVINCIAS").strip()


def _ubicaciones_sheet_title_cantones() -> str:
    return (_env_or_secret(PLAN_UBICACIONES_SHEET_CANTONES_ENV, "CANTONES") or "CANTONES").strip()


def _ubicaciones_sheet_title_parroquias() -> str:
    return (_env_or_secret(PLAN_UBICACIONES_SHEET_PARROQUIAS_ENV, "PARROQUIAS") or "PARROQUIAS").strip()


def _ubicaciones_raw_df_from_vals(vals: list[list[str]]) -> pd.DataFrame | None:
    """Devuelve un DataFrame con la primera fila de encabezado plausible (como Excel INEC)."""
    if not vals or len(vals) < 2:
        return None
    for header_ix in (1, 0, 2, 3):
        if header_ix >= len(vals):
            continue
        header = [str(h).strip() for h in vals[header_ix]]
        if not header or not any(h for h in header):
            continue
        data = vals[header_ix + 1 :]
        ncols = len(header)
        padded: list[list[str]] = []
        for row in data:
            cells = [str(c) if c is not None else "" for c in row]
            if len(cells) < ncols:
                cells = cells + [""] * (ncols - len(cells))
            else:
                cells = cells[:ncols]
            if any(str(c).strip() for c in cells):
                padded.append(cells)
        if padded:
            df = pd.DataFrame(padded, columns=header)
            return df.fillna("").astype(str)
    return None


def _pick_first_col(df: pd.DataFrame, markers: tuple[str, ...], *, avoid: tuple[str, ...] = ()) -> str | None:
    for c in df.columns:
        u = _col_norm(c)
        if any(av in u for av in avoid):
            continue
        for m in markers:
            if m in u:
                return c
    return None


def _ubicaciones_merge_innec_three_tabs(spreadsheet_id: str) -> pd.DataFrame | None:
    """
    Une pestañas tipo INEC: PROVINCIAS, CANTONES, PARROQUIAS (códigos + nombres DPA_*).
    """
    gc = _gspread_client()
    try:
        sh = gc.open_by_key(spreadsheet_id)
    except Exception:
        return None
    titles = {ws.title.strip() for ws in sh.worksheets()}
    pt, ct, rt = _ubicaciones_sheet_title_provinces(), _ubicaciones_sheet_title_cantones(), _ubicaciones_sheet_title_parroquias()
    if pt not in titles or ct not in titles or rt not in titles:
        return None

    def _read_tab(title: str) -> pd.DataFrame | None:
        try:
            ws = sh.worksheet(title)
            raw = _ubicaciones_raw_df_from_vals(ws.get_all_values())
            return raw
        except Exception:
            return None

    pdf, cdf, rdf = _read_tab(pt), _read_tab(ct), _read_tab(rt)
    if pdf is None or cdf is None or rdf is None or pdf.empty or cdf.empty or rdf.empty:
        return None

    p_code = _pick_first_col(pdf, ("DPA_PROVIN", "COD_PROVIN", "COD_PROV"))
    p_name = _pick_first_col(pdf, ("DPA_DESPRO", "DESPRO", "NOMBRE PROVINCIA", "PROVINCIA"))
    c_code = _pick_first_col(cdf, ("DPA_CANTON", "COD_CANTON", "COD_CANT"))
    c_name = _pick_first_col(cdf, ("DPA_DESCAN", "DESCAN", "NOMBRE CANTON", "CANTON", "CANTÓN"))
    c_prov = _pick_first_col(cdf, ("DPA_PROVIN", "COD_PROVIN", "COD_PROV"))
    r_par = _pick_first_col(rdf, ("DPA_PARROQ", "COD_PARROQ", "COD_PARRO"))
    r_name = _pick_first_col(rdf, ("DPA_DESPAR", "DESPAR", "NOMBRE PARROQUIA", "PARROQUIA"))
    r_cant = _pick_first_col(rdf, ("DPA_CANTON", "COD_CANTON", "COD_CANT"))

    if not all([p_code, p_name, c_code, c_name, c_prov, r_par, r_name, r_cant]):
        return None

    p = pdf[[p_code, p_name]].copy()
    p.columns = ["prov_c", "PROVINCIA"]
    c = cdf[[c_code, c_name, c_prov]].copy()
    c.columns = ["cant_c", "CANTON", "prov_c"]
    r = rdf[[r_par, r_name, r_cant]].copy()
    r.columns = ["par_c", "PARROQUIA", "cant_c"]
    for d in (p, c, r):
        for col in d.columns:
            d[col] = d[col].astype(str).str.strip().replace({"nan": "", "None": ""})
    p = p[(p["prov_c"] != "") & (p["PROVINCIA"] != "")].drop_duplicates(subset=["prov_c"])
    c = c[(c["cant_c"] != "") & (c["CANTON"] != "") & (c["prov_c"] != "")].drop_duplicates(subset=["cant_c"])
    r = r[(r["cant_c"] != "") & (r["PARROQUIA"] != "")].drop_duplicates(subset=["par_c"])
    if p.empty or c.empty or r.empty:
        return None
    m = r.merge(c, on="cant_c", how="inner").merge(p, on="prov_c", how="inner")
    if m.empty:
        return None
    out = m[["PROVINCIA", "CANTON", "PARROQUIA"]].drop_duplicates()
    out = out[(out["PROVINCIA"] != "") & (out["CANTON"] != "") & (out["PARROQUIA"] != "")]
    return out if not out.empty else None


def _ubicaciones_norm_from_sheet_matrix(vals: list[list[str]]) -> pd.DataFrame | None:
    """
    CODIFICACIÓN INEC en Google Sheets suele igual que en Excel: títulos en fila 2 (índice 1).
    Prueba varias filas de encabezado hasta que _normalize_ubicaciones_df reconozca columnas.
    """
    if not vals or len(vals) < 2:
        return None
    for header_ix in (1, 0, 2, 3):
        if header_ix >= len(vals):
            continue
        header = [str(h).strip() for h in vals[header_ix]]
        if not header or not any(h for h in header):
            continue
        data = vals[header_ix + 1 :]
        ncols = len(header)
        padded: list[list[str]] = []
        for row in data:
            cells = [str(c) if c is not None else "" for c in row]
            if len(cells) < ncols:
                cells = cells + [""] * (ncols - len(cells))
            else:
                cells = cells[:ncols]
            if any(str(c).strip() for c in cells):
                padded.append(cells)
        if not padded:
            continue
        df = pd.DataFrame(padded, columns=header)
        df = df.fillna("").astype(str)
        norm = _normalize_ubicaciones_df(df)
        if norm is not None:
            return norm
    return None


def _dataframe_to_worksheet(ws: gspread.Worksheet, df: pd.DataFrame) -> None:
    """Sustituye el contenido de la hoja por el DataFrame (cabecera + filas)."""
    ws.clear()
    if df.empty:
        headers = [str(c) for c in df.columns]
        ws.update([headers], value_input_option="USER_ENTERED")
        return
    headers = [str(c) for c in df.columns]
    body_df = df.fillna("").astype(str).replace({"<NA>": "", "nan": ""})
    rows = body_df.values.tolist()
    all_rows = [headers] + rows
    ws.update(all_rows, value_input_option="USER_ENTERED")


def gsheet_write_dataframe(spreadsheet_id: str, worksheet: str | int, df: pd.DataFrame) -> None:
    """Escribe un DataFrame a una pestaña de Google Sheets (API pública para mantener paridad con to_csv/to_excel)."""
    ws = _open_worksheet(spreadsheet_id, worksheet)
    _dataframe_to_worksheet(ws, df)


def gsheet_read_dataframe(spreadsheet_id: str, worksheet: str | int) -> pd.DataFrame:
    ws = _open_worksheet(spreadsheet_id, worksheet)
    return _worksheet_to_dataframe(ws)


def _prepare_raw_person_df(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.copy()
    raw.columns = [str(c).strip() for c in raw.columns]
    return raw


def _coerce_people_dataframe(raw: pd.DataFrame) -> pd.DataFrame:
    col_nombre = _pick_csv_column(
        raw,
        [
            "APELLIDOS Y NOMBRES",
            "NOMBRES Y APELLIDOS",
            "NOMBRE COMPLETO",
            "NOMBRES",
            "APELLIDOS NOMBRES",
        ],
    )
    col_cel = _pick_csv_column(
        raw,
        ["CELULAR", "TELEFONO", "TELÉFONO", "TEL", "CEL", "MOVIL", "MÓVIL", "TELEFONO MOVIL"],
    )
    col_id = _pick_csv_column(
        raw,
        ["C. IDENTIDAD", "CEDULA", "CÉDULA", "CI", "C.I.", "C.I", "IDENTIFICACION", "IDENTIFICACIÓN"],
    )
    if not col_nombre:
        return pd.DataFrame(columns=["APELLIDOS Y NOMBRES", "CELULAR", "C. IDENTIDAD"])

    n = len(raw)
    out = pd.DataFrame(
        {
            "APELLIDOS Y NOMBRES": raw[col_nombre].astype(str),
            "CELULAR": raw[col_cel].astype(str)
            if col_cel
            else pd.Series([""] * n, index=raw.index, dtype=object),
            "C. IDENTIDAD": raw[col_id].astype(str)
            if col_id
            else pd.Series([""] * n, index=raw.index, dtype=object),
        }
    )
    out["APELLIDOS Y NOMBRES"] = out["APELLIDOS Y NOMBRES"].str.strip()
    out["CELULAR"] = out["CELULAR"].str.strip()
    out["C. IDENTIDAD"] = out["C. IDENTIDAD"].str.strip()
    out = out[out["APELLIDOS Y NOMBRES"].str.len() > 0]
    return out


@st.cache_data(show_spinner=False)
def _cached_load_people_gsheet(_spreadsheet_id: str, _worksheet: str | int) -> pd.DataFrame:
    """Lee BD de personal desde Google Sheets (caché por id de libro + pestaña)."""
    if not _spreadsheet_id:
        return pd.DataFrame(columns=["APELLIDOS Y NOMBRES", "CELULAR", "C. IDENTIDAD"])
    try:
        ws = _open_worksheet(_spreadsheet_id, _worksheet)
        raw = _prepare_raw_person_df(_worksheet_to_dataframe(ws))
        return _coerce_people_dataframe(raw)
    except Exception:
        return pd.DataFrame(columns=["APELLIDOS Y NOMBRES", "CELULAR", "C. IDENTIDAD"])


def _load_people() -> pd.DataFrame:
    return _cached_load_people_gsheet(_gsheet_bd_spreadsheet_id(), _bd_worksheet_spec())


def _coerce_vehicles_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza hoja de vehículos (misma lógica que antes con CSV).

    - Columnas tipo Camionetas: PLACA, MODELO, COLOR (tipo por defecto: Camioneta)
    - O: PLACA, TIPO, MODELO_ANIO
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["PLACA", "TIPO", "MODELO_ANIO"])
    df = df.copy()
    df.columns = [c.strip().upper() for c in df.columns]
    if "PLACA" not in df.columns:
        return pd.DataFrame(columns=["PLACA", "TIPO", "MODELO_ANIO"])

    placa = df["PLACA"].astype(str).str.strip()

    if "TIPO" in df.columns:
        tipo = df["TIPO"].astype(str).str.strip()
        tipo = tipo.mask(tipo.isin(["", "nan", "None"]), "Camioneta")
    else:
        tipo = pd.Series(["Camioneta"] * len(df), index=df.index, dtype=object)

    if "MODELO_ANIO" in df.columns and not df["MODELO_ANIO"].fillna("").astype(str).str.strip().eq("").all():
        modelo = df["MODELO_ANIO"].astype(str).str.strip()
    elif "MODELO" in df.columns:
        m = df["MODELO"].astype(str).str.strip()
        if "COLOR" in df.columns:
            c = df["COLOR"].astype(str).str.strip().replace({"nan": "", "None": ""})
            modelo = m + c.map(lambda x: f" ({x})" if x else "")
        else:
            modelo = m
    else:
        modelo = pd.Series([""] * len(df), index=df.index)

    out = pd.DataFrame({"PLACA": placa, "TIPO": tipo, "MODELO_ANIO": modelo.astype(str).str.strip()})
    out = out[out["PLACA"].str.len() > 0]
    return out


@st.cache_data(show_spinner=False)
def _cached_load_vehicles_gsheet(_spreadsheet_id: str, _worksheet: str | int) -> pd.DataFrame:
    if not _spreadsheet_id:
        return pd.DataFrame(columns=["PLACA", "TIPO", "MODELO_ANIO"])
    try:
        ws = _open_worksheet(_spreadsheet_id, _worksheet)
        raw = _worksheet_to_dataframe(ws)
        return _coerce_vehicles_dataframe(raw)
    except Exception:
        return pd.DataFrame(columns=["PLACA", "TIPO", "MODELO_ANIO"])


def _load_vehicles() -> pd.DataFrame:
    return _cached_load_vehicles_gsheet(_gsheet_vehicles_spreadsheet_id(), _vehicles_worksheet_spec())


def _vehicle_options(df: pd.DataFrame) -> list[str]:
    if df is None or df.empty:
        return ["—"]
    # etiqueta: PLACA — TIPO — MODELO
    labels = []
    for _, r in df.iterrows():
        placa = str(r.get("PLACA") or "").strip()
        tipo = str(r.get("TIPO") or "").strip()
        modelo = str(r.get("MODELO_ANIO") or "").strip()
        labels.append(" — ".join([x for x in [placa, tipo, modelo] if x]))
    return ["—"] + sorted(set(labels))


def _vehicle_lookup(df: pd.DataFrame, label: str) -> dict[str, str]:
    if label == "—" or not label.strip() or df is None or df.empty:
        return {"placa": "", "tipo": "", "modelo_anio": ""}
    placa = label.split(" — ", 1)[0].strip()
    row = df[df["PLACA"] == placa].head(1)
    if row.empty:
        return {"placa": placa, "tipo": "", "modelo_anio": ""}
    r = row.iloc[0].to_dict()
    return {
        "placa": str(r.get("PLACA", "")).strip(),
        "tipo": str(r.get("TIPO", "")).strip(),
        "modelo_anio": str(r.get("MODELO_ANIO", "")).strip(),
    }


def _person_options(df: pd.DataFrame) -> list[str]:
    return ["—"] + sorted(df["APELLIDOS Y NOMBRES"].unique().tolist())


def _person_lookup(df: pd.DataFrame, name: str) -> dict[str, str]:
    if name == "—":
        return {"name": "", "cel": "", "id": ""}
    row = df[df["APELLIDOS Y NOMBRES"] == name].head(1)
    if row.empty:
        return {"name": name, "cel": "", "id": ""}
    r = row.iloc[0].to_dict()
    return {
        "name": str(r.get("APELLIDOS Y NOMBRES", "")).strip(),
        "cel": str(r.get("CELULAR", "")).strip(),
        "id": str(r.get("C. IDENTIDAD", "")).strip(),
    }


@dataclass
class Stop:
    n: int
    lugar: str = ""
    motivo: str = ""
    tiempo_min: str = ""


@dataclass
class PlanData:
    # Encabezado
    empresa_nombre: str = DEFAULT_COMPANY_NAME
    empresa_logo_bytes: bytes | None = None
    empresa_logo_mime: str | None = None
    doc_code: str = DEFAULT_DOC_CODE
    doc_rev: str = DEFAULT_REV
    doc_date: str = DEFAULT_DOC_DATE

    # 1. Datos generales — técnicos (cantidad según _max_tecnicos_viaje()), solo filas con persona elegida (no "—")
    conductores: list[str] = field(default_factory=list)
    cedulas_conductores: list[str] = field(default_factory=list)
    celulares_conductores: list[str] = field(default_factory=list)
    fecha_elab: date | None = None
    cargo: str = "Técnico de Operaciones"
    origen: str = ""
    placa: str = ""
    tipo_vehiculo: str = ""
    modelo_anio: str = ""
    emergencia_1: str = ""
    emergencia_2: str = ""
    tel_emergencia_1: str = ""
    tel_emergencia_2: str = ""
    cedula_emergencia_1: str = ""
    cedula_emergencia_2: str = ""

    # 2. Planificación
    destino: str = ""
    empresa: str = ""
    orden_trabajo: str = ""
    fecha_salida: date | None = None
    hora_salida: str = ""
    fecha_llegada: date | None = None
    hora_llegada: str = ""
    distancia_km: str = ""
    duracion_horas: str = ""
    proposito: str = ""
    condiciones_camino: str = ""

    # 3. Paradas (ida)
    paradas_ida: list[Stop] = field(default_factory=list)

    # 4. Peligros
    peligro_lluvia: bool = False
    peligro_niebla: bool = False
    peligro_nieve_hielo: bool = False
    peligro_nocturna: bool = False
    peligro_carretera_mala: bool = False
    peligro_delincuencia: bool = False
    peligro_accidentes_transito: bool = False
    otros_peligros: str = ""
    observaciones: str = ""
    international_sos_text: str = ""
    international_sos_imagen_bytes: bytes | None = None
    international_sos_imagen_mime: str | None = None

    # pasajeros (listas)
    pasajeros_ida: list[str] = field(default_factory=list)

    # 5. Vuelta
    vuelta_hora_salida: str = ""
    vuelta_hora_llegada: str = ""
    vuelta_fecha_salida: date | None = None
    vuelta_fecha_llegada: date | None = None
    pasajeros_vuelta: list[str] = field(default_factory=list)
    paradas_vuelta: list[Stop] = field(default_factory=list)

    # 6. Aprobación
    firma_elabora: str = ""
    firma_conductor_1: str = ""
    firma_conductor_2: str = ""
    firma_aprueba_1: str = ""
    firma_aprueba_2: str = ""
    fecha_firma: date | None = None

    # Imagen ruta
    ruta_imagen_bytes: bytes | None = None
    ruta_imagen_mime: str | None = None
    ruta_vuelta_imagen_bytes: bytes | None = None
    ruta_vuelta_imagen_mime: str | None = None


def _safe_pdf_filename_segment(text: str, max_len: int = 80) -> str:
    """Fragmento seguro para nombre de archivo (Windows: sin \\ / : * ? \" < > |)."""
    s = (text or "").strip()
    if not s:
        return "Viaje"
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_.")
    if not s:
        return "Viaje"
    return s[:max_len]


def _default_plan_pdf_filename(empresa: str, empresa_nombre: str, destino: str, d: date) -> str:
    """Patrón sugerido: PV_(empresa del viaje | destino | nombre en encabezado)_fecha.pdf."""
    mid = (
        (empresa or "").strip()
        or (destino or "").strip()
        or (empresa_nombre or "").strip()
    )
    slug = _safe_pdf_filename_segment(mid)
    return f"PV_{slug}_{d.isoformat()}.pdf"


def _ensure_stop_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Asegura columnas esperadas para editar paradas."""
    if df is None:
        df = pd.DataFrame()
    out = df.copy()
    if "Motivo" not in out.columns:
        out["Motivo"] = ""
    if "Motivo (Otros)" in out.columns:
        # Compatibilidad con sesiones anteriores: unifica texto libre en la celda principal.
        out["Motivo"] = out["Motivo"].fillna("").astype(str).str.strip()
        otros = out["Motivo (Otros)"].fillna("").astype(str).str.strip()
        mask = out["Motivo"].eq("") & otros.ne("")
        out.loc[mask, "Motivo"] = otros[mask]
        out = out.drop(columns=["Motivo (Otros)"])
    return out


def _init_state():
    if "paradas_ida" not in st.session_state:
        st.session_state.paradas_ida = pd.DataFrame(
            [{"N°": 1, "Lugar / Ciudad": "", "Motivo": "", "Tiempo (min)": ""}]
        )
    else:
        st.session_state.paradas_ida = _ensure_stop_columns(st.session_state.paradas_ida)
    if "paradas_vuelta" not in st.session_state:
        st.session_state.paradas_vuelta = pd.DataFrame(
            columns=["N°", "Lugar / Ciudad", "Motivo", "Tiempo (min)"]
        )
    else:
        st.session_state.paradas_vuelta = _ensure_stop_columns(st.session_state.paradas_vuelta)
    if "paradas_ida_rows" not in st.session_state:
        migrated_ida = _stops_df_to_ui_rows(_ensure_stop_columns(st.session_state.paradas_ida))
        st.session_state.paradas_ida_rows = migrated_ida if migrated_ida else [_new_parada_row(1)]
    if "paradas_vuelta_rows" not in st.session_state:
        st.session_state.paradas_vuelta_rows = _stops_df_to_ui_rows(
            _ensure_stop_columns(st.session_state.paradas_vuelta)
        )
    if "shared_pdf_folder" not in st.session_state:
        st.session_state.shared_pdf_folder = os.environ.get(PDF_OUTPUT_DIR_ENV, "")
    if "ubicaciones_df" not in st.session_state:
        st.session_state.ubicaciones_df = None
    if "ubicaciones_source" not in st.session_state:
        st.session_state.ubicaciones_source = "Sin datos"


def _col_norm(c: Any) -> str:
    return (
        str(c)
        .strip()
        .upper()
        .replace("Ó", "O")
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ú", "U")
    )


def _pick_columna_geografia(df: pd.DataFrame, tipo: str) -> str | None:
    """
    tipo: PROVINCIA | CANTON | PARROQUIA
    Prioriza columnas con NOMBRE / DESCRIPCIÓN sobre códigos (COD_*).
    Compatible con tablas tipo INEC / CODIFICACIÓN DPA (DPA_DESPRO, DPA_DESCAN, DPA_DESPAR).
    """
    cols = list(df.columns)

    # Nomenclatura estándar INEC en Excel CODIFICACIÓN (nombres geográficos, no códigos)
    dpa_tokens = {
        "PROVINCIA": ("DESPRO", "NOMPRO", "NOM_PRO", "DPA_DESPRO"),
        "CANTON": ("DESCAN", "NOMCAN", "NOM_CAN", "DPA_DESCAN", "DPA_DESCANT"),
        "PARROQUIA": ("DESPAR", "NOMPAR", "NOM_PAR", "DPA_DESPAR"),
    }
    for tok in dpa_tokens[tipo]:
        for c in cols:
            u = _col_norm(c)
            if tok in u:
                return c

    def score(col_name: str) -> tuple[int, int]:
        u = _col_norm(col_name)
        s = 0
        if "NOMBRE" in u or "DESCRIP" in u:
            s += 10
        if "COD" in u or u.startswith("DPA") and "COD" in u:
            s -= 5
        return (s, -len(u))

    candidates: list[str] = []
    for c in cols:
        u = _col_norm(c)
        if tipo == "PROVINCIA" and "PROVINCIA" in u:
            candidates.append(c)
        elif tipo == "CANTON" and "CANT" in u and "CANTIDAD" not in u:
            candidates.append(c)
        elif tipo == "PARROQUIA" and "PARROQ" in u:
            candidates.append(c)
    if not candidates:
        return None
    candidates.sort(key=lambda x: score(x), reverse=True)
    return candidates[0]


def _normalize_ubicaciones_df(df: pd.DataFrame) -> pd.DataFrame | None:
    """Provincia / cantón / parroquia: nombres de columna flexibles (CSV o Excel INEC)."""
    if df is None or df.empty:
        return None
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    pcol = _pick_columna_geografia(df, "PROVINCIA")
    ccol = _pick_columna_geografia(df, "CANTON")
    rcol = _pick_columna_geografia(df, "PARROQUIA")

    # Respaldo: nombres exactos simples
    if not pcol or not ccol or not rcol:
        col_map: dict[str, str] = {}
        for c in df.columns:
            x = _col_norm(c)
            if x == "PROVINCIA":
                col_map["PROVINCIA"] = c
            elif x in ("CANTON", "CANTÓN"):
                col_map["CANTON"] = c
            elif x.startswith("PARROQ"):
                col_map["PARROQUIA"] = c
        if len(col_map) == 3:
            pcol, ccol, rcol = col_map["PROVINCIA"], col_map["CANTON"], col_map["PARROQUIA"]

    if not pcol or not ccol or not rcol:
        return None

    out = df[[pcol, ccol, rcol]].copy()
    out.columns = ["PROVINCIA", "CANTON", "PARROQUIA"]
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip().replace({"nan": "", "None": ""})
    out = out[(out["PROVINCIA"] != "") & (out["CANTON"] != "") & (out["PARROQUIA"] != "")]
    out = out.drop_duplicates().sort_values(["PROVINCIA", "CANTON", "PARROQUIA"])
    return out if not out.empty else None


def _ruta_excel_codificacion() -> str | None:
    """Busca el Excel de codificación en la carpeta de trabajo (nombre exacto o *CODIFIC*2025*.xlsx)."""
    for name in (UBICACIONES_XLSX_PRIMARY, UBICACIONES_XLSX_ALT):
        if os.path.isfile(name):
            return name
    # Windows: el nombre con tilde en disco puede no coincidir con el literal; se busca por listdir
    try:
        for fn in os.listdir("."):
            if not str(fn).lower().endswith(".xlsx"):
                continue
            u = str(fn).upper().replace("Ó", "O").replace("Á", "A")
            if "2025" not in u:
                continue
            if "CODIFIC" in u and os.path.isfile(fn):
                return fn
    except Exception:
        pass
    try:
        import glob

        for p in glob.glob("*.xlsx"):
            u = os.path.basename(p).upper().replace("Ó", "O")
            if "2025" in u and "CODIFIC" in u:
                return p
    except Exception:
        pass
    return None


def _sheet_excel_ubicaciones() -> str | int:
    s = os.environ.get(UBICACIONES_SHEET_ENV, "0")
    if s.isdigit():
        return int(s)
    return s


def _read_excel_kw() -> dict[str, Any]:
    return {
        "sheet_name": _sheet_excel_ubicaciones(),
        "dtype": str,
        "engine": "openpyxl",
        "keep_default_na": False,
    }


def _leer_excel_ubicaciones_con_header(path_or_bytes: str | bytes, *, from_path: bool) -> pd.DataFrame:
    """
    El archivo CODIFICACIÓN INEC suele tener títulos en la fila 2 (header=1).
    Se prueba header 1, 0, 2, 3 hasta que _normalize reconozca provincia/cantón/parroquia.
    """
    kw = _read_excel_kw()
    for header in (1, 0, 2, 3):
        try:
            if from_path:
                raw = pd.read_excel(path_or_bytes, header=header, **kw)
            else:
                raw = pd.read_excel(io.BytesIO(path_or_bytes), header=header, **kw)
            if _normalize_ubicaciones_df(raw) is not None:
                return raw
        except Exception:
            pass
    if from_path:
        return pd.read_excel(path_or_bytes, header=1, **kw)
    return pd.read_excel(io.BytesIO(path_or_bytes), header=1, **kw)


def _leer_excel_ubicaciones(path: str) -> pd.DataFrame:
    return _leer_excel_ubicaciones_con_header(path, from_path=True)


def _leer_excel_ubicaciones_bytes(data: bytes) -> pd.DataFrame:
    return _leer_excel_ubicaciones_con_header(data, from_path=False)


@st.cache_data(show_spinner=False)
def _cached_norm_ubicaciones_excel(path: str, _mtime: float) -> pd.DataFrame | None:
    """Evita releer y normalizar el Excel INEC en cada clic (lo más pesado al abrir)."""
    try:
        raw = _leer_excel_ubicaciones(path)
        return _normalize_ubicaciones_df(raw)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def _cached_norm_ubicaciones_csv(path: str, _mtime: float) -> pd.DataFrame | None:
    try:
        raw = pd.read_csv(path, dtype=str, keep_default_na=False)
        return _normalize_ubicaciones_df(raw)
    except Exception:
        return None


def _load_ubicaciones_desde_archivo_local() -> None:
    if st.session_state.get("ubicaciones_df") is not None:
        return
    _ubi_id = _gsheet_ubicaciones_spreadsheet_id()
    if _ubi_id:
        try:
            ws = _open_worksheet(_ubi_id, _ubicaciones_gsheet_worksheet_spec())
            vals = ws.get_all_values()
            norm = _ubicaciones_norm_from_sheet_matrix(vals)
            if norm is not None:
                st.session_state.ubicaciones_df = norm
                st.session_state.ubicaciones_source = (
                    f"Google Sheets ({_ubi_id}) · pestaña `{_ubicaciones_gsheet_worksheet_spec()}`"
                )
                return
            norm = _ubicaciones_merge_innec_three_tabs(_ubi_id)
            if norm is not None:
                st.session_state.ubicaciones_df = norm
                st.session_state.ubicaciones_source = (
                    f"Google Sheets ({_ubi_id}) · pestañas "
                    f"`{_ubicaciones_sheet_title_provinces()}` + `{_ubicaciones_sheet_title_cantones()}` + "
                    f"`{_ubicaciones_sheet_title_parroquias()}`"
                )
                return
            st.session_state.ubicaciones_source = (
                f"Error: Google Sheets ({_ubi_id}) · no se pudo leer una sola tabla (prov/cant/parr) "
                f"ni unir pestañas INEC. Revisa columnas DPA_* o nombres en "
                f"`{_ubicaciones_gsheet_worksheet_spec()}` / PROVINCIAS / CANTONES / PARROQUIAS."
            )
            return
        except Exception:
            st.session_state.ubicaciones_source = (
                f"Error: no se pudo leer Google Sheets ({_ubi_id})"
            )
            return
    xlsx_path = _ruta_excel_codificacion()
    if xlsx_path and os.path.isfile(xlsx_path):
        norm = _cached_norm_ubicaciones_excel(xlsx_path, os.path.getmtime(xlsx_path))
        if norm is not None:
            st.session_state.ubicaciones_df = norm
            st.session_state.ubicaciones_source = f"Excel local ({os.path.basename(xlsx_path)})"
            return
    if not os.path.isfile(UBICACIONES_CSV_PATH):
        return
    norm = _cached_norm_ubicaciones_csv(UBICACIONES_CSV_PATH, os.path.getmtime(UBICACIONES_CSV_PATH))
    if norm is not None:
        st.session_state.ubicaciones_df = norm
        st.session_state.ubicaciones_source = f"CSV local ({UBICACIONES_CSV_PATH})"


def _ubicacion_campo(df: pd.DataFrame | None, etiqueta: str, key_prefix: str) -> str:
    """
    Devuelve texto para origen/destino. Con BD: cascada provincia/cantón/parroquia.
    Sin BD: texto libre. Debe estar fuera de st.form para que los desplegables reaccionen al instante.
    """
    if df is None or df.empty:
        return st.text_input(
            etiqueta,
            value="",
            placeholder="Ej.: Tonsupa / Esmeraldas",
            key=f"loc_{key_prefix}_manual",
        )
    provs = ["—"] + sorted(df["PROVINCIA"].dropna().unique().tolist())
    prov = st.selectbox(f"{etiqueta} — Provincia", provs, key=f"loc_{key_prefix}_prov")
    sub = df[df["PROVINCIA"] == prov] if prov != "—" else df.iloc[0:0]
    cants = ["—"] + sorted(sub["CANTON"].dropna().unique().tolist()) if not sub.empty else ["—"]
    cant = st.selectbox(f"{etiqueta} — Cantón", cants, key=f"loc_{key_prefix}_cant")
    sub2 = sub[sub["CANTON"] == cant] if cant != "—" and not sub.empty else df.iloc[0:0]
    parrs = ["—"] + sorted(sub2["PARROQUIA"].dropna().unique().tolist()) if not sub2.empty else ["—"]
    parr = st.selectbox(f"{etiqueta} — Parroquia", parrs, key=f"loc_{key_prefix}_parr")
    if prov != "—" and cant != "—" and parr != "—":
        return f"{parr}, {cant} — {prov}"
    return ""


def _df_to_stops(df: pd.DataFrame) -> list[Stop]:
    if df is None or df.empty:
        return []
    out: list[Stop] = []
    for _, r in df.iterrows():
        n_raw = str(r.get("N°") or "").strip()
        if not n_raw:
            continue
        try:
            n = int(float(n_raw))
        except Exception:
            continue
        if n <= 0:
            continue
        out.append(
            Stop(
                n=n,
                lugar=str(r.get("Lugar / Ciudad") or "").strip(),
                motivo=str(r.get("Motivo") or "").strip(),
                tiempo_min=str(r.get("Tiempo (min)") or "").strip(),
            )
        )
    return out


def _parada_motivo_catalog_index(motivo_catalogo: str) -> int:
    try:
        return STOP_MOTIVO_OPTIONS.index(motivo_catalogo)
    except ValueError:
        return 0


def _new_parada_row(n: int) -> dict[str, Any]:
    return {
        "id": str(uuid.uuid4()),
        "n": n,
        "lugar": "",
        "motivo_catalogo": STOP_MOTIVO_OPTIONS[0],
        "motivo_libre": "",
        "tiempo_min": STOP_TIEMPO_MIN_OPTIONS[0],
    }


def _stops_df_to_ui_rows(df: pd.DataFrame) -> list[dict[str, Any]]:
    df = _ensure_stop_columns(df)
    if df.empty:
        return []
    out: list[dict[str, Any]] = []
    for _, r in df.iterrows():
        n_raw = str(r.get("N°") or "").strip()
        try:
            n = int(float(n_raw)) if n_raw else len(out) + 1
        except Exception:
            n = len(out) + 1
        motivo_final = str(r.get("Motivo") or "").strip()
        if motivo_final in STOP_MOTIVO_OPTIONS:
            cat, libre = motivo_final, ""
        else:
            cat, libre = STOP_MOTIVO_OPTIONS[0], motivo_final
        out.append(
            {
                "id": str(uuid.uuid4()),
                "n": n,
                "lugar": str(r.get("Lugar / Ciudad") or "").strip(),
                "motivo_catalogo": cat,
                "motivo_libre": libre,
                "tiempo_min": _coerce_tiempo_min(r.get("Tiempo (min)")),
            }
        )
    return out


def _paradas_widget_state_to_df(prefix: str, rows: list[dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=["N°", "Lugar / Ciudad", "Motivo", "Tiempo (min)"])
    recs: list[dict[str, Any]] = []
    for row in rows:
        rid = row["id"]
        n_raw = st.session_state.get(f"{prefix}_n_{rid}", row.get("n", ""))
        try:
            n = int(float(n_raw))
        except (TypeError, ValueError):
            n = 0
        lugar = str(st.session_state.get(f"{prefix}_lugar_{rid}", row.get("lugar", "")) or "").strip()
        cat = st.session_state.get(
            f"{prefix}_cat_{rid}", row.get("motivo_catalogo", STOP_MOTIVO_OPTIONS[0])
        )
        libre = str(
            st.session_state.get(f"{prefix}_libre_{rid}", row.get("motivo_libre", "")) or ""
        ).strip()
        tiempo_raw = st.session_state.get(
            f"{prefix}_tiempo_{rid}", row.get("tiempo_min", STOP_TIEMPO_MIN_OPTIONS[0])
        )
        tiempo = str(_coerce_tiempo_min(tiempo_raw))
        cat_s = str(cat or "").strip()
        motivo = libre if libre else cat_s
        recs.append(
            {"N°": n, "Lugar / Ciudad": lugar, "Motivo": motivo, "Tiempo (min)": tiempo}
        )
    return pd.DataFrame(recs)


def _render_paradas_form_block(
    *,
    subheader: str,
    caption: str,
    rows_key: str,
    prefix: str,
    count_label: str,
    min_rows: int,
    max_rows: int = 40,
) -> None:
    st.subheader(subheader)
    st.caption(caption)
    rows: list[dict[str, Any]] = st.session_state[rows_key]
    n_paradas = st.number_input(
        count_label,
        min_value=min_rows,
        max_value=max_rows,
        value=max(min_rows, len(rows)) if min_rows > 0 else len(rows),
        step=1,
        key=f"{prefix}_paradas_count",
        help="Aumenta o disminuye cuántas filas de paradas se muestran (las últimas se quitan al bajar el número).",
    )
    n_paradas = int(n_paradas)
    while len(rows) > n_paradas:
        rows.pop()
    while len(rows) < n_paradas:
        rows.append(_new_parada_row(len(rows) + 1))

    for idx, row in enumerate(rows):
        rid = row["id"]
        st.markdown(f"**Parada {idx + 1}**")
        c0, c1, c2, c3, c4 = st.columns([0.55, 1.65, 1.45, 1.85, 0.85], gap="small")
        with c0:
            st.number_input(
                "N°",
                min_value=1,
                value=int(row.get("n", idx + 1) or idx + 1),
                step=1,
                key=f"{prefix}_n_{rid}",
            )
        with c1:
            st.text_input(
                "Lugar / Ciudad",
                value=str(row.get("lugar", "")),
                key=f"{prefix}_lugar_{rid}",
            )
        with c2:
            st.selectbox(
                "Motivo (lista)",
                options=STOP_MOTIVO_OPTIONS,
                index=_parada_motivo_catalog_index(str(row.get("motivo_catalogo", ""))),
                key=f"{prefix}_cat_{rid}",
            )
        with c3:
            st.text_input(
                "O escribir otro motivo",
                value=str(row.get("motivo_libre", "")),
                key=f"{prefix}_libre_{rid}",
                placeholder="Si escribes aquí, sustituye a la opción de la lista",
            )
        with c4:
            st.selectbox(
                "Tiempo (min)",
                options=STOP_TIEMPO_MIN_OPTIONS,
                index=STOP_TIEMPO_MIN_OPTIONS.index(_coerce_tiempo_min(row.get("tiempo_min"))),
                key=f"{prefix}_tiempo_{rid}",
            )


def main():
    st.set_page_config(page_title="Plan de Gestión de Viaje ALS Ecuador", layout="wide")

    _init_state()
    _load_ubicaciones_desde_archivo_local()

    people_df = _load_people()
    people_opts = _person_options(people_df)

    vehicles_df = _load_vehicles()
    vehicles_opts = _vehicle_options(vehicles_df)

    st.markdown(
        """
<style>
/* Ajuste visual / responsivo */
div[data-testid="stHorizontalBlock"] { gap: 1rem; flex-wrap: wrap; }
/* Inputs con alto consistente */
div[data-baseweb="input"] > div { min-height: 42px; }
</style>
        """.strip(),
        unsafe_allow_html=True,
    )

    colA, colB = st.columns([2, 1], gap="large")

    data = PlanData()

    with colB:
        if st.button("Limpiar formulario"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        st.subheader("Encabezado")
        data.empresa_nombre = st.text_input("Nombre de la empresa", value=DEFAULT_COMPANY_NAME)
        data.doc_code = st.text_input("Código documento", value=DEFAULT_DOC_CODE)
        data.doc_rev = st.text_input("Revisión", value=DEFAULT_REV)
        data.doc_date = st.text_input("Fecha documento", value=DEFAULT_DOC_DATE)

        # Logo: si existe en assets, úsalo por defecto; si el usuario sube uno, guardarlo para próximos PDFs.
        try:
            if os.path.exists(DEFAULT_LOGO_PATH) and data.empresa_logo_bytes is None:
                with open(DEFAULT_LOGO_PATH, "rb") as f:
                    data.empresa_logo_bytes = f.read()
                data.empresa_logo_mime = "image/png"
        except Exception:
            pass

        logo = st.file_uploader("Logo (opcional)", type=["png", "jpg", "jpeg"])
        if logo is not None:
            logo_bytes = logo.read()
            data.empresa_logo_bytes = logo_bytes
            data.empresa_logo_mime = logo.type
            try:
                os.makedirs(os.path.dirname(DEFAULT_LOGO_PATH), exist_ok=True)
                with open(DEFAULT_LOGO_PATH, "wb") as f:
                    f.write(logo_bytes)
            except Exception:
                # Si no se puede guardar, igual lo usamos para este PDF.
                pass
            st.image(data.empresa_logo_bytes, caption="Logo cargado", width="stretch")

        st.subheader("Datos rápidos (catálogo)")
        if st.button(
            "Recargar BD y catálogos",
            help="Ejecuta esto después de editar las hojas de Google Sheets (BD / camionetas). "
            "Limpia la caché y vuelve a leer desde la nube.",
        ):
            st.cache_data.clear()
            try:
                st.cache_resource.clear()
            except Exception:
                pass
            st.rerun()
        st.write(f"Personas en BD: **{len(people_df)}**")
        st.caption(
            f"Fuente: **Google Sheets** · libro `{_gsheet_bd_spreadsheet_id()}` · pestaña "
            f"`{_bd_worksheet_spec()}` (`st.secrets` / `{PLAN_BD_WORKSHEET_ENV}`)."
        )
        st.caption(
            f"Credenciales: **st.secrets** (producción) o **`{CREDENTIALS_JSON_PATH}`** / "
            "`GOOGLE_APPLICATION_CREDENTIALS` (local). "
            f"Tras editar la hoja, usa **Recargar BD**. "
            f"Hasta **{_max_tecnicos_viaje()}** técnicos por viaje (`{MAX_TECNICOS_ENV}`, 1–20)."
        )
        st.write(f"Vehículos en BD: **{len(vehicles_df)}**")
        st.caption(
            f"Fuente: **Google Sheets** · libro `{_gsheet_vehicles_spreadsheet_id()}` · pestaña "
            f"`{_vehicles_worksheet_spec()}` (`st.secrets` / `{PLAN_VEHICLES_WORKSHEET_ENV}`)."
        )

        st.subheader("Ubicaciones (opcional)")
        _xlsx_here = _ruta_excel_codificacion()
        _ubi_cfg_id = _gsheet_ubicaciones_spreadsheet_id()
        if not _ubi_cfg_id:
            st.warning(
                "La app **no está leyendo** `PLAN_UBICACIONES_SPREADSHEET_ID` (sale vacío). "
                "Por eso se usa Excel/CSV local si existe. Revisa **App settings → Secrets**: la clave debe estar "
                "en el TOML (puede ser al inicio del archivo o dentro de un bloque `[...]`), pulsa **Save changes** "
                "y luego **Reboot app**."
            )
        if _ubi_cfg_id:
            st.caption(
                "Modo **solo Google Sheets**: con **`PLAN_UBICACIONES_SPREADSHEET_ID`** definido no se usa el Excel "
                f"ni el CSV de la carpeta del servidor. Pestaña principal: **`{_ubicaciones_gsheet_worksheet_spec()}`** "
                f"(`{PLAN_UBICACIONES_WORKSHEET_ENV}`). Si esa pestaña es tipo **CODIGOS** y no sirve para listas, "
                "la app intenta unir automáticamente **PROVINCIAS + CANTONES + PARROQUIAS** (nombres configurables con "
                f"`{PLAN_UBICACIONES_SHEET_PROVINCIAS_ENV}`, etc.). Comparte el Sheet con la cuenta de servicio."
            )
            st.caption(f"ID de ubicaciones activo (últimos 10 caracteres): `…{_ubi_cfg_id[-10:]}`")
        else:
            st.caption(
                f"Opcional: define **`PLAN_UBICACIONES_SPREADSHEET_ID`** para cargar DPA desde Google Sheets. "
                f"Si no, se usa **`{UBICACIONES_XLSX_PRIMARY}`** / **`{UBICACIONES_CSV_PATH}`** en esta carpeta. "
                f"También puedes subir un **.xlsx** o **.csv** aquí. "
                f"Hoja Excel local: variable `{UBICACIONES_SHEET_ENV}` (índice o nombre)."
            )
        if _xlsx_here and not _ubi_cfg_id:
            st.write(f"Detectado en disco: `{_xlsx_here}`")
        elif _xlsx_here and _ubi_cfg_id:
            st.caption(
                f"(Hay un archivo `{os.path.basename(_xlsx_here)}` en el despliegue, pero **no se usa** mientras exista "
                "`PLAN_UBICACIONES_SPREADSHEET_ID`.)"
            )
        ubi_upload = st.file_uploader(
            "Subir Excel (CODIFICACIÓN) o CSV de ubicaciones",
            type=["csv", "xlsx"],
            key="ubicaciones_file_upload",
        )
        if ubi_upload is not None:
            try:
                name_l = ubi_upload.name.lower()
                raw_data = ubi_upload.read()
                if name_l.endswith(".xlsx"):
                    raw = _leer_excel_ubicaciones_bytes(raw_data)
                else:
                    raw = pd.read_csv(io.BytesIO(raw_data), dtype=str, keep_default_na=False)
                norm = _normalize_ubicaciones_df(raw)
                if norm is not None:
                    st.session_state.ubicaciones_df = norm
                    st.session_state.ubicaciones_source = f"Archivo subido ({ubi_upload.name})"
                    st.success(f"Cargado en memoria: **{len(norm)}** filas.")
                else:
                    st.error(
                        "No se detectaron columnas de provincia / cantón / parroquia. "
                        "Revisa que el archivo tenga nombres o códigos DPA reconocibles."
                    )
            except Exception as ex:
                st.error(f"No se pudo leer el archivo: {ex}")
        ubi_n = st.session_state.get("ubicaciones_df")
        ubi_src = (st.session_state.get("ubicaciones_source") or "Sin datos").strip()
        st.caption(f"Fuente activa de ubicaciones: **{ubi_src}**")
        if ubi_src.startswith("Error:"):
            st.error(
                f"{ubi_src} · Revisa permisos del Sheet, el ID y que la pestaña tenga columnas "
                "reconocibles (provincia / cantón / parroquia o DPA_DESPRO, DPA_DESCAN, DPA_DESPAR)."
            )
        if ubi_n is not None and not getattr(ubi_n, "empty", True):
            st.write(f"Filas activas: **{len(ubi_n)}**")
        if st.button("Usar solo texto (sin lista)", help="Quita la lista y vuelve a escribir origen/destino a mano"):
            st.session_state.ubicaciones_df = None
            st.session_state.ubicaciones_source = "Texto manual (sin lista)"
            st.rerun()

    ubi_df: pd.DataFrame | None = st.session_state.get("ubicaciones_df")

    with colA:
        header_left, header_right = st.columns([1, 3], vertical_alignment="center")
        with header_left:
            if data.empresa_logo_bytes:
                st.image(data.empresa_logo_bytes, width="stretch")
        with header_right:
            st.title(f"PLAN DE GESTIÓN DE VIAJE — {data.empresa_nombre}".strip())
            st.caption(
                "Formulario para viajes >300 km (ida y vuelta), >4 horas, o con riesgos particulares. "
                "Completa y luego exporta a PDF."
            )

        # Fuera del st.form: al cambiar un selectbox el script debe re-ejecutarse para autollenar
        # celular/cédula. Dentro de un form eso no ocurre hasta presionar "Generar PDF".
        st.subheader("1. Datos generales")
        col_tecnicos, col_meta = st.columns([3, 2])
        with col_tecnicos:
            _mx = _max_tecnicos_viaje()
            n_tecnicos = st.number_input(
                "Cantidad de técnicos en el viaje",
                min_value=1,
                max_value=_mx,
                value=min(2, _mx),
                step=1,
                help=f"Un selector por cada técnico (máx. {_mx} en esta instalación).",
            )
            n_tecnicos = int(n_tecnicos)
            conductores_nombres: list[str] = []
            conductores_ced: list[str] = []
            conductores_tel: list[str] = []
            tec_opts = _opciones_con_manual(people_opts)
            for i in range(n_tecnicos):
                if i:
                    st.divider()
                sel = st.selectbox(
                    f"Técnico {i + 1}",
                    tec_opts,
                    index=0,
                    key=f"conductor_sel_{i}",
                )
                manual = sel == PERSON_SEL_OTRO_MANUAL
                m_n = st.text_input(
                    "Nombre completo (si «Otro»)",
                    key=f"conductor_manual_nom_{i}",
                    disabled=not manual,
                )
                m_id = st.text_input(
                    "Cédula (si «Otro»)",
                    key=f"conductor_manual_ced_{i}",
                    disabled=not manual,
                )
                m_cel = st.text_input(
                    "Celular (si «Otro»)",
                    key=f"conductor_manual_cel_{i}",
                    disabled=not manual,
                )
                d = _person_lookup(people_df, sel) if not manual else {"name": "", "cel": "", "id": ""}
                c_a, c_b = st.columns(2)
                with c_a:
                    st.caption("Celular")
                    cel_show = (m_cel or "").strip() if manual else (d.get("cel") or "")
                    st.markdown(f"**{cel_show or '—'}**")
                with c_b:
                    st.caption("Cédula")
                    id_show = (m_id or "").strip() if manual else (d.get("id") or "")
                    st.markdown(f"**{id_show or '—'}**")
                if manual:
                    nom = (m_n or "").strip()
                    if nom:
                        conductores_nombres.append(nom)
                        conductores_ced.append((m_id or "").strip())
                        conductores_tel.append((m_cel or "").strip())
                elif sel != "—":
                    conductores_nombres.append(d["name"])
                    conductores_ced.append(d["id"])
                    conductores_tel.append(d["cel"])
            data.conductores = conductores_nombres
            data.cedulas_conductores = conductores_ced
            data.celulares_conductores = conductores_tel
        with col_meta:
            data.fecha_elab = st.date_input("Fecha de Elaboración", value=date.today())
            data.cargo = st.text_input("Cargo / Posición", value="Técnico de Operaciones")

        st.markdown("##### Origen y destino")
        st.caption(
            "Con un CSV de provincia / cantón / parroquia verás listas en cascada. "
            "Sin archivo, puedes escribir el texto a mano."
        )
        u_o, u_d = st.columns(2)
        with u_o:
            data.origen = _ubicacion_campo(ubi_df, "Punto de origen", "origen")
        with u_d:
            data.destino = _ubicacion_campo(ubi_df, "Destino", "destino")

        # Vehículo y emergencias fuera del form: el selector debe actualizar placa/teléfono al instante
        g3, g4 = st.columns(2)
        with g3:
            veh_sel = st.selectbox("Vehículo (desde BD) (opcional)", vehicles_opts, index=0)
            veh = _vehicle_lookup(vehicles_df, veh_sel)
            data.placa = st.text_input("Placa del Vehículo", value=veh["placa"])
            data.tipo_vehiculo = st.text_input("Tipo de Vehículo", value=veh["tipo"])
        with g4:
            data.modelo_anio = st.text_input("Modelo / Año", value=veh["modelo_anio"])

        st.caption(
            "Contactos de emergencia: atajos **Santiago Montalvan** / **David Solano** (datos desde la BD), "
            "**cualquier persona de la BD**, o **Otro** para escribir nombre y teléfono a mano."
        )
        e1_otro_nombre, e1_otro_tel = "", ""
        e2_otro_nombre, e2_otro_tel = "", ""
        e1_bd_pick, e2_bd_pick = "", ""
        emc1, emc2 = st.columns(2)
        with emc1:
            e1_choice = st.selectbox("Contacto de Emergencia (1)", EMERGENCY_SELECT_OPTIONS, index=0, key="em1_choice")
            if e1_choice in EMERGENCY_PRESET_TO_BD_NAME:
                _p1 = _person_lookup(people_df, EMERGENCY_PRESET_TO_BD_NAME[e1_choice])
                st.caption("Teléfono (Emergencia 1)")
                st.markdown(f"**{_p1['cel'] or '—'}**")
                st.caption("Cédula (Emergencia 1)")
                st.markdown(f"**{_p1['id'] or '—'}**")
            elif e1_choice == EMERGENCY_FROM_BD:
                bd_em1 = [p for p in people_opts if p != "—"]
                e1_bd_pick = st.selectbox(
                    "Persona en BD (Emergencia 1)",
                    ["—"] + bd_em1,
                    index=0,
                    key="em1_bd_pick",
                )
                _p1b = _person_lookup(people_df, e1_bd_pick) if e1_bd_pick != "—" else {"cel": "", "id": ""}
                st.caption("Teléfono (Emergencia 1)")
                st.markdown(f"**{_p1b['cel'] or '—'}**")
                st.caption("Cédula (Emergencia 1)")
                st.markdown(f"**{_p1b.get('id') or '—'}**")
            elif e1_choice == EMERGENCY_OTRO_TEXTO:
                e1_otro_nombre = st.text_input("Nombre completo (Emergencia 1)", key="em1_otro_nombre")
                e1_otro_tel = st.text_input("Teléfono (Emergencia 1)", key="em1_otro_tel")
        with emc2:
            e2_choice = st.selectbox("Contacto de Emergencia (2)", EMERGENCY_SELECT_OPTIONS, index=0, key="em2_choice")
            if e2_choice in EMERGENCY_PRESET_TO_BD_NAME:
                _p2 = _person_lookup(people_df, EMERGENCY_PRESET_TO_BD_NAME[e2_choice])
                st.caption("Teléfono (Emergencia 2)")
                st.markdown(f"**{_p2['cel'] or '—'}**")
                st.caption("Cédula (Emergencia 2)")
                st.markdown(f"**{_p2['id'] or '—'}**")
            elif e2_choice == EMERGENCY_FROM_BD:
                bd_em2 = [p for p in people_opts if p != "—"]
                e2_bd_pick = st.selectbox(
                    "Persona en BD (Emergencia 2)",
                    ["—"] + bd_em2,
                    index=0,
                    key="em2_bd_pick",
                )
                _p2b = _person_lookup(people_df, e2_bd_pick) if e2_bd_pick != "—" else {"cel": "", "id": ""}
                st.caption("Teléfono (Emergencia 2)")
                st.markdown(f"**{_p2b['cel'] or '—'}**")
                st.caption("Cédula (Emergencia 2)")
                st.markdown(f"**{_p2b.get('id') or '—'}**")
            elif e2_choice == EMERGENCY_OTRO_TEXTO:
                e2_otro_nombre = st.text_input("Nombre completo (Emergencia 2)", key="em2_otro_nombre")
                e2_otro_tel = st.text_input("Teléfono (Emergencia 2)", key="em2_otro_tel")

        (
            data.emergencia_1,
            data.tel_emergencia_1,
            data.cedula_emergencia_1,
        ) = _resolve_emergency_contact(
            e1_choice,
            e1_otro_nombre if e1_choice == EMERGENCY_OTRO_TEXTO else "",
            e1_otro_tel if e1_choice == EMERGENCY_OTRO_TEXTO else "",
            people_df,
            bd_pick_name=e1_bd_pick if e1_choice == EMERGENCY_FROM_BD else "",
        )
        (
            data.emergencia_2,
            data.tel_emergencia_2,
            data.cedula_emergencia_2,
        ) = _resolve_emergency_contact(
            e2_choice,
            e2_otro_nombre if e2_choice == EMERGENCY_OTRO_TEXTO else "",
            e2_otro_tel if e2_choice == EMERGENCY_OTRO_TEXTO else "",
            people_df,
            bd_pick_name=e2_bd_pick if e2_choice == EMERGENCY_FROM_BD else "",
        )

        # Fuera del form: el pegado desde portapapeles no es confiable dentro de st.form.
        st.subheader("Adjuntos de imagen")
        st.caption(
            "Puedes pegar un screenshot con Ctrl+V (Chrome/Edge) o subir archivo. "
            "Estas imágenes se guardan para el PDF al presionar Generar PDF."
        )
        sos_img_bytes = _image_bytes_input(
            upload_label="Captura de pantalla o imagen desde la app (PNG/JPG)",
            upload_key="international_sos_uploader",
            paste_label="Pegar screenshot (International SOS)",
            allowed_types=["png", "jpg", "jpeg"],
        )
        if sos_img_bytes:
            st.image(sos_img_bytes, caption="Vista previa International SOS", width="stretch")
        ruta_bytes = _image_bytes_input(
            upload_label="Sube captura de Google Maps/Waze (IDA) (PNG/JPG)",
            upload_key="ruta_ida_uploader",
            paste_label="Pegar screenshot (Ruta IDA)",
            allowed_types=["png", "jpg", "jpeg"],
        )
        if ruta_bytes:
            st.image(ruta_bytes, caption="Ruta IDA cargada", width="stretch")
        ruta_vuelta_bytes = _image_bytes_input(
            upload_label="Sube captura de Google Maps/Waze (VUELTA) (PNG/JPG)",
            upload_key="ruta_vuelta_uploader",
            paste_label="Pegar screenshot (Ruta VUELTA)",
            allowed_types=["png", "jpg", "jpeg"],
        )
        if ruta_vuelta_bytes:
            st.image(ruta_vuelta_bytes, caption="Ruta VUELTA cargada", width="stretch")

        with st.form("plan_form", clear_on_submit=False):
            st.subheader("2. Planificación de viaje")
            p1, p2 = st.columns(2)
            with p1:
                data.empresa = st.text_input("Empresa", value="")
                data.orden_trabajo = st.text_input("Orden de Trabajo", value="")
                data.fecha_salida = st.date_input("Fecha de Salida", value=date.today(), key="fecha_salida")
                _hs = st.time_input(
                    "Hora de Salida",
                    value=time(17, 0),
                    step=300,
                    key="hora_salida_picker",
                )
                data.hora_salida = _format_hora_plan(_hs)
                data.distancia_km = st.text_input("Distancia Total (km)", value="")
            with p2:
                data.fecha_llegada = st.date_input(
                    "Fecha Estimada de Llegada", value=date.today(), key="fecha_llegada"
                )
                _hl = st.time_input(
                    "Hora Estimada de Llegada",
                    value=time(20, 0),
                    step=300,
                    key="hora_llegada_picker",
                )
                data.hora_llegada = _format_hora_plan(_hl)
                data.duracion_horas = st.text_input("Duración Estimada (horas)", value="")
                condiciones_camino_opts = [
                    "Asfalto",
                    "Lastre",
                    "Mixto",
                    "Tierra",
                    "Caminos de segundo orden",
                    "Otro",
                ]
                condiciones_camino_sel = st.multiselect(
                    "Condiciones del camino",
                    condiciones_camino_opts,
                    default=["Asfalto"],
                    help="Puedes seleccionar una o varias condiciones.",
                )
                data.condiciones_camino = ", ".join(condiciones_camino_sel) if condiciones_camino_sel else "—"

            data.observaciones = st.text_area("Observaciones adicionales", height=90, value="")

            st.markdown("**APP International SOS** *(opcional: texto y/o captura de lo que muestra la app)*")
            data.international_sos_text = st.text_area(
                "Notas o resumen (International SOS)",
                height=80,
                value="",
                placeholder="Ej.: nivel de riesgo, recomendaciones, enlace o comentario breve…",
            )

            data.proposito = st.text_area("Propósito del viaje", height=120, value="")

            _render_paradas_form_block(
                subheader="3. Paradas planificadas (IDA)",
                caption=(
                    "Motivo: elige una opción de la lista o escribe en «O escribir otro motivo» "
                    "(el texto escrito tiene prioridad). Usa el número de paradas para agregar o quitar filas."
                ),
                rows_key="paradas_ida_rows",
                prefix="ida",
                count_label="Número de paradas (IDA)",
                min_rows=1,
            )

            st.subheader("4. Peligros conocidos (marca con X)")
            ph1, ph2, ph3 = st.columns(3)
            with ph1:
                data.peligro_lluvia = st.checkbox("Lluvia")
                data.peligro_niebla = st.checkbox("Niebla")
            with ph2:
                data.peligro_nieve_hielo = st.checkbox("Nieve / hielo")
                data.peligro_nocturna = st.checkbox("Conducción nocturna")
            with ph3:
                data.peligro_carretera_mala = st.checkbox("Carreteras en mal estado")
                data.peligro_delincuencia = st.checkbox("Zona de alta delincuencia")
                data.peligro_accidentes_transito = st.checkbox("Accidentes de tránsito")

            data.otros_peligros = st.text_area("Otros peligros / detalles adicionales", height=90, value="")

            st.subheader("Pasajeros (IDA)")
            pasajeros_ida_sel = st.multiselect(
                "Selecciona pasajeros (IDA) (opcional)",
                options=[p for p in people_opts if p != "—"],
                default=[],
            )
            pasajeros_ida_extra = st.text_input(
                "Otros pasajeros (IDA) no en la lista — separar con coma o una línea por nombre",
                value="",
                key="pasajeros_ida_extra",
            )
            data.pasajeros_ida = _merge_passenger_lists(
                [p for p in pasajeros_ida_sel if p and p != "—"],
                pasajeros_ida_extra,
            )
            if data.pasajeros_ida:
                st.caption(f"Cantidad de pasajeros (IDA): {len(data.pasajeros_ida)}")

            st.subheader("5. Viaje de Vuelta")
            vuelta_con_horas = st.checkbox(
                "Definir horas de vuelta (reloj)",
                value=False,
                key="vuelta_horas_chk",
                help="Si no marcas esto, las horas de vuelta quedan en blanco en el PDF (como antes).",
            )
            v1, v2 = st.columns(2)
            with v1:
                data.vuelta_fecha_salida = st.date_input(
                    "Fecha de salida (VUELTA)", value=date.today(), key="vuelta_fecha_salida"
                )
                if vuelta_con_horas:
                    _vs = st.time_input(
                        "Hora de salida (VUELTA)",
                        value=time(8, 0),
                        step=300,
                        key="vuelta_hora_salida_picker",
                    )
                    data.vuelta_hora_salida = _format_hora_plan(_vs)
                else:
                    data.vuelta_hora_salida = ""
            with v2:
                data.vuelta_fecha_llegada = st.date_input(
                    "Fecha estimada de llegada (VUELTA)", value=date.today(), key="vuelta_fecha_llegada"
                )
                if vuelta_con_horas:
                    _vl = st.time_input(
                        "Hora estimada de llegada (VUELTA)",
                        value=time(18, 0),
                        step=300,
                        key="vuelta_hora_llegada_picker",
                    )
                    data.vuelta_hora_llegada = _format_hora_plan(_vl)
                else:
                    data.vuelta_hora_llegada = ""

            st.subheader("Pasajeros (VUELTA)")
            pasajeros_vuelta_sel = st.multiselect(
                "Selecciona pasajeros (VUELTA) (opcional)",
                options=[p for p in people_opts if p != "—"],
                default=[],
            )
            pasajeros_vuelta_extra = st.text_input(
                "Otros pasajeros (VUELTA) no en la lista — separar con coma o una línea por nombre",
                value="",
                key="pasajeros_vuelta_extra",
            )
            data.pasajeros_vuelta = _merge_passenger_lists(
                [p for p in pasajeros_vuelta_sel if p and p != "—"],
                pasajeros_vuelta_extra,
            )
            if data.pasajeros_vuelta:
                st.caption(f"Cantidad de pasajeros (VUELTA): {len(data.pasajeros_vuelta)}")

            _render_paradas_form_block(
                subheader="Paradas planificadas (VUELTA)",
                caption=(
                    "Misma lógica que en IDA. Pon 0 paradas si no aplica. "
                    "El texto en «O escribir otro motivo» sustituye a la lista si no está vacío."
                ),
                rows_key="paradas_vuelta_rows",
                prefix="vuelta",
                count_label="Número de paradas (VUELTA)",
                min_rows=0,
            )

            st.subheader("6. Aprobación")
            firma_opts = _opciones_con_manual(people_opts)
            ap1, ap2 = st.columns(2)
            with ap1:
                data.firma_elabora = st.text_input(
                    "Firma responsable elaboración plan", value="Cristina Aguirre"
                )
                fc1_sel = st.selectbox("Firma conductor responsable (1)", firma_opts, index=0)
                fc1_man = st.text_input(
                    "Nombre completo (conductor 1)",
                    key="firma_c1_manual",
                    disabled=fc1_sel != PERSON_SEL_OTRO_MANUAL,
                )
                data.firma_conductor_1 = _valor_firma(fc1_sel, fc1_man)
                fc2_sel = st.selectbox("Firma conductor responsable (2)", firma_opts, index=0)
                fc2_man = st.text_input(
                    "Nombre completo (conductor 2)",
                    key="firma_c2_manual",
                    disabled=fc2_sel != PERSON_SEL_OTRO_MANUAL,
                )
                data.firma_conductor_2 = _valor_firma(fc2_sel, fc2_man)
            with ap2:
                fa1_sel = st.selectbox("Firma responsable de aprobación (1)", firma_opts, index=0)
                fa1_man = st.text_input(
                    "Nombre completo (aprueba 1)",
                    key="firma_a1_manual",
                    disabled=fa1_sel != PERSON_SEL_OTRO_MANUAL,
                )
                data.firma_aprueba_1 = _valor_firma(fa1_sel, fa1_man)
                fa2_sel = st.selectbox("Firma responsable de aprobación (2)", firma_opts, index=0)
                fa2_man = st.text_input(
                    "Nombre completo (aprueba 2)",
                    key="firma_a2_manual",
                    disabled=fa2_sel != PERSON_SEL_OTRO_MANUAL,
                )
                data.firma_aprueba_2 = _valor_firma(fa2_sel, fa2_man)
                data.fecha_firma = st.date_input("Fecha (firmas)", value=date.today())

            st.divider()
            gen_col1, gen_col2 = st.columns([1, 2])
            with gen_col1:
                _default_pdf = _default_plan_pdf_filename(
                    data.empresa, data.empresa_nombre, data.destino, date.today()
                )
                filename = st.text_input("Nombre del archivo PDF", value=_default_pdf)
            with gen_col2:
                st.caption(
                    "Al generar, podrás descargar el PDF en tu equipo. Opcionalmente, en servidores con "
                    "disco accesible, puedes guardar una copia en una carpeta local o de red."
                )

            submitted = st.form_submit_button("Generar PDF", type="primary")

        if submitted:
            paradas_ida_df = _paradas_widget_state_to_df("ida", st.session_state.paradas_ida_rows)
            paradas_vuelta_df = _paradas_widget_state_to_df(
                "vuelta", st.session_state.paradas_vuelta_rows
            )
            st.session_state.paradas_ida = paradas_ida_df
            st.session_state.paradas_vuelta = paradas_vuelta_df
            data.paradas_ida = _df_to_stops(paradas_ida_df)
            data.paradas_vuelta = _df_to_stops(paradas_vuelta_df)
            data.ruta_imagen_bytes = ruta_bytes
            data.ruta_vuelta_imagen_bytes = ruta_vuelta_bytes
            data.international_sos_imagen_bytes = sos_img_bytes
            data.international_sos_imagen_mime = "image/png" if sos_img_bytes else None

            pdf_bytes = build_plan_pdf(data)
            final_pdf_name = filename if filename.lower().endswith(".pdf") else f"{filename}.pdf"
            st.session_state[SESSION_PLAN_PDF_BYTES] = pdf_bytes
            st.session_state[SESSION_PLAN_PDF_FILENAME] = final_pdf_name
        _pdf_dl = st.session_state.get(SESSION_PLAN_PDF_BYTES)
        _pdf_name_dl = (st.session_state.get(SESSION_PLAN_PDF_FILENAME) or "").strip() or _default_plan_pdf_filename(
            "", "", "", date.today()
        )
        if _pdf_dl is not None:
            st.download_button(
                "Descargar PDF",
                data=_pdf_dl,
                file_name=_pdf_name_dl,
                mime="application/pdf",
                key="download_plan_pdf_bytes",
            )

        st.caption("Tip: el PDF solo se genera cuando presionas Generar PDF.")


if __name__ == "__main__":
    main()


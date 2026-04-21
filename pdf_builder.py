from __future__ import annotations

import io
from datetime import date
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.utils import simpleSplit
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# Paleta (PDF): títulos de sección azules; etiquetas vs valores en tablas llave-valor
SECTION_BLUE = colors.HexColor("#0d47a1")
LABEL_BG = colors.HexColor("#e8eef6")
VALUE_BG = colors.HexColor("#ffffff")
LABEL_TEXT = colors.HexColor("#37474f")


def _fmt_date(d: date | None) -> str:
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")


def _hazard_label_paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    """Etiqueta con ajuste de línea dentro del ancho de celda (evita solaparse con la columna X)."""
    t = (text or "").strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    if not t:
        return Paragraph("", style)
    return Paragraph(t.replace("\n", "<br/>"), style)


def _hazard_x_paragraph(checked: bool, style: ParagraphStyle) -> Paragraph:
    return Paragraph("X" if checked else "", style)


def _p(text: str, style: ParagraphStyle) -> Paragraph:
    t = (text or "").strip()
    if not t:
        return Paragraph("", style)
    return Paragraph(t.replace("\n", "<br/>"), style)


def _join_lines(items: list[str]) -> str:
    clean = [x.strip() for x in items if x and str(x).strip()]
    return "<br/>".join(clean)


def _boxed_text(label: str, text: str, label_style: ParagraphStyle, value_style: ParagraphStyle, height_cm: float) -> Table:
    """
    Caja con borde para texto multilinea (tipo formulario).
    """
    content = _p((text or "").strip() or "&nbsp;", value_style)
    t = Table(
        [[_p(label, label_style)], [content]],
        colWidths=[17.2 * cm],
        rowHeights=[None, height_cm * cm],
        hAlign="LEFT",
    )
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("BACKGROUND", (0, 0), (0, 0), LABEL_BG),
                ("BACKGROUND", (0, 1), (0, 1), VALUE_BG),
            ]
        )
    )
    return t


def _table(
    data: list[list[Any]],
    col_widths: list[float] | None = None,
    *,
    kv_shading: bool = False,
    header_row: bool = False,
) -> Table:
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    style_cmds: list[tuple[Any, ...]] = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
    ]
    if kv_shading and col_widths and len(col_widths) == 2:
        style_cmds.extend(
            [
                ("BACKGROUND", (0, 0), (0, -1), LABEL_BG),
                ("BACKGROUND", (1, 0), (1, -1), VALUE_BG),
            ]
        )
    if header_row and data:
        style_cmds.extend(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LABEL_BG),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 0), (-1, 0), SECTION_BLUE),
            ]
        )
    t.setStyle(TableStyle(style_cmds))
    return t


def _itinerary_cover(plan: Any, title_style: ParagraphStyle, small_label: ParagraphStyle, small_value: ParagraphStyle) -> list[Any]:
    """
    Portada tipo "itinerario" inspirada en plantilla visual del usuario.
    Conserva datos clave y deja el detalle completo para las secciones siguientes.
    """
    out: list[Any] = []
    usable_w = 17.2 * cm

    logo_cell: Any = _p("TU LOGO", small_label)
    logo_bytes = getattr(plan, "empresa_logo_bytes", None)
    if logo_bytes:
        try:
            logo = Image(io.BytesIO(logo_bytes))
            logo._restrictSize(2.6 * cm, 2.6 * cm)
            logo_cell = logo
        except Exception:
            pass

    head = Table(
        [[Paragraph("<b>PLAN DE GESTIÓN<br/>DE VIAJE</b>", title_style), logo_cell]],
        colWidths=[usable_w - 3.1 * cm, 3.1 * cm],
        hAlign="LEFT",
    )
    head.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    out.append(head)
    out.append(Spacer(1, 6))

    duracion = str(getattr(plan, "duracion_horas", "") or "").strip()
    if not duracion and getattr(plan, "fecha_salida", None) and getattr(plan, "fecha_llegada", None):
        duracion = f"{_fmt_date(getattr(plan, 'fecha_salida', None))} - {_fmt_date(getattr(plan, 'fecha_llegada', None))}"

    info_tbl = Table(
        [
            [
                _p("<b>DESTINO</b><br/>" + str(getattr(plan, "destino", "") or ""), small_value),
                _p("<b>DURACIÓN DE LA ESTANCIA</b><br/>" + duracion, small_value),
            ],
            [
                _p("<b>SALIDA DEL VUELO / VIAJE</b><br/>" + str(getattr(plan, "hora_salida", "") or ""), small_value),
                _p("<b>LLEGADA DEL VUELO / VIAJE</b><br/>" + str(getattr(plan, "hora_llegada", "") or ""), small_value),
            ],
        ],
        colWidths=[usable_w / 2, usable_w / 2],
        hAlign="LEFT",
    )
    info_tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    out.append(info_tbl)

    paradas = getattr(plan, "paradas_ida", []) or []
    max_rows = max(4, len(paradas))
    day_rows: list[list[Any]] = []
    for i in range(max_rows):
        stop = paradas[i] if i < len(paradas) else None
        que_hacer = ""
        presupuesto = ""
        if stop:
            lugar = str(getattr(stop, "lugar", "") or "").strip()
            motivo = str(getattr(stop, "motivo", "") or "").strip()
            tiempo = str(getattr(stop, "tiempo_min", "") or "").strip()
            que_hacer = " - ".join([x for x in [lugar, motivo] if x])
            presupuesto = f"{tiempo} min" if tiempo else ""
        day_rows.append(
            [
                _p(f"<b>DÍA {i + 1:02d}</b>", small_label),
                _p("<b>QUÉ HACER</b><br/>" + que_hacer, small_value),
                _p("<b>PRESUPUESTO</b><br/>" + presupuesto, small_value),
            ]
        )

    days_tbl = Table(
        day_rows,
        colWidths=[1.4 * cm, 10.0 * cm, usable_w - 11.4 * cm],
        rowHeights=[2.2 * cm] * len(day_rows),
        hAlign="LEFT",
    )
    days_tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    out.append(days_tbl)

    footer_txt = (
        f"<font color='white'><b>Fecha de inicio:</b> {_fmt_date(getattr(plan, 'fecha_salida', None)) or '-'}"
        f" &nbsp;&nbsp; <b>Fecha de finalización:</b> {_fmt_date(getattr(plan, 'fecha_llegada', None)) or '-'}</font>"
    )
    footer_tbl = Table([[Paragraph(footer_txt, small_value)]], colWidths=[usable_w], hAlign="LEFT")
    footer_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), colors.black),
                ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    out.append(footer_tbl)
    out.append(Spacer(1, 8))
    return out


def build_plan_pdf(plan: Any) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=LETTER,
        leftMargin=1.3 * cm,
        rightMargin=1.3 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="PLAN DE GESTIÓN DE VIAJE ALS ECUADOR",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleSmall",
        parent=styles["Title"],
        fontSize=24,
        leading=26,
        spaceAfter=8,
        fontName="Helvetica-Bold",
    )
    h_style = ParagraphStyle(
        "H",
        parent=styles["Heading2"],
        fontSize=11,
        leading=13,
        spaceBefore=10,
        spaceAfter=6,
        textColor=SECTION_BLUE,
        fontName="Helvetica-Bold",
    )
    small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=9, leading=11)
    small_label = ParagraphStyle(
        "SmallLabel",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        textColor=LABEL_TEXT,
        fontName="Helvetica-Bold",
    )
    small_value = ParagraphStyle(
        "SmallValue",
        parent=styles["Normal"],
        fontSize=8.8,
        leading=10.4,
        textColor=colors.black,
        fontName="Helvetica",
    )
    hazard_label_style = ParagraphStyle(
        "HazardLabel",
        parent=small_label,
        alignment=TA_LEFT,
        fontSize=9,
        leading=11,
        textColor=LABEL_TEXT,
        fontName="Helvetica-Bold",
    )
    hazard_x_style = ParagraphStyle(
        "HazardX",
        parent=styles["Normal"],
        fontSize=10,
        leading=12,
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
        textColor=colors.black,
    )

    story: list[Any] = []
    story.extend(_itinerary_cover(plan, title_style, small_label, small_value))
    story.append(Spacer(1, 4))

    # Encabezado (estilo RU-40: logo | título | código/rev/fecha)
    empresa = (getattr(plan, "empresa_nombre", "") or "").strip() or "ALS ECUADOR"
    doc_code = (getattr(plan, "doc_code", "") or "").strip()
    doc_rev = (getattr(plan, "doc_rev", "") or "").strip()
    doc_date = (getattr(plan, "doc_date", "") or "").strip()
    logo_bytes = getattr(plan, "empresa_logo_bytes", None)
    center_style = ParagraphStyle(
        "HeaderCenter",
        parent=styles["Normal"],
        fontSize=12,
        leading=14,
        alignment=1,  # center
        textColor=colors.grey,
    )
    right_style = ParagraphStyle(
        "HeaderRight",
        parent=styles["Normal"],
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.grey,
    )

    logo_cell: Any = ""
    if logo_bytes:
        try:
            logo = Image(io.BytesIO(logo_bytes))
            logo._restrictSize(2.6 * cm, 2.6 * cm)
            logo_cell = logo
        except Exception:
            logo_cell = ""

    center_cell = Paragraph(
        "<b>PLAN GESTIÓN DE VIAJE</b><br/><b>SISTEMA INTEGRADO DE GESTIÓN</b>",
        center_style,
    )
    right_cell = Paragraph(
        "<b>{}</b><br/>{}<br/>{}".format(
            doc_code or "&nbsp;",
            doc_rev or "&nbsp;",
            doc_date or "&nbsp;",
        ),
        right_style,
    )

    header_tbl = Table(
        [[logo_cell, center_cell, right_cell]],
        colWidths=[3.2 * cm, 11.0 * cm, 3.0 * cm],
        hAlign="LEFT",
    )
    header_tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(header_tbl)
    story.append(Spacer(1, 6))

    story.append(
        Paragraph(
            "Instrucción: Este formulario debe ser completado para viajes de más de 300 km (ida y vuelta), "
            "que duren más de 4 horas, o cualquier viaje identificado con riesgos particulares.",
            small,
        )
    )
    story.append(Spacer(1, 8))

    # 1. Datos generales
    story.append(Paragraph("1. DATOS GENERALES", h_style))
    cond_list = getattr(plan, "conductores", None) or []
    ced_list = getattr(plan, "cedulas_conductores", None) or []
    cel_list = getattr(plan, "celulares_conductores", None) or []
    if not cond_list:
        cond_list = [x for x in [getattr(plan, "conductor_1", ""), getattr(plan, "conductor_2", "")] if x]
    if not ced_list:
        ced_list = [x for x in [getattr(plan, "cedula_1", ""), getattr(plan, "cedula_2", "")] if x]
    if not cel_list:
        cel_list = [x for x in [getattr(plan, "conductor_cel_1", ""), getattr(plan, "conductor_cel_2", "")] if x]
    conductores = _join_lines([str(x) for x in cond_list])
    cedulas = _join_lines([str(x) for x in ced_list])
    tels_cond = _join_lines([str(x) for x in cel_list])
    emerg = _join_lines([getattr(plan, "emergencia_1", ""), getattr(plan, "emergencia_2", "")])
    tel_emerg = _join_lines([getattr(plan, "tel_emergencia_1", ""), getattr(plan, "tel_emergencia_2", "")])
    ced_emerg = _join_lines(
        [getattr(plan, "cedula_emergencia_1", ""), getattr(plan, "cedula_emergencia_2", "")]
    )

    t1 = _table(
        [
            [_p("Nombre del Conductor/Responsable", small_label), _p(conductores, small_value)],
            [_p("Fecha de Elaboración", small_label), _p(_fmt_date(getattr(plan, "fecha_elab", None)), small_value)],
            [_p("Cédula de Identidad", small_label), _p(cedulas, small_value)],
            [_p("Celular", small_label), _p(tels_cond, small_value)],
            [_p("Cargo / Posición", small_label), _p(str(getattr(plan, "cargo", "") or ""), small_value)],
            [_p("Punto de Origen (Ciudad/Provincia)", small_label), _p(str(getattr(plan, "origen", "") or ""), small_value)],
            [_p("Placa del Vehículo", small_label), _p(str(getattr(plan, "placa", "") or ""), small_value)],
            [_p("Tipo de Vehículo", small_label), _p(str(getattr(plan, "tipo_vehiculo", "") or ""), small_value)],
            [_p("Modelo / Año", small_label), _p(str(getattr(plan, "modelo_anio", "") or ""), small_value)],
            [_p("Nombre del Contacto de Emergencia", small_label), _p(emerg, small_value)],
            [_p("Teléfono de Emergencia", small_label), _p(tel_emerg, small_value)],
            [_p("Cédula de identidad (contacto emergencia)", small_label), _p(ced_emerg, small_value)],
        ],
        col_widths=[6.2 * cm, 11.0 * cm],
        kv_shading=True,
    )
    story.append(t1)

    # 2. Planificación
    story.append(Paragraph("2. PLANIFICACIÓN DE VIAJE", h_style))
    t2 = _table(
        [
            [_p("Destino (Ciudad/Provincia)", small_label), _p(str(getattr(plan, "destino", "") or ""), small_value)],
            [_p("Empresa", small_label), _p(str(getattr(plan, "empresa", "") or ""), small_value)],
            [_p("Orden de Trabajo", small_label), _p(str(getattr(plan, "orden_trabajo", "") or ""), small_value)],
            [_p("Fecha de Salida", small_label), _p(_fmt_date(getattr(plan, "fecha_salida", None)), small_value)],
            [_p("Hora de Salida", small_label), _p(str(getattr(plan, "hora_salida", "") or ""), small_value)],
            [_p("Fecha Estimada de Llegada", small_label), _p(_fmt_date(getattr(plan, "fecha_llegada", None)), small_value)],
            [_p("Hora Estimada de Llegada", small_label), _p(str(getattr(plan, "hora_llegada", "") or ""), small_value)],
            [_p("Distancia Total (km)", small_label), _p(str(getattr(plan, "distancia_km", "") or ""), small_value)],
            [_p("Duración Estimada (horas)", small_label), _p(str(getattr(plan, "duracion_horas", "") or ""), small_value)],
        ],
        col_widths=[6.2 * cm, 11.0 * cm],
        kv_shading=True,
    )
    story.append(t2)
    story.append(Spacer(1, 6))

    story.append(
        _boxed_text(
            "Propósito del viaje",
            str(getattr(plan, "proposito", "") or ""),
            small_label,
            small_value,
            height_cm=2.0,
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        _boxed_text(
            "Condiciones del camino",
            str(getattr(plan, "condiciones_camino", "") or ""),
            small_label,
            small_value,
            height_cm=0.9,
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        _boxed_text(
            "Observaciones adicionales",
            str(getattr(plan, "observaciones", "") or ""),
            small_label,
            small_value,
            height_cm=1.6,
        )
    )

    sos_txt = (getattr(plan, "international_sos_text", "") or "").strip()
    sos_img = getattr(plan, "international_sos_imagen_bytes", None)
    if sos_txt or sos_img:
        story.append(Spacer(1, 8))
        story.append(
            Paragraph(
                "<font color='#0d47a1'><b>APP International SOS</b></font>",
                small_label,
            )
        )
        story.append(Spacer(1, 4))
        if sos_txt:
            story.append(
                _boxed_text(
                    "Texto o notas (APP International SOS)",
                    sos_txt,
                    small_label,
                    small_value,
                    height_cm=1.4,
                )
            )
            if sos_img:
                story.append(Spacer(1, 6))
        if sos_img:
            story.append(Paragraph("<b>Captura / imagen desde la app:</b>", small))
            try:
                im_sos = Image(io.BytesIO(sos_img))
                im_sos._restrictSize(17.5 * cm, 9.5 * cm)
                story.append(im_sos)
            except Exception:
                story.append(
                    Paragraph("No se pudo incrustar la imagen de International SOS.", small_value)
                )

    # 3. Paradas ida
    story.append(Paragraph("3. PARADAS PLANIFICADAS (IDA)", h_style))
    paradas_ida = getattr(plan, "paradas_ida", []) or []
    stops_rows = [["N°", "Lugar / Ciudad", "Motivo", "Tiempo estimado (min)"]]
    for s in paradas_ida:
        stops_rows.append([getattr(s, "n", ""), getattr(s, "lugar", ""), getattr(s, "motivo", ""), getattr(s, "tiempo_min", "")])
    story.append(
        _table(
            stops_rows,
            col_widths=[1.0 * cm, 5.6 * cm, 7.6 * cm, 4.0 * cm],
            header_row=True,
        )
    )

    # 4. Peligros (Paragraph en etiquetas para que el texto haga wrap y no invada la celda de la X)
    story.append(Paragraph("4. PELIGROS CONOCIDOS / MARCA CON UNA X", h_style))
    hazard_data = [
        [
            _hazard_label_paragraph("Lluvia", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_lluvia", False)), hazard_x_style),
            _hazard_label_paragraph("Niebla", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_niebla", False)), hazard_x_style),
            _hazard_label_paragraph("Nieve / hielo", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_nieve_hielo", False)), hazard_x_style),
        ],
        [
            _hazard_label_paragraph("Conducción nocturna", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_nocturna", False)), hazard_x_style),
            _hazard_label_paragraph("Carreteras en mal estado", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_carretera_mala", False)), hazard_x_style),
            _hazard_label_paragraph("Zona de alta delincuencia", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_delincuencia", False)), hazard_x_style),
        ],
        [
            _hazard_label_paragraph("Accidentes de tránsito", hazard_label_style),
            _hazard_x_paragraph(bool(getattr(plan, "peligro_accidentes_transito", False)), hazard_x_style),
            _hazard_label_paragraph("", hazard_label_style),
            _hazard_x_paragraph(False, hazard_x_style),
            _hazard_label_paragraph("", hazard_label_style),
            _hazard_x_paragraph(False, hazard_x_style),
        ],
    ]
    hazards = Table(
        hazard_data,
        colWidths=[5.25 * cm, 1.05 * cm, 5.25 * cm, 1.05 * cm, 5.25 * cm, 1.05 * cm],
        hAlign="LEFT",
    )
    hazards.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [LABEL_BG, colors.HexColor("#f5f7fb")]),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("ALIGN", (3, 0), (3, -1), "CENTER"),
                ("ALIGN", (5, 0), (5, -1), "CENTER"),
                ("LEFTPADDING", (1, 0), (1, -1), 2),
                ("RIGHTPADDING", (1, 0), (1, -1), 2),
                ("LEFTPADDING", (3, 0), (3, -1), 2),
                ("RIGHTPADDING", (3, 0), (3, -1), 2),
                ("LEFTPADDING", (5, 0), (5, -1), 2),
                ("RIGHTPADDING", (5, 0), (5, -1), 2),
            ]
        )
    )
    story.append(hazards)
    story.append(Spacer(1, 6))
    story.append(Paragraph("<font color='#0d47a1'><b>Otros peligros / detalles adicionales:</b></font>", small))
    story.append(Paragraph((getattr(plan, "otros_peligros", "") or "").replace("\n", "<br/>"), small_value))

    # Ruta (imagen)
    img_bytes = getattr(plan, "ruta_imagen_bytes", None)
    if img_bytes:
        story.append(Spacer(1, 8))
        story.append(Paragraph("<font color='#0d47a1'><b>Ruta para tomar (imagen):</b></font>", small))
        try:
            im = Image(io.BytesIO(img_bytes))
            im._restrictSize(17.5 * cm, 9.5 * cm)
            story.append(im)
        except Exception:
            story.append(Paragraph("No se pudo incrustar la imagen de ruta (formato no compatible).", small))

    # Pasajeros ida (1 columna)
    story.append(Spacer(1, 8))
    story.append(Paragraph("Pasajeros (IDA)", h_style))
    pasajeros_ida = getattr(plan, "pasajeros_ida", []) or []
    story.append(
        _table(
            [
                [
                    _p("Cantidad de Pasajeros", small_label),
                    _p(str(len(pasajeros_ida)) if pasajeros_ida else "", small_value),
                ],
                [_p("Listado", small_label), _p(_join_lines([str(x) for x in pasajeros_ida]), small_value)],
            ],
            col_widths=[6.2 * cm, 11.0 * cm],
            kv_shading=True,
        )
    )

    # 5. Vuelta (1 columna)
    story.append(Paragraph("5. VIAJE DE VUELTA", h_style))
    pasajeros_vuelta = getattr(plan, "pasajeros_vuelta", []) or []
    story.append(
        _table(
            [
                [_p("Hora de salida", small_label), _p(str(getattr(plan, "vuelta_hora_salida", "") or ""), small_value)],
                [
                    _p("Hora estimada de llegada", small_label),
                    _p(str(getattr(plan, "vuelta_hora_llegada", "") or ""), small_value),
                ],
                [_p("Fecha de salida", small_label), _p(_fmt_date(getattr(plan, "vuelta_fecha_salida", None)), small_value)],
                [
                    _p("Fecha estimada de llegada", small_label),
                    _p(_fmt_date(getattr(plan, "vuelta_fecha_llegada", None)), small_value),
                ],
                [
                    _p("Cantidad de Pasajeros", small_label),
                    _p(str(len(pasajeros_vuelta)) if pasajeros_vuelta else "", small_value),
                ],
                [_p("Listado", small_label), _p(_join_lines([str(x) for x in pasajeros_vuelta]), small_value)],
            ],
            col_widths=[6.2 * cm, 11.0 * cm],
            kv_shading=True,
        )
    )

    # Ruta vuelta (imagen)
    img_vuelta_bytes = getattr(plan, "ruta_vuelta_imagen_bytes", None)
    if img_vuelta_bytes:
        story.append(Spacer(1, 8))
        story.append(Paragraph("<font color='#0d47a1'><b>Ruta para tomar (imagen) (VUELTA):</b></font>", small))
        try:
            imv = Image(io.BytesIO(img_vuelta_bytes))
            imv._restrictSize(17.5 * cm, 9.5 * cm)
            story.append(imv)
        except Exception:
            story.append(Paragraph("No se pudo incrustar la imagen de ruta (VUELTA).", small))

    story.append(Paragraph("PARADAS PLANIFICADAS (VUELTA)", h_style))
    paradas_vuelta = getattr(plan, "paradas_vuelta", []) or []
    v_rows = [["N°", "Lugar / Ciudad", "Motivo", "Tiempo estimado (min)"]]
    for s in paradas_vuelta:
        v_rows.append([getattr(s, "n", ""), getattr(s, "lugar", ""), getattr(s, "motivo", ""), getattr(s, "tiempo_min", "")])
    if len(v_rows) == 1:
        v_rows.append(["", "", "", ""])
    story.append(
        _table(
            v_rows,
            col_widths=[1.0 * cm, 5.6 * cm, 7.6 * cm, 4.0 * cm],
            header_row=True,
        )
    )

    # 6. Aprobación (1 columna)
    story.append(Paragraph("6. APROBACIÓN", h_style))
    firma_conductores = _join_lines(
        [x for x in [getattr(plan, "firma_conductor_1", ""), getattr(plan, "firma_conductor_2", "")] if x and x != "—"]
    )
    firma_aprueba = _join_lines(
        [x for x in [getattr(plan, "firma_aprueba_1", ""), getattr(plan, "firma_aprueba_2", "")] if x and x != "—"]
    )
    story.append(
        _table(
            [
                [
                    _p("Firma responsable elaboración plan", small_label),
                    _p(str(getattr(plan, "firma_elabora", "") or ""), small_value),
                ],
                [_p("Firma conductor responsable", small_label), _p(firma_conductores, small_value)],
                [_p("Firma responsable de aprobación plan", small_label), _p(firma_aprueba, small_value)],
                [_p("Fecha", small_label), _p(_fmt_date(getattr(plan, "fecha_firma", None)), small_value)],
            ],
            col_widths=[6.2 * cm, 11.0 * cm],
            kv_shading=True,
        )
    )

    # 7. Direcciones y contactos (estático)
    story.append(Paragraph("7. DIRECCIONES Y CONTACTOS PARA CONSULTA", h_style))
    story.append(
        Paragraph(
            "Estado Vial Oficial (Tiempo Real): Servicio Integrado de Seguridad ECU 911 — "
            "https://www.ecu911.gob.ec/consulta-de-vias/ — @ECU911_ (Google Play / App Store)<br/>"
            "Infraestructura y Obras Públicas: MTOP — https://www.obraspublicas.gob.ec/ — @ObrasPublicasEc<br/>"
            "Gestión de Riesgos y Amenazas: SGR — https://www.gestionderiesgos.gob.ec/ — @Riesgos_Ec<br/>"
            "Navegación Comunitaria: Waze — https://www.waze.com/es-419/live-map/ — @waze<br/>"
            "Noticias: Teleamazonas — https://www.teleamazonas.com/ — @teleamazonasec<br/>"
            "Noticias: Ecuavisa — https://www.ecuavisa.com/ — @ecuavisa<br/>"
            "Noticias: El Universo — https://www.eluniverso.com/ — @eluniversocom<br/>"
            "Noticias: El Comercio — https://www.elcomercio.com/ — @elcomerciocom<br/>"
            "Noticias: Primicias — https://www.primicias.ec/ — @Primicias<br/>"
            "Seguridad y Asistencia: Policía Nacional — https://www.gob.ec/pn — @PoliciaEcuador",
            small,
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        Paragraph(
            "Nota: Los formularios completos se conservarán durante un período de 1 año. "
            "Es prioritario elegir la ruta más segura sobre la más corta.",
            small,
        )
    )

    doc.build(story)
    return buf.getvalue()


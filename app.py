
import os
import re
import shutil
import tempfile
import subprocess
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


st.set_page_config(page_title="Balance Publicación a Word", layout="wide")

st.title("Generador de Word - Balance de Publicación")
st.caption("Versión por áreas de impresión: cada área de impresión del Excel se convierte en una página del Word.")


# -------------------------
# Utilidades generales
# -------------------------

def safe_filename(name: str) -> str:
    name = (name or "Balance Publicacion.docx").strip()
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def save_upload(uploaded_file) -> str:
    temp_dir = tempfile.mkdtemp(prefix="balance_publicacion_")
    path = os.path.join(temp_dir, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def normalize_print_area(print_area) -> str:
    if print_area is None:
        return ""
    return str(print_area)


def split_print_areas(print_area: str, sheet_name: str) -> list[str]:
    """
    Recibe algo tipo:
    'Bal Public 2025 (REEX)'!$A$18:$E$85,'Bal Public 2025 (REEX)'!$A$89:$E$145
    y devuelve:
    ['$A$18:$E$85', '$A$89:$E$145']
    """
    if not print_area:
        return []

    raw_parts = []
    current = []
    in_quote = False

    for ch in print_area:
        if ch == "'":
            in_quote = not in_quote
            current.append(ch)
        elif ch == "," and not in_quote:
            raw_parts.append("".join(current).strip())
            current = []
        else:
            current.append(ch)

    if current:
        raw_parts.append("".join(current).strip())

    ranges = []
    for part in raw_parts:
        if "!" in part:
            part = part.split("!", 1)[1]
        part = part.replace("'", "").strip()
        if part:
            ranges.append(part)

    return ranges


def range_size_score(ws, area: str) -> tuple[float, float]:
    """
    Estima ancho y alto visual usando anchos de columnas y altos de filas.
    Sirve para decidir vertical/horizontal.
    """
    min_col, min_row, max_col, max_row = range_boundaries(area)

    width = 0.0
    for c in range(min_col, max_col + 1):
        letter = ws.cell(1, c).column_letter
        dim = ws.column_dimensions.get(letter)
        width += float(dim.width or 8.43)

    height = 0.0
    for r in range(min_row, max_row + 1):
        dim = ws.row_dimensions.get(r)
        height += float(dim.height or 15)

    return width, height


def auto_orientation(ws, area: str) -> str:
    width, height = range_size_score(ws, area)

    # Los anchos de columna y altos de fila no están en la misma unidad exacta,
    # pero esta relación funciona bien para separar páginas muy anchas.
    return "H" if width * 7 > height * 0.85 else "V"


def make_area_lines(ws, areas: list[str]) -> str:
    lines = []
    for a in areas:
        ori = auto_orientation(ws, a)
        lines.append(f"{a} | {ori}")
    return "\n".join(lines)


def parse_area_lines(text: str) -> list[tuple[str, str]]:
    result = []

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        if "|" in line:
            area, ori = [x.strip().upper() for x in line.split("|", 1)]
        else:
            area, ori = line.upper(), "AUTO"

        # validar rango
        range_boundaries(area)

        if ori not in ("V", "H", "AUTO"):
            raise ValueError(f"Orientación inválida en línea: {line}. Usá V, H o AUTO.")

        result.append((area, ori))

    return result


# -------------------------
# Preparación de Excel temporal por página
# -------------------------

def hide_all_except(wb, sheet_name: str):
    for ws in wb.worksheets:
        ws.sheet_state = "visible" if ws.title == sheet_name else "hidden"
    wb.active = wb.sheetnames.index(sheet_name)


def prepare_single_area_workbook(
    original_xlsx: str,
    output_xlsx: str,
    sheet_name: str,
    area: str,
    orientation: str
):
    """
    Crea una copia temporal del Excel con:
    - solo visible la hoja elegida,
    - área de impresión = una sola área,
    - ajuste a 1 página,
    - orientación V/H.
    """
    wb = load_workbook(original_xlsx)
    ws = wb[sheet_name]

    hide_all_except(wb, sheet_name)

    ws.print_area = area
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if orientation == "H" else "portrait"

    # Márgenes chicos para aprovechar la hoja.
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0
    ws.page_margins.footer = 0

    wb.save(output_xlsx)


def libreoffice_to_pdf(soffice: str, xlsx_path: str, out_dir: str) -> str:
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        xlsx_path,
    ]

    res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    expected = os.path.join(out_dir, Path(xlsx_path).with_suffix(".pdf").name)

    if res.returncode != 0 or not os.path.exists(expected):
        raise RuntimeError(
            "LibreOffice no pudo convertir el Excel a PDF.\n\n"
            f"STDOUT:\n{res.stdout}\n\nSTDERR:\n{res.stderr}"
        )

    return expected


def pdf_first_page_to_png(pdf_path: str, png_path: str, zoom: float = 3.0):
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    pix.save(png_path)
    pdf.close()


# -------------------------
# Word
# -------------------------

def add_image_page(doc: Document, png_path: str, orientation: str, first: bool):
    if not first:
        doc.add_section()

    section = doc.sections[-1]

    if orientation == "H":
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Cm(29.7)
        section.page_height = Cm(21.0)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)

    # Como la imagen ya es una página PDF renderizada, usamos márgenes 0.
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    usable_width = section.page_width
    usable_height = section.page_height

    p = doc.add_paragraph()
    p.paragraph_format.space_before = 0
    p.paragraph_format.space_after = 0
    run = p.add_run()
    pic = run.add_picture(png_path, width=usable_width)

    # Seguridad por si la imagen queda apenas más alta.
    if pic.height > usable_height:
        pic.height = usable_height


def build_word(
    original_xlsx: str,
    soffice: str,
    caratula_sheet: str,
    caratula_area: str,
    balance_sheet: str,
    balance_areas: list[tuple[str, str]],
    output_name: str,
):
    temp_dir = tempfile.mkdtemp(prefix="balance_render_")
    doc = Document()

    pages = []

    # Página 1: carátula
    pages.append((caratula_sheet, caratula_area, "V", "Carátula"))

    # Páginas siguientes
    wb_read = load_workbook(original_xlsx, read_only=True, data_only=False)
    ws_bal_read = wb_read[balance_sheet]

    for idx, (area, ori) in enumerate(balance_areas, start=1):
        if ori == "AUTO":
            ori = auto_orientation(ws_bal_read, area)
        pages.append((balance_sheet, area, ori, f"Balance {idx}"))

    for idx, (sheet, area, ori, label) in enumerate(pages, start=1):
        st.write(f"Procesando página {idx}: {sheet} - {area} - {'Horizontal' if ori == 'H' else 'Vertical'}")

        xlsx_tmp = os.path.join(temp_dir, f"page_{idx:03d}.xlsx")
        pdf_tmp_dir = os.path.join(temp_dir, f"pdf_{idx:03d}")
        os.makedirs(pdf_tmp_dir, exist_ok=True)
        png_tmp = os.path.join(temp_dir, f"page_{idx:03d}.png")

        prepare_single_area_workbook(original_xlsx, xlsx_tmp, sheet, area, ori)
        pdf_path = libreoffice_to_pdf(soffice, xlsx_tmp, pdf_tmp_dir)
        pdf_first_page_to_png(pdf_path, png_tmp)

        add_image_page(doc, png_tmp, ori, first=(idx == 1))

    output_path = os.path.join(temp_dir, safe_filename(output_name))
    doc.save(output_path)
    return output_path, len(pages)


# -------------------------
# UI
# -------------------------

soffice_path = find_soffice()

if soffice_path:
    st.success(f"LibreOffice detectado: {soffice_path}")
else:
    st.error(
        "No encontré LibreOffice. Instalá LibreOffice y volvé a abrir la app. "
        "Esta versión no abre Excel ni Word: usa LibreOffice en modo oculto."
    )

uploaded = st.file_uploader("Cargar Excel con áreas de impresión definidas", type=["xlsx", "xlsm"])

if uploaded:
    xlsx_path = save_upload(uploaded)
    wb = load_workbook(xlsx_path, read_only=False, data_only=False)
    sheetnames = wb.sheetnames

    st.success("Excel cargado correctamente.")

    # Sugerencias de hojas
    car_idx = 0
    bal_idx = 0

    for i, name in enumerate(sheetnames):
        low = name.lower()
        if "car" in low:
            car_idx = i
        if "bal public" in low or "reex" in low:
            bal_idx = i

    col1, col2 = st.columns(2)

    with col1:
        car_sheet = st.selectbox("Hoja de carátula", sheetnames, index=car_idx)

    with col2:
        bal_sheet = st.selectbox("Hoja del balance", sheetnames, index=bal_idx)

    ws_car = wb[car_sheet]
    ws_bal = wb[bal_sheet]

    car_areas = split_print_areas(normalize_print_area(ws_car.print_area), car_sheet)
    bal_areas = split_print_areas(normalize_print_area(ws_bal.print_area), bal_sheet)

    if not car_areas:
        st.warning("La hoja de carátula no tiene área de impresión definida.")
        car_area_default = ""
    else:
        car_area_default = car_areas[0]

    if not bal_areas:
        st.warning("La hoja del balance no tiene áreas de impresión definidas.")
        bal_lines_default = ""
    else:
        bal_lines_default = make_area_lines(ws_bal, bal_areas)

    st.subheader("Áreas detectadas")

    car_area = st.text_input("Área de impresión de carátula", value=car_area_default)

    st.write("Áreas del balance. Formato: `RANGO | V` o `RANGO | H`. También podés usar `AUTO`.")
    balance_area_text = st.text_area(
        "Áreas de impresión del balance, una por línea",
        value=bal_lines_default,
        height=280,
    )

    try:
        parsed_areas = parse_area_lines(balance_area_text)
        st.info(f"Se generarán {1 + len(parsed_areas)} páginas: 1 carátula + {len(parsed_areas)} páginas del balance.")
    except Exception as e:
        parsed_areas = []
        st.error(f"Hay un error en las áreas: {e}")

    output_name = st.text_input("Nombre del Word de salida", value="Balance Publicacion.docx")

    if st.button("Generar Word", disabled=(not soffice_path or not car_area or not parsed_areas)):
        try:
            with st.spinner("Generando Word. Esto puede tardar unos minutos..."):
                range_boundaries(car_area)  # validar
                output_path, total_pages = build_word(
                    original_xlsx=xlsx_path,
                    soffice=soffice_path,
                    caratula_sheet=car_sheet,
                    caratula_area=car_area,
                    balance_sheet=bal_sheet,
                    balance_areas=parsed_areas,
                    output_name=output_name,
                )

            st.success(f"Word generado correctamente. Total de páginas: {total_pages}")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="Descargar Word",
                    data=f,
                    file_name=safe_filename(output_name),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
            st.warning("Mandame captura del error y reviso el ajuste.")

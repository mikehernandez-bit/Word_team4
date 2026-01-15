import os
import json
import platform
import subprocess
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# -------------------------
# UTILIDADES JSON / PATHS
# -------------------------

def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def resolve_path(base_dir: str, maybe_rel: str) -> str:
    return maybe_rel if os.path.isabs(maybe_rel) else os.path.join(base_dir, maybe_rel)

def resolve_config_path(base_dir: str, config_path: str) -> str:
    if os.path.isabs(config_path):
        return config_path
    candidate = os.path.join(base_dir, config_path)
    if os.path.exists(candidate):
        return candidate
    return os.path.abspath(config_path)

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


# -------------------------
# WORD: SETUP / HELPERS
# -------------------------

def open_document(path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as exc:
        print("[WARN] No se pudo abrir el documento:", exc)

def set_page_setup(doc: Document, cfg: dict):
    page_setup = cfg.get("page_setup", {})
    margins = page_setup.get("margins_cm", {})

    sec = doc.sections[0]

    # A4
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)

    # Márgenes
    sec.left_margin = Cm(float(margins.get("left", 3.5)))
    sec.right_margin = Cm(float(margins.get("right", 2.5)))
    sec.top_margin = Cm(float(margins.get("top", 3.0)))
    sec.bottom_margin = Cm(float(margins.get("bottom", 3.0)))

    # Fuente base
    font_cfg = page_setup.get("font", {})
    font_name = font_cfg.get("name", "Arial")
    font_size = float(font_cfg.get("size_pt", 12))

    normal = doc.styles["Normal"]
    normal.font.name = font_name
    normal.font.size = Pt(font_size)


def add_center_line(doc: Document, text: str, size=12, bold=False, uppercase=False, spacing_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text.upper() if uppercase else text)
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(spacing_after)
    return p


def add_heading(doc: Document, text: str, level: int = 1, spacing_after: int = 0):
    p = doc.add_paragraph(text)
    p.style = f"Heading {level}"
    if p.runs:
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(12)
    p.paragraph_format.space_after = Pt(spacing_after)
    return p


def add_center_logo(doc: Document, logo_path: str, width_cm: float = 3.5, spacing_after_pt: int = 6):
    if not logo_path or not os.path.exists(logo_path):
        print(f"[WARN] No se encontró el logo en: {logo_path}")
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(logo_path, width=Cm(width_cm))
    p.paragraph_format.space_after = Pt(spacing_after_pt)


# -------------------------
# WORD: CAMPOS (ÍNDICES)
# -------------------------

def add_field(paragraph, instr: str):
    run = paragraph.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run._r.append(fld)


def add_toc_page(doc: Document, toc_cfg: dict):
    p = doc.add_paragraph("ÍNDICE")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.runs[0]
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(12)

    min_lv = int(toc_cfg.get("min_level", 1))
    max_lv = int(toc_cfg.get("max_level", 3))

    toc_p = doc.add_paragraph()
    add_field(toc_p, f'TOC \\\\o "{min_lv}-{max_lv}" \\\\h \\\\z \\\\u')

    doc.add_page_break()


def add_list_of_tables(doc: Document):
    p = doc.add_paragraph("ÍNDICE DE TABLAS")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.runs[0]
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(12)

    lot_p = doc.add_paragraph()
    add_field(lot_p, 'TOC \\\\h \\\\z \\\\c "Tabla"')

    doc.add_page_break()


def add_list_of_figures(doc: Document):
    p = doc.add_paragraph("ÍNDICE DE FIGURAS")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.runs[0]
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(12)

    lof_p = doc.add_paragraph()
    add_field(lof_p, 'TOC \\\\h \\\\z \\\\c "Figura"')

    doc.add_page_break()


# -------------------------
# CARÁTULA DESDE JSON
# -------------------------

def add_cover_from_cfg(doc: Document, cfg: dict, base_dir: str):
    cover = cfg["cover"]

    logo_path = resolve_path(base_dir, cfg.get("logo_path", ""))
    logo_width = float(cover.get("logo_width_cm", 3.5))

    # Logo
    add_center_logo(doc, logo_path, width_cm=logo_width, spacing_after_pt=6)

    text_size = int(cover.get("text_size_pt", 12))
    title_size = int(cover.get("title_size_pt", 14))

    add_center_line(doc, cover["universidad_linea"], size=text_size, bold=True, uppercase=True, spacing_after=6)
    add_center_line(doc, cover["unidad"], size=text_size, bold=False, uppercase=True, spacing_after=18)

    add_center_line(doc, f"“{cover['titulo']}”", size=title_size, bold=True, uppercase=True, spacing_after=12)
    add_center_line(
        doc,
        f"TESIS PARA OPTAR EL GRADO ACADÉMICO DE {cover['grado_maestria']}",
        size=text_size,
        bold=True,
        uppercase=True,
        spacing_after=18
    )

    add_center_line(doc, cover["autor"], size=text_size, uppercase=True, spacing_after=6)
    add_center_line(doc, cover["asesor"], size=text_size, uppercase=True, spacing_after=6)
    add_center_line(doc, f"LINEA DE INVESTIGACIÓN: {cover['linea']}", size=text_size, uppercase=True, spacing_after=18)

    add_center_line(doc, f"{cover.get('ciudad','')}, {cover['anio']}", size=text_size, uppercase=True, spacing_after=6)
    add_center_line(doc, cover.get("pais", "PERÚ"), size=text_size, uppercase=True, spacing_after=0)

    doc.add_page_break()


# -------------------------
# ESTRUCTURA DESDE JSON
# -------------------------

def add_structure_from_cfg(doc: Document, cfg: dict):
    rules = cfg.get("structure_rules", {})
    add_placeholder = bool(rules.get("add_placeholder_after_heading", True))
    break_after_level1 = bool(rules.get("page_break_after_level_1", True))

    structure = cfg.get("structure", [])
    total = len(structure)

    for i, item in enumerate(structure):
        lvl = int(item["level"])
        title = item["title"]

        add_heading(doc, title, level=lvl)

        if add_placeholder:
            doc.add_paragraph("{{COMPLETAR}}")

        # Salto de página: por default después de cada nivel 1, excepto el último elemento
        if break_after_level1 and lvl == 1 and i < total - 1:
            doc.add_page_break()


# -------------------------
# CATALOGO (METADATOS)
# -------------------------

def update_catalog(base_dir: str, cfg: dict, output_path: str):
    catalog_entry = {
        "id": cfg["id"],
        "universidad": cfg.get("universidad"),
        "tipo": cfg.get("tipo"),
        "version": cfg.get("version"),
        "descripcion": cfg.get("descripcion"),
        "file": os.path.basename(output_path),
        "format": "docx",
        "size_bytes": os.path.getsize(output_path),
        "last_modified": datetime.fromtimestamp(os.path.getmtime(output_path)).isoformat(timespec="seconds")
    }

    catalog_path = os.path.join(base_dir, "catalog.json")

    if os.path.exists(catalog_path):
        with open(catalog_path, "r", encoding="utf-8") as f:
            catalog = json.load(f)
        if not isinstance(catalog, list):
            catalog = []
    else:
        catalog = []

    # upsert
    catalog = [x for x in catalog if x.get("id") != catalog_entry["id"]]
    catalog.append(catalog_entry)

    with open(catalog_path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)

    print("CATALOG OK ->", catalog_path)


# -------------------------
# MAIN
# -------------------------

def generate(config_path: str):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = resolve_config_path(base_dir, config_path)
    if not os.path.exists(config_path):
        print("Uso: python generate_from_json.py <ruta_config.json>")
        print("No se encontro el archivo:", config_path)
        raise SystemExit(1)
    cfg = load_json(config_path)

    # Documento
    doc = Document()
    set_page_setup(doc, cfg)

    # Carátula
    add_cover_from_cfg(doc, cfg, base_dir)

    # Índices
    add_toc_page(doc, cfg.get("toc", {"min_level": 1, "max_level": 3}))
    if cfg.get("include_list_of_tables", False):
        add_list_of_tables(doc)
    if cfg.get("include_list_of_figures", False):
        add_list_of_figures(doc)

    # Estructura
    add_structure_from_cfg(doc, cfg)

    # Guardar
    output_name = cfg.get("output_name", "output.docx")
    output_path = resolve_path(base_dir, output_name)
    doc.save(output_path)

    print("DOCX OK ->", output_path)

    # Actualizar catálogo (metadatos reales)
    update_catalog(base_dir, cfg, output_path)

    if cfg.get("open_after_generate", True):
        open_document(output_path)


if __name__ == "__main__":
    # Ejemplo:
    # python generate_from_json.py formats/unac_maestria_cuant.json
    import sys

    default_config = os.path.join("formats", "unac_maestria_cuant.json")
    config_path = sys.argv[1] if len(sys.argv) > 1 else default_config
    generate(config_path)

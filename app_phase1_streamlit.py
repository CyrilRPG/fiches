import io, zipfile, re
import streamlit as st
from typing import List, Tuple, Dict, Optional
from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table
from docx.oxml.ns import qn

OLD = "2024-2025"
NEW = "2025-2026"

st.set_page_config(page_title="Fiches – Phase 1 (macOS) : texte & styles", page_icon="📝", layout="centered")
st.title("📝 Traitement Fiches – Phase 1 (macOS) : texte & styles")

st.markdown("""
**Ce que fait l’outil (Phase 1)**  
- Remplace **2024-2025 → 2025-2026** (et normalise les variantes avec espaces).  
- Force **Calibri** partout.  
- Applique les tailles :  
  - Paragraphes courants : **9 pt**  
  - **Université + année** (y compris pied de page) : **10 pt**  
  - **Titre du cours** : **20 pt**  
  - **Matière** : **18 pt**  
  - **“Fiche de cours …”** : **22 pt**  
  - **Titre de tableau** : **12 pt gras**  
  - **Numérotation de tableau** (ex.: *I. Introduction générale*) : **10 pt gras italique**  
  
ℹ️ *La détection “Titre du cours” et “Matière” est automatique.*
""")

uploaded = st.file_uploader("Dépose un ou plusieurs fichiers .docx", type=["docx"], accept_multiple_files=True)

ROMAN_LINE = re.compile(r"^\s*[IVXLC]+\.[\s\S]+", flags=re.IGNORECASE)
YEAR_VARIANTS = [
    "2024-2025", "2024 - 2025", "2024\u00A0-\u00A02025"  # include NBSP
]

def set_run_font(run, size_pt: Optional[float]=None, name: str="Calibri", bold: Optional[bool]=None, italic: Optional[bool]=None):
    if name:
        run.font.name = name
        r = run._element.rPr
        if r is not None:
            r_fonts = r.rFonts
            if r_fonts is not None:
                r_fonts.set(qn('w:ascii'), name)
                r_fonts.set(qn('w:hAnsi'), name)
                r_fonts.set(qn('w:cs'), name)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic

def replace_in_paragraph(par: Paragraph) -> int:
    count = 0
    for run in par.runs:
        text = run.text
        for variant in YEAR_VARIANTS:
            if variant in text:
                count += text.count(variant)
                text = text.replace(variant, NEW)
        if OLD in text:
            count += text.count(OLD)
            text = text.replace(OLD, NEW)
        run.text = text
    return count

def set_paragraph_font(par: Paragraph, size_pt: float, bold: Optional[bool]=None, italic: Optional[bool]=None):
    for run in par.runs:
        set_run_font(run, size_pt=size_pt, bold=bold, italic=italic)

def process_paragraph(par: Paragraph, context: Dict) -> Dict:
    txt = par.text.strip()
    counters = {k:0 for k in ["body9","univ10","course20","subject18","fiche22","table_title12","table_num10","replacements"]}
    counters["replacements"] += replace_in_paragraph(par)
    for run in par.runs:
        set_run_font(run, name="Calibri")

    if txt and "fiche de cours" in txt.lower():
        set_paragraph_font(par, 22.0)
        counters["fiche22"] += 1
        return counters

    if txt and (("université" in txt.lower()) or (NEW in txt)):
        set_paragraph_font(par, 10.0)
        counters["univ10"] += 1
        return counters

    if ROMAN_LINE.match(txt):
        set_paragraph_font(par, 10.0, bold=True, italic=True)
        counters["table_num10"] += 1
        context["prev_was_table_num"] = True
        return counters

    if txt:
        set_paragraph_font(par, 9.0)
        counters["body9"] += 1

    context["last_non_empty_par"] = par
    return counters

def mark_table_title(prev_par: Optional[Paragraph], counters: Dict):
    if prev_par is None:
        return
    if prev_par.text.strip():
        set_paragraph_font(prev_par, 12.0, bold=True, italic=False)
        counters["table_title12"] += 1

def detect_and_set_titles(doc: Document, counters: Dict):
    lowered = lambda s: s.strip().lower() if s else ""
    paragraphs = [p for p in doc.paragraphs]
    candidates = [p for p in paragraphs if p.text and len(p.text.strip()) > 2]

    course_par = None
    for p in candidates[:20]:
        l = lowered(p.text)
        if ("fiche de cours" in l) or ("université" in l) or (NEW in l) or ("2024" in l):
            continue
        course_par = p
        break
    if course_par is not None:
        set_paragraph_font(course_par, 20.0)
        counters["course20"] += 1

    subject_par = None
    if course_par is not None:
        try:
            idx = paragraphs.index(course_par)
            for p in paragraphs[idx+1: idx+6]:
                if p.text and len(p.text.strip()) > 1 and "fiche de cours" not in lowered(p.text):
                    subject_par = p
                    break
        except ValueError:
            pass
    if subject_par is None:
        for p in candidates:
            l = lowered(p.text)
            if any(k in l for k in ["matière", "discipline", "ue ", "unité d'enseignement"]):
                subject_par = p
                break
    if subject_par is not None:
        set_paragraph_font(subject_par, 18.0)
        counters["subject18"] += 1

def walk_tables(tables: List[Table], process_fn):
    for t in tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_fn(p)
                if cell.tables:
                    walk_tables(cell.tables, process_fn)

def process_headers_footers(doc: Document, counters: Dict, force_footer_10: bool = True):
    for section in doc.sections:
        for p in section.header.paragraphs:
            c = process_paragraph(p, context={})
            for k,v in c.items():
                counters[k] += v
        for p in section.footer.paragraphs:
            c = process_paragraph(p, context={})
            for k,v in c.items():
                counters[k] += v
            if force_footer_10 and p.text.strip():
                set_paragraph_font(p, 10.0)

def process_document(doc_bytes: bytes) -> Tuple[bytes, Dict]:
    doc = Document(io.BytesIO(doc_bytes))
    counters = {k:0 for k in ["body9","univ10","course20","subject18","fiche22","table_title12","table_num10","replacements"]}
    context = {"prev_was_table_num": False, "last_non_empty_par": None}

    def handle_par(par: Paragraph):
        nonlocal context, counters
        c = process_paragraph(par, context)
        if context.get("prev_was_table_num", False):
            mark_table_title(context.get("last_non_empty_par"), counters)
            context["prev_was_table_num"] = False
        for k,v in c.items():
            counters[k] += v

    for p in doc.paragraphs:
        handle_par(p)

    walk_tables(doc.tables, handle_par)
    process_headers_footers(doc, counters, force_footer_10=True)

    detect_and_set_titles(doc, counters)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue(), counters

if uploaded:
    rows = []
    outs = []
    with st.spinner("Traitement en cours..."):
        for f in uploaded:
            data = f.read()
            out_bytes, counters = process_document(data)
            outs.append((f.name, out_bytes, counters))
            rows.append({"fichier": f.name, **counters})
    st.success("Terminé ✅")
    st.subheader("Résumé")
    st.dataframe(rows, use_container_width=True)

    st.subheader("Téléchargements")
    for name, out_bytes, _ in outs:
        st.download_button(
            label=f"Télécharger {name} (traité)",
            data=out_bytes,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl-{name}"
        )

    if len(outs) > 1:
        zip_bytes = io.BytesIO()
        with zipfile.ZipFile(zip_bytes, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name, out_bytes, _ in outs:
                zf.writestr(name, out_bytes)
        zip_bytes.seek(0)
        st.download_button(
            label="Télécharger tous les fichiers (ZIP)",
            data=zip_bytes.getvalue(),
            file_name="fiches_traitees_phase1.zip",
            mime="application/zip",
            key="dl-zip"
        )

st.caption("Phase 1 : texte & styles (python-docx). Les formes/images de la page de garde seront traitées en Phase 2.")

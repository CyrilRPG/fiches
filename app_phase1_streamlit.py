# -*- coding: utf-8 -*-
import io
import zipfile
import re
import os
import unicodedata
import xml.etree.ElementTree as ET
from typing import Dict, Tuple, List
import streamlit as st

# ───────────────────────── Espaces de noms ─────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
VML_NS = "urn:schemas-microsoft-com:vml"

NS = {"w": W, "wp": WP, "a": A, "pic": PIC, "r": R, "wps": WPS, "v": VML_NS}
for k, v in NS.items():
    ET.register_namespace(k, v)

# ───────────────────────── Règles/constantes ───────────────────────
TARGETS = ["2024-2025", "2024 - 2025", "2024\u00A0-\u00A02025"]
REPL = "2025 - 2026"
ROMAN_RE = re.compile(r"^\s*[IVXLC]+\s*[.)]?\s+.+", re.IGNORECASE)

# ───────────────────────── Utils ───────────────────────────────────
def cm_to_emu(cm: float) -> int:
    return int(round(cm * 360000))

def emu_to_cm(emu: int) -> float:
    return emu / 360000.0

def get_text(p) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))

def set_run_props(run, size=None, bold=None, italic=None, color=None, calibri=False):
    """
    Applique sélectivement certaines propriétés de mise en forme.
    On **n’impose plus Calibri** globalement afin de préserver le design du modèle.
    """
    rPr = run.find("w:rPr", NS) or ET.SubElement(run, f"{{{W}}}rPr")
    if size is not None:
        v = str(int(round(size * 2)))
        (rPr.find("w:sz", NS) or ET.SubElement(rPr, f"{{{W}}}sz")).set(f"{{{W}}}val", v)
        (rPr.find("w:szCs", NS) or ET.SubElement(rPr, f"{{{W}}}szCs")).set(f"{{{W}}}val", v)
    if bold is not None:
        b = rPr.find("w:b", NS)
        if bold:
            (b or ET.SubElement(rPr, f"{{{W}}}b")).set(f"{{{W}}}val", "1")
        elif b is not None:
            rPr.remove(b)
    if italic is not None:
        i = rPr.find("w:i", NS)
        if italic:
            (i or ET.SubElement(rPr, f"{{{W}}}i")).set(f"{{{W}}}val", "1")
        elif i is not None:
            rPr.remove(i)
    if color is not None:
        (rPr.find("w:color", NS) or ET.SubElement(rPr, f"{{{W}}}color")).set(f"{{{W}}}val", color)

def set_dml_text_size_in_txbody(txbody, pt: float):
    val = str(int(round(pt * 100)))
    for r in txbody.findall(".//a:r", NS):
        rPr = r.find("a:rPr", NS) or ET.SubElement(r, f"{{{A}}}rPr")
        rPr.set("sz", val)

def redistribute(nodes, new):
    lens = [len(n.text or "") for n in nodes]
    pos = 0
    for i, n in enumerate(nodes):
        n.text = new[pos:pos + lens[i]] if i < len(nodes) - 1 else new[pos:]
        if i < len(nodes) - 1:
            pos += lens[i]

def normalize_spaces(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" - ", " - ")
    return s

# ───────────────────────── Remplacements texte ─────────────────────
def replace_years(root: ET.Element):
    for p in root.findall(".//w:p", NS):
        wts = p.findall(".//w:t", NS)
        if not wts:
            continue
        txt = "".join(t.text or "" for t in wts)
        new = txt
        for tgt in TARGETS:
            new = new.replace(tgt, REPL)
        if new != txt:
            redistribute(wts, new)
    ats = root.findall(".//a:txBody//a:t", NS)
    if ats:
        txt = "".join(t.text or "" for t in ats)
        new = txt
        for tgt in TARGETS:
            new = new.replace(tgt, REPL)
        if new != txt:
            redistribute(ats, new)

def strip_actualisation_everywhere(root: ET.Element):
    for t in root.findall(".//w:t", NS) + root.findall(".//a:t", NS):
        if t.text:
            t.text = re.sub(r"(?iu)\bactualisation\b", "", t.text)

def red_to_black(root: ET.Element):
    for r in root.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS)
        c = None if rPr is None else rPr.find("w:color", NS)
        if c is not None and (c.get(f"{{{W}}}val", "").upper() == "FF0000"):
            c.set(f"{{{W}}}val", "000000")

# ───────────────────────── Mise en forme couverture ────────────────
def holder_pos_cm(holder: ET.Element) -> Tuple[float, float]:
    try:
        x = int((holder.find("wp:positionH/wp:posOffset", NS) or ET.Element("x")).text or "0")
        y = int((holder.find("wp:positionV/wp:posOffset", NS) or ET.Element("y")).text or "0")
        return (emu_to_cm(x), emu_to_cm(y))
    except Exception:
        return (0.0, 0.0)

def get_tx_text(holder: ET.Element) -> str:
    tx = holder.find(".//a:txBody", NS)
    if tx is not None:
        return "".join(t.text or "" for t in tx.findall(".//a:t", NS)).strip()
    txbx = holder.find(".//wps:txbx/w:txbxContent", NS)
    if txbx is not None:
        return "".join(t.text or "" for t in txbx.findall(".//w:t", NS)).strip()
    return ""

def set_tx_size(holder: ET.Element, pt: float):
    tx = holder.find(".//a:txBody", NS)
    if tx is not None:
        set_dml_text_size_in_txbody(tx, pt)
    txbx = holder.find(".//wps:txbx/w:txbxContent", NS)
    if txbx is not None:
        for r in txbx.findall(".//w:r", NS):
            set_run_props(r, size=pt)

def cover_sizes_cleanup(root: ET.Element):
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]

    def set_size(p, pt):
        for r in p.findall(".//w:r", NS):
            set_run_props(r, size=pt)

    last_was_fiche = False
    in_plan = False

    for i, txt in enumerate(texts):
        low = txt.lower()

        if txt.strip().upper() == "ACTUALISATION":
            for t in paras[i].findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)actualisation", "", t.text)
            continue

        if "fiche de cours" in low:
            set_size(paras[i], 22)
            last_was_fiche = True
            in_plan = False
            continue
        if last_was_fiche and txt:
            set_size(paras[i], 20)
            last_was_fiche = False

        if "université" in low and any(
            x.replace("\u00A0", " ") in txt.replace("\u00A0", " ") for x in TARGETS + [REPL]
        ):
            set_size(paras[i], 10)

        if txt.strip().upper() == "PLAN":
            set_size(paras[i], 11)
            in_plan = True
            continue
        if in_plan:
            if txt.strip() == "" or low.startswith("légende"):
                in_plan = False
            else:
                set_size(paras[i], 11)

def tune_cover_shapes_spatial(root: ET.Element):
    holders = []
    for holder in root.findall(".//wp:anchor", NS) + root.findall(".//wp:inline", NS):
        txt = get_tx_text(holder)
        if not txt:
            continue
        x, y = holder_pos_cm(holder)
        holders.append((y, x, holder, txt))

    if not holders:
        return

    holders.sort(key=lambda t: (t[0], t[1]))

    # université + année en 10
    for _, _, h, txt in holders:
        low = txt.lower()
        if ("universite" in low or "université" in low) and any(
            t.replace("\u00A0", " ") in txt.replace("\u00A0", " ") for t in TARGETS + [REPL]
        ):
            set_tx_size(h, 10.0)

    # Fiche de cours (22) + cours (20)
    idx_fiche = None
    for i, (_, _, h, txt) in enumerate(holders):
        if "fiche de cours" in txt.lower():
            set_tx_size(h, 22.0)
            idx_fiche = i
            break
    if idx_fiche is not None:
        for j in range(idx_fiche + 1, len(holders)):
            txt = holders[j][3].strip()
            if txt and "fiche de cours" not in txt.lower():
                set_tx_size(holders[j][2], 20.0)
                break

    # PLAN -> tout en 11
    for _, _, h, txt in holders:
        if "plan" in txt.lower():
            set_tx_size(h, 11.0)

    # retirer ACTUALISATION éventuel dans les formes
    for _, _, h, _ in holders:
        tx = h.find(".//a:txBody", NS)
        if tx is not None:
            for t in tx.findall(".//a:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)\bactualisation\b", "", t.text)
        txbx = h.find(".//wps:txbx/w:txbxContent", NS)
        if txbx is not None:
            for t in txbx.findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)\bactualisation\b", "", t.text)

def tables_and_numbering(root: ET.Element):
    for tbl in root.findall(".//w:tbl", NS):
        rows = tbl.findall(".//w:tr", NS)
        if not rows:
            continue
        for p in rows[0].findall(".//w:p", NS):
            for r in p.findall(".//w:r", NS):
                set_run_props(r, size=12, bold=True)
        for tr in rows[1:]:
            for p in tr.findall(".//w:p", NS):
                for r in p.findall(".//w:r", NS):
                    set_run_props(r, size=9)
    for p in root.findall(".//w:p", NS):
        if ROMAN_RE.match(get_text(p).strip() or ""):
            for r in p.findall(".//w:r", NS):
                set_run_props(r, size=10, bold=True, italic=True, color="FFFFFF")

# ───────────────────────── Suppression du panneau gris ─────────────
def extract_theme_colors(parts: Dict[str, bytes]) -> Dict[str, str]:
    data = parts.get("word/theme/theme1.xml")
    if not data:
        return {}
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return {}
    colors: Dict[str, str] = {}
    cs = root.find(".//a:clrScheme", NS)
    if cs is None:
        return colors
    for el in list(cs):
        tag = re.sub(r"{.*}", "", el.tag)
        srgb = el.find("a:srgbClr", NS)
        if srgb is not None:
            colors[tag.lower()] = srgb.get("val", "").upper()
        else:
            sys = el.find("a:sysClr", NS)
            if sys is not None:
                colors[tag.lower()] = sys.get("lastClr", "").upper()
    return colors

def remove_large_grey_rectangles(root: ET.Element, theme_colors: Dict[str, str]):
    parent_map = {child: parent for parent in root.iter() for child in parent}

    for drawing in root.findall(".//w:drawing", NS):
        holder = drawing.find(".//wp:anchor", NS) or drawing.find(".//wp:inline", NS)
        if holder is None:
            continue
        extent = holder.find("wp:extent", NS)
        if extent is None:
            continue
        try:
            cx = int(extent.get("cx", "0"))
            cy = int(extent.get("cy", "0"))
        except Exception:
            continue

        if holder.find(".//pic:pic", NS) is not None:
            continue

        # calcul des positions
        x_el = holder.find(".//wp:positionH/wp:posOffset", NS)
        y_el = holder.find(".//wp:positionV/wp:posOffset", NS)
        try:
            x_cm = emu_to_cm(int(x_el.text)) if x_el is not None else 0.0
            y_cm = emu_to_cm(int(y_el.text)) if y_el is not None else 0.0
        except Exception:
            x_cm = y_cm = 0.0

        # type de la géométrie
        is_rect = holder.find(".//a:prstGeom[@prst='rect']", NS) is not None or \
                  holder.find(".//a:prstGeom[@prst='roundRect']", NS) is not None
        if not is_rect:
            continue

        # couleurs explicites (srgbClr)
        srgb_colors = [el.get("val", "").upper() for el in holder.findall(".//a:srgbClr", NS)]
        looks_f2 = False
        for v in srgb_colors:
            try:
                r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
                if abs(r - 0xF2) <= 16 and abs(g - 0xF2) <= 16 and abs(b - 0xF2) <= 16:
                    looks_f2 = True
                    break
            except Exception:
                pass

        width_cm = emu_to_cm(cx)
        height_cm = emu_to_cm(cy)
        big_on_cover = y_cm < 20.0 and width_cm >= 6.0 and height_cm >= 8.0
        very_big_anywhere = width_cm >= 10.0 and height_cm >= 12.0
        very_big_right = (x_cm >= 9.0) and (width_cm >= 7.0) and (height_cm >= 12.0)

        if ((srgb_colors and looks_f2 and (big_on_cover or x_cm >= 9.0)) or
            very_big_anywhere or
            very_big_right):
            parent = parent_map.get(drawing)
            if parent is not None:
                parent.remove(drawing)

# ───────────────────────── Traitement du document ──────────────────
def process_bytes(
    docx_bytes: bytes,
    legend_bytes: bytes = None,
    icon_left=15.3,
    icon_top=11.0,
    legend_left=2.25,
    legend_top=25.0,
    legend_w=5.68,
    legend_h=3.77,
) -> bytes:

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        parts: Dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}
    theme_colors = extract_theme_colors(parts)

    for name, data in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue

        # texte : remplacements
        replace_years(root)
        strip_actualisation_everywhere(root)
        red_to_black(root)

        if name == "word/document.xml":
            cover_sizes_cleanup(root)
            tune_cover_shapes_spatial(root)
            tables_and_numbering(root)
            # pas de repositionnement forcé de l’icône d’écriture
            remove_large_grey_rectangles(root, theme_colors)

        # on ne touche plus aux pieds de page : la numérotation reste intacte

        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # Légende optionnelle
    if legend_bytes and "word/document.xml" in parts and "word/_rels/document.xml.rels" in parts:
        parts["word/document.xml"] = remove_legend_text(parts["word/document.xml"])
        new_doc, new_rels, media = insert_legend_image(
            parts["word/document.xml"],
            parts["word/_rels/document.xml.rels"],
            legend_bytes,
            left_cm=legend_left,
            top_cm=legend_top,
            width_cm=legend_w,
            height_cm=legend_h,
        )
        parts["word/document.xml"] = new_doc
        parts["word/_rels/document.xml.rels"] = new_rels
        parts[media[0]] = media[1]

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, d in parts.items():
            zout.writestr(n, d)
    return out_buf.getvalue()

# ───────────────────────── Nom de fichier de sortie ────────────────
def cleaned_filename(original_name: str) -> str:
    base, ext = os.path.splitext(original_name)
    base = re.sub(r"(?iu)\bactu\b", "", base)
    base = normalize_spaces(base)
    base = re.sub(r"\s+([\-_,])", r"\1", base)
    if not ext.lower().endswith(".docx"):
        ext = ".docx"
    return f"{base}{ext}"

# ───────────────────────── Interface Streamlit ─────────────────────
st.set_page_config(page_title="Fiches Diploma", page_icon="🧠", layout="centered")
st.title("🧠 Fiches Diploma")
st.caption("Transforme tes .docx des fiches actualisées de 2024-2025 en fiches de cours non actualisées de 2025-2026 en respectant toutes les règles")

with st.sidebar:
    st.subheader("Paramètres (cm)")
    icon_left = st.number_input("Icône écriture — gauche", value=15.3, step=0.1)
    icon_top = st.number_input("Icône écriture — haut", value=11.0, step=0.1)
    legend_left = st.number_input("Image Légendes — gauche", value=2.25, step=0.1)
    legend_top = st.number_input("Image Légendes — haut", value=25.0, step=0.1)
    legend_w = st.number_input("Image Légendes — largeur", value=5.68, step=0.01)
    legend_h = st.number_input("Image Légendes — hauteur", value=3.77, step=0.01)

st.markdown("**1) Glisse/dépose un ou plusieurs fichiers `.docx`**")
files = st.file_uploader("DOCX à traiter", type=["docx"], accept_multiple_files=True)
st.markdown("**2) (Optionnel) Ajoute l’image de la Légende (PNG/JPG)**")
legend_file = st.file_uploader("Image Légendes", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=False)

if st.button("⚙️ Lancer le traitement", type="primary", disabled=not files):
    if not files:
        st.warning("Ajoute au moins un fichier .docx")
    else:
        legend_bytes = legend_file.read() if legend_file else None
        for up in files:
            try:
                out_bytes = process_bytes(
                    up.read(),
                    legend_bytes=legend_bytes,
                    icon_left=icon_left,
                    icon_top=icon_top,
                    legend_left=legend_left,
                    legend_top=legend_top,
                    legend_w=legend_w,
                    legend_h=legend_h,
                )
                out_name = cleaned_filename(up.name)
                st.success(f"✅ Terminé : {up.name}")
                st.download_button(
                    "⬇️ Télécharger " + out_name,
                    data=out_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"❌ Échec pour {up.name} : {e}")

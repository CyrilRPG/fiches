# app_phase1_streamlit.py
import io, zipfile, re, xml.etree.ElementTree as ET
from typing import Dict, Tuple
import streamlit as st

# ---------- Namespaces ----------
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
VML_NS = "urn:schemas-microsoft-com:vml"

NS = {"w":W, "wp":WP, "a":A, "pic":PIC, "r":R, "wps":WPS, "v":VML_NS}
for k,v in NS.items(): ET.register_namespace(k, v)

# ---------- Rules ----------
TARGETS = ["2024-2025","2024 - 2025","2024\u00A0-\u00A02025"]
REPL    = "2025 - 2026"
ROMAN_RE = re.compile(r"^\s*[IVXLC]+\s*[.)]?\s+.+", re.IGNORECASE)

# ---------- Utils ----------
def cm_to_emu(cm: float) -> int: return int(round(cm * 360000))
def emu_to_cm(emu: int) -> float: return emu/360000.0
def get_text(p) -> str: return "".join(t.text or "" for t in p.findall(".//w:t", NS))

def set_run_props(run, size=None, bold=None, italic=None, color=None, calibri=False):
    rPr = run.find("w:rPr", NS) or ET.SubElement(run, f"{{{W}}}rPr")
    if calibri:
        rFonts = rPr.find("w:rFonts", NS) or ET.SubElement(rPr, f"{{{W}}}rFonts")
        for k in ("ascii","hAnsi","cs"): rFonts.set(f"{{{W}}}{k}", "Calibri")
    if size is not None:
        v = str(int(round(size*2)))
        (rPr.find("w:sz", NS)  or ET.SubElement(rPr, f"{{{W}}}sz")).set(f"{{{W}}}val", v)
        (rPr.find("w:szCs", NS) or ET.SubElement(rPr, f"{{{W}}}szCs")).set(f"{{{W}}}val", v)
    if bold is not None:
        b = rPr.find("w:b", NS)
        if bold:  (b or ET.SubElement(rPr, f"{{{W}}}b")).set(f"{{{W}}}val","1")
        elif b is not None: rPr.remove(b)
    if italic is not None:
        i = rPr.find("w:i", NS)
        if italic: (i or ET.SubElement(rPr, f"{{{W}}}i")).set(f"{{{W}}}val","1")
        elif i is not None: rPr.remove(i)
    if color is not None:
        (rPr.find("w:color", NS) or ET.SubElement(rPr, f"{{{W}}}color")).set(f"{{{W}}}val", color)

def set_dml_text_size_in_txbody(txbody, pt: float):
    """Set size in DrawingML text body (pt -> *100)."""
    val = str(int(round(pt*100)))
    for r in txbody.findall(".//a:r", NS):
        rPr = r.find("a:rPr", NS) or ET.SubElement(r, f"{{{A}}}rPr")
        rPr.set("sz", val)

def redistribute(nodes, new):
    lens = [len(n.text or "") for n in nodes]; pos = 0
    for i,n in enumerate(nodes):
        n.text = new[pos:pos+lens[i]] if i < len(nodes)-1 else new[pos:]
        if i < len(nodes)-1: pos += lens[i]

# ---------- Text replacements & styling ----------
def replace_years(root):
    # w:t
    for p in root.findall(".//w:p", NS):
        wts = p.findall(".//w:t", NS)
        if not wts: continue
        txt = "".join(t.text or "" for t in wts); new = txt
        for tgt in TARGETS: new = new.replace(tgt, REPL)
        if new != txt: redistribute(wts, new)
    # a:t (shapes)
    ats = root.findall(".//a:txBody//a:t", NS)
    if ats:
        txt = "".join(t.text or "" for t in ats); new = txt
        for tgt in TARGETS: new = new.replace(tgt, REPL)
        if new != txt: redistribute(ats, new)

def force_calibri(root):
    for r in root.findall(".//w:r", NS): set_run_props(r, calibri=True)

def red_to_black(root):
    for r in root.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS);  c = None if rPr is None else rPr.find("w:color", NS)
        if c is not None and (c.get(f"{{{W}}}val","").upper() == "FF0000"): c.set(f"{{{W}}}val","000000")

def cover_sizes_cleanup(root):
    """Keep layout as-is but enforce sizes:
    - University + year: 10
    - 'Fiche de cours …': 22
    - Course name (next non-empty para after 'Fiche …'): 20
    - PLAN (title + items on cover): 11
    - Remove 'ACTUALISATION'
    """
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]
    def set_size(p, pt):  [set_run_props(r, size=pt) for r in p.findall(".//w:r", NS)]

    last_was_fiche = False
    in_plan = False

    for i, txt in enumerate(texts):
        low = txt.lower()

        # remove ACTUALISATION
        if txt.strip().upper() == "ACTUALISATION":
            for t in paras[i].findall(".//w:t", NS):
                if t.text: t.text = re.sub(r"(?i)ACTUALISATION","",t.text)
            continue

        # sizes
        if "fiche de cours" in low:
            set_size(paras[i], 22)
            last_was_fiche = True
            in_plan = False
            continue

        if last_was_fiche and txt:
            set_size(paras[i], 20)  # course name
            last_was_fiche = False

        if "université" in low and (any(x.replace("\u00A0"," ") in txt.replace("\u00A0"," ") for x in TARGETS+[REPL])):
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

def tables_and_numbering(root):
    # table title 12 bold; body 9
    for tbl in root.findall(".//w:tbl", NS):
        rows = tbl.findall(".//w:tr", NS)
        if not rows: continue
        for p in rows[0].findall(".//w:p", NS):
            for r in p.findall(".//w:r", NS): set_run_props(r, size=12, bold=True)
        for tr in rows[1:]:
            for p in tr.findall(".//w:p", NS):
                for r in p.findall(".//w:r", NS): set_run_props(r, size=9)
    # roman numbering line
    for p in root.findall(".//w:p", NS):
        if ROMAN_RE.match(get_text(p).strip() or ""):
            for r in p.findall(".//w:r", NS):
                set_run_props(r, size=10, bold=True, italic=True, color="FFFFFF")

# ---------- Grey cover rectangle removal (robust) ----------
def build_parent_map(root): return {child: parent for parent in root.iter() for child in parent}
def hex_to_rgb(h): h=h.strip().lstrip("#"); return (int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)) if len(h)==6 else (0,0,0)
def is_close_grey(val, target="#F2F2F2", tol=16):
    try:
        r,g,b   = hex_to_rgb(val.upper()); rt,gt,bt = hex_to_rgb(target)
        return abs(r-rt)<=tol and abs(g-gt)<=tol and abs(b-bt)<=tol
    except: return False

def extract_theme_colors(parts: Dict[str, bytes]) -> Dict[str, str]:
    data = parts.get("word/theme/theme1.xml")
    if not data: return {}
    try: root = ET.fromstring(data)
    except ET.ParseError: return {}
    colors = {}
    cs = root.find(".//a:clrScheme", NS)
    if cs is None: return colors
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

def looks_like_cover_shape(holder):
    posV = holder.find("wp:positionV/wp:posOffset", NS)
    if posV is None: return True
    try:  return emu_to_cm(int(posV.text)) < 20.0
    except: return True

def remove_large_grey_rectangles(root, theme_colors: Dict[str, str]):
    parent_map = build_parent_map(root)

    # DrawingML shapes
    for drawing in root.findall(".//w:drawing", NS):
        holder = drawing.find(".//wp:anchor", NS) or drawing.find(".//wp:inline", NS)
        if holder is None: continue
        extent = holder.find("wp:extent", NS)
        if extent is None: continue
        try:
            cx = int(extent.get("cx","0")); cy = int(extent.get("cy","0"))
        except: continue

        has_pic = holder.find(".//pic:pic", NS) is not None

        # gather colors (srgb + scheme with luminance)
        srgb = [el.get("val","").upper() for el in holder.findall(".//a:srgbClr", NS)]
        for sc in holder.findall(".//a:schemeClr", NS):
            val = sc.get("val", "").lower()
            base = theme_colors.get(val)
            if not base: continue
            r, g, b = hex_to_rgb(base)
            lm = sc.find("a:lumMod", NS); lo = sc.find("a:lumOff", NS)
            mod = int(lm.get("val", "100000")) / 100000 if lm is not None else 1.0
            off = int(lo.get("val", "0")) / 100000 if lo is not None else 0.0
            r = min(255, int(r * mod + 255 * off))
            g = min(255, int(g * mod + 255 * off))
            b = min(255, int(b * mod + 255 * off))
            srgb.append(f"{r:02X}{g:02X}{b:02X}")

        looks_f2 = any(is_close_grey(v, "#F2F2F2", 16) for v in srgb)
        is_rect = holder.find(".//a:prstGeom[@prst='rect']", NS) is not None \
               or holder.find(".//a:prstGeom[@prst='roundRect']", NS) is not None

        big_on_cover = looks_like_cover_shape(holder) and cx >= cm_to_emu(6) and cy >= cm_to_emu(8)
        very_big_anywhere = cx >= cm_to_emu(10) and cy >= cm_to_emu(10)

        if (not has_pic) and looks_f2 and is_rect and (big_on_cover or very_big_anywhere):
            parent = parent_map.get(drawing)
            if parent is not None: parent.remove(drawing)

    # VML (old shapes)
    for pict in root.findall(".//w:pict", NS):
        # rect / roundrect / shape
        for tag in ("rect","roundrect","shape"):
            for shape in pict.findall(f".//v:{tag}", NS):
                fill = (shape.get("fillcolor", "") or "").upper()
                if not is_close_grey(fill, "#F2F2F2", 16):  # only the light grey cover box
                    continue
                style = shape.get("style", "")
                m_w = re.search(r"width:([0-9.]+)cm", style)
                m_h = re.search(r"height:([0-9.]+)cm", style)
                if m_w and m_h:
                    w = float(m_w.group(1)); h = float(m_h.group(1))
                    big_on_cover = w >= 6 and h >= 8
                    very_big_anywhere = w >= 10 and h >= 10
                    if big_on_cover or very_big_anywhere:
                        parent = parent_map.get(pict)
                        if parent is not None: parent.remove(pict)

# ---------- Legend image ----------
def build_anchored_image(rId, width_cm, height_cm, left_cm, top_cm, name="Legende"):
    cx, cy = cm_to_emu(width_cm), cm_to_emu(height_cm)
    xoff, yoff = cm_to_emu(left_cm), cm_to_emu(top_cm)
    drawing = ET.Element(f"{{{W}}}drawing")
    anchor = ET.SubElement(drawing, f"{{{WP}}}anchor", {
        "distT":"0","distB":"0","distL":"0","distR":"0",
        "simplePos":"0","relativeHeight":"0","behindDoc":"0",
        "locked":"0","layoutInCell":"1","allowOverlap":"1"
    })
    ET.SubElement(anchor, f"{{{WP}}}simplePos", {"x":"0","y":"0"})
    posH = ET.SubElement(anchor, f"{{{WP}}}positionH", {"relativeFrom":"page"})
    ET.SubElement(posH, f"{{{WP}}}posOffset").text = str(xoff)
    posV = ET.SubElement(anchor, f"{{{WP}}}positionV", {"relativeFrom":"page"})
    ET.SubElement(posV, f"{{{WP}}}posOffset").text = str(yoff)
    ET.SubElement(anchor, f"{{{WP}}}extent", {"cx":str(cx),"cy":str(cy)})
    ET.SubElement(anchor, f"{{{WP}}}effectExtent", {"l":"0","t":"0","r":"0","b":"0"})
    ET.SubElement(anchor, f"{{{WP}}}wrapNone")
    ET.SubElement(anchor, f"{{{WP}}}docPr", {"id":"10","name":name})
    ET.SubElement(anchor, f"{{{WP}}}cNvGraphicFramePr")
    graphic = ET.SubElement(anchor, f"{{{A}}}graphic")
    gData = ET.SubElement(graphic, f"{{{A}}}graphicData", {"uri":"http://schemas.openxmlformats.org/drawingml/2006/picture"})
    pic = ET.SubElement(gData, f"{{{PIC}}}pic")
    nvPicPr = ET.SubElement(pic, f"{{{PIC}}}nvPicPr")
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPr", {"id":"0","name":name+".img"})
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPicPr")
    blipFill = ET.SubElement(pic, f"{{{PIC}}}blipFill")
    ET.SubElement(blipFill, f"{{{A}}}blip", {f"{{{R}}}embed": rId})
    stretch = ET.SubElement(blipFill, f"{{{A}}}stretch"); ET.SubElement(stretch, f"{{{A}}}fillRect")
    spPr = ET.SubElement(pic, f"{{{PIC}}}spPr")
    xfrm = ET.SubElement(spPr, f"{{{A}}}xfrm")
    ET.SubElement(xfrm, f"{{{A}}}off", {"x":"0","y":"0"})
    ET.SubElement(xfrm, f"{{{A}}}ext", {"cx":str(cx), "cy":str(cy)})
    prst = ET.SubElement(spPr, f"{{{A}}}prstGeom", {"prst":"rect"}); ET.SubElement(prst, f"{{{A}}}avLst")
    return drawing

def remove_legend_text(document_xml: bytes) -> bytes:
    root = ET.fromstring(document_xml)
    # "Légendes" title
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip().lower() == "légendes":
            for t in p.findall(".//w:t", NS): t.text = ""
    # lines
    lines = {"Notion nouvelle cette année","Notion hors programme","Notion déjà tombée au concours","Astuces et méthodes"}
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip() in lines:
            for t in p.findall(".//w:t", NS): t.text = ""
            for d in p.findall(".//w:drawing", NS): d.clear()
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def insert_legend_image(document_xml: bytes, rels_xml: bytes, image_bytes: bytes,
                        left_cm=2.25, top_cm=25.0, width_cm=5.68, height_cm=3.77) -> Tuple[bytes, bytes, Tuple[str, bytes]]:
    root = ET.fromstring(document_xml)
    rels = ET.fromstring(rels_xml)
    paras = root.findall(".//w:p", NS)
    idx = None
    for i,p in enumerate(paras):
        if get_text(p).strip().lower().startswith("légendes"): idx = i; break
    if idx is None and paras: idx = 0

    nums = []
    for rel in rels.findall(f".//{{{P_REL}}}Relationship"):
        rid = rel.get("Id","")
        if rid.startswith("rId"):
            try: nums.append(int(rid[3:]))
            except: pass
    new_rid = f"rId{(max(nums) if nums else 0)+1}"
    media_name = "media/image_legende.png"
    rel = ET.SubElement(rels, f"{{{P_REL}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    rel.set("Target", media_name)

    drawing = build_anchored_image(new_rid, width_cm, height_cm, left_cm, top_cm, "Legende")
    (ET.SubElement(paras[idx], f"{{{W}}}r") if idx is not None else ET.SubElement(ET.SubElement(root, f"{{{W}}}p"), f"{{{W}}}r")).append(drawing)

    return (ET.tostring(root, encoding="utf-8", xml_declaration=True),
            ET.tostring(rels, encoding="utf-8", xml_declaration=True),
            (f"word/{media_name}", image_bytes))

# ---------- Icon reposition ----------
def reposition_small_icon(root, left_cm=15.3, top_cm=11.0):
    for anchor in root.findall(".//wp:anchor", NS):
        extent = anchor.find("wp:extent", NS)
        if extent is None: continue
        try: cx = int(extent.get("cx","0")); cy = int(extent.get("cy","0"))
        except: continue
        if anchor.find(".//pic:pic", NS) is not None and cx <= cm_to_emu(3.0) and cy <= cm_to_emu(3.0):
            posH = anchor.find("wp:positionH", NS) or ET.SubElement(anchor, f"{{{WP}}}positionH")
            for ch in list(posH): posH.remove(ch)
            posH.set("relativeFrom","page"); ET.SubElement(posH, f"{{{WP}}}posOffset").text = str(cm_to_emu(left_cm))
            posV = anchor.find("wp:positionV", NS) or ET.SubElement(anchor, f"{{{WP}}}positionV")
            for ch in list(posV): posV.remove(ch)
            posV.set("relativeFrom","page"); ET.SubElement(posV, f"{{{WP}}}posOffset").text = str(cm_to_emu(top_cm))
            break

# ---------- Footer normalization (10pt) ----------
def set_dml_text_size(root, pt: float):
    val = str(int(round(pt*100)))
    for r in root.findall(".//a:r", NS):
        rPr = r.find("a:rPr", NS) or ET.SubElement(r, f"{{{A}}}rPr")
        rPr.set("sz", val)

def force_footer_size_10(root):
    for r in root.findall(".//w:r", NS): set_run_props(r, size=10)
    set_dml_text_size(root, 10.0)

# ---------- DML/WPS text sizing on cover (shapes) ----------
def tune_cover_dml_textsizes(root):
    last_was_fiche = False
    for holder in root.findall(".//wp:anchor", NS) + root.findall(".//wp:inline", NS):
        tx = holder.find(".//a:txBody", NS)
        if tx is None: continue
        text = "".join(t.text or "" for t in tx.findall(".//a:t", NS)).strip().lower()
        if not text: continue
        if ("universite" in text) or ("université" in text):
            set_dml_text_size_in_txbody(tx, 10.0); last_was_fiche = False
        elif "fiche de cours" in text:
            set_dml_text_size_in_txbody(tx, 22.0); last_was_fiche = True
        elif last_was_fiche:
            set_dml_text_size_in_txbody(tx, 20.0); last_was_fiche = False
        elif ("introduction" in text) and ("biologie" in text):
            set_dml_text_size_in_txbody(tx, 20.0); last_was_fiche = False
        # also strip ACTUALISATION if it appears inside a shape
        if "actualisation" in text:
            for t in tx.findall(".//a:t", NS):
                if t.text: t.text = re.sub(r"(?i)actualisation","",t.text)

def tune_cover_wps_textsizes(root):
    last_was_fiche = False
    for holder in root.findall(".//wp:anchor", NS) + root.findall(".//wp:inline", NS):
        if not looks_like_cover_shape(holder): continue
        txbx = holder.find(".//wps:txbx/w:txbxContent", NS)
        if txbx is None: continue
        full = "".join(t.text or "" for t in txbx.findall(".//w:t", NS)).strip().lower()
        if not full: continue
        if "universite" in full or "université" in full:
            for r in txbx.findall(".//w:r", NS): set_run_props(r, size=10)
            last_was_fiche = False
        elif "fiche de cours" in full:
            for r in txbx.findall(".//w:r", NS): set_run_props(r, size=22)
            last_was_fiche = True
        elif last_was_fiche:
            for r in txbx.findall(".//w:r", NS): set_run_props(r, size=20)
            last_was_fiche = False
        elif "introduction" in full and "biologie" in full:
            for r in txbx.findall(".//w:r", NS): set_run_props(r, size=20)
            last_was_fiche = False
        # remove ACTUALISATION in WPS textboxes too
        if "actualisation" in full:
            for t in txbx.findall(".//w:t", NS):
                if t.text: t.text = re.sub(r"(?i)actualisation","",t.text)

# ---------- Pipeline ----------
def process_bytes(docx_bytes: bytes,
                  legend_bytes: bytes = None,
                  icon_left=15.3, icon_top=11.0,
                  legend_left=2.25, legend_top=25.0,
                  legend_w=5.68, legend_h=3.77) -> bytes:

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        parts: Dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}
    theme_colors = extract_theme_colors(parts)

    for name, data in list(parts.items()):
        if not name.endswith(".xml"): continue
        try: root = ET.fromstring(data)
        except ET.ParseError: continue

        replace_years(root)
        force_calibri(root)
        red_to_black(root)

        if name == "word/document.xml":
            cover_sizes_cleanup(root)           # w:p
            tune_cover_dml_textsizes(root)      # shapes DML
            tune_cover_wps_textsizes(root)      # shapes WPS
            tables_and_numbering(root)
            reposition_small_icon(root, icon_left, icon_top)
            remove_large_grey_rectangles(root, theme_colors)

        if name.startswith("word/footer"):
            force_footer_size_10(root)

        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # legend image (optional)
    if legend_bytes and "word/document.xml" in parts and "word/_rels/document.xml.rels" in parts:
        parts["word/document.xml"] = remove_legend_text(parts["word/document.xml"])
        new_doc, new_rels, media = insert_legend_image(
            parts["word/document.xml"], parts["word/_rels/document.xml.rels"], legend_bytes,
            left_cm=legend_left, top_cm=legend_top, width_cm=legend_w, height_cm=legend_h
        )
        parts["word/document.xml"] = new_doc
        parts["word/_rels/document.xml.rels"] = new_rels
        parts[media[0]] = media[1]

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, d in parts.items(): zout.writestr(n, d)
    return out_buf.getvalue()

# ---------- UI ----------
st.set_page_config(page_title="Diploma Santé – Phase I", page_icon="🧠", layout="centered")
st.title("🧠 Programme ultime – Phase I")
st.caption("Modifie le .docx en gardant 100% du design + règles demandées.")

with st.sidebar:
    st.subheader("Paramètres (cm)")
    icon_left  = st.number_input("Icône écriture — gauche", value=15.3, step=0.1)
    icon_top   = st.number_input("Icône écriture — haut",   value=11.0, step=0.1)
    legend_left= st.number_input("Image Légendes — gauche", value=2.25, step=0.1)
    legend_top = st.number_input("Image Légendes — haut",   value=25.0, step=0.1)
    legend_w   = st.number_input("Image Légendes — largeur",value=5.68, step=0.01)
    legend_h   = st.number_input("Image Légendes — hauteur",value=3.77, step=0.01)

st.markdown("**1) Glisse/dépose un ou plusieurs fichiers `.docx`**")
files = st.file_uploader("DOCX à traiter", type=["docx"], accept_multiple_files=True)
st.markdown("**2) (Optionnel) Ajoute l’image de la Légende (PNG/JPG)**")
legend_file = st.file_uploader("Image Légendes", type=["png","jpg","jpeg","webp"], accept_multiple_files=False)

if st.button("⚙️ Lancer le traitement", type="primary", disabled=not files):
    if not files:
        st.warning("Ajoute au moins un .docx")
    else:
        legend_bytes = legend_file.read() if legend_file else None
        for up in files:
            try:
                out_bytes = process_bytes(
                    up.read(),
                    legend_bytes=legend_bytes,
                    icon_left=icon_left, icon_top=icon_top,
                    legend_left=legend_left, legend_top=legend_top,
                    legend_w=legend_w, legend_h=legend_h
                )
                out_name = up.name.replace(".docx","") + "_PHASE1_ULTIME.docx"
                st.success(f"✅ Terminé : {up.name}")
                st.download_button("⬇️ Télécharger " + out_name, data=out_bytes,
                                   file_name=out_name,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"❌ Échec pour {up.name} : {e}")

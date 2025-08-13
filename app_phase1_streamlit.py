import io, zipfile, re, os, xml.etree.ElementTree as ET
from typing import Dict, Tuple, List
import streamlit as st

# ------------- Namespaces & utils -------------
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC= "http://schemas.openxmlformats.org/drawingml/2006/picture"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"w":W, "wp":WP, "a":A, "pic":PIC, "r":R}
for k,v in NS.items():
    ET.register_namespace(k, v)

TARGETS = ["2024-2025","2024 - 2025","2024\u00A0-\u00A02025"]
REPL    = "2025 - 2026"
ROMAN_RE = re.compile(r"^\s*[IVXLC]+\s*[.)]?\s+.+", re.IGNORECASE)

def cm_to_emu(cm: float) -> int:
    # 1 cm ≈ 360000 EMU
    return int(round(cm * 360000))

def get_text(p) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS))

def set_run_props(run, size=None, bold=None, italic=None, color=None, calibri=False):
    rPr = run.find("w:rPr", NS) or ET.SubElement(run, f"{{{W}}}rPr")
    if calibri:
        rFonts = rPr.find("w:rFonts", NS) or ET.SubElement(rPr, f"{{{W}}}rFonts")
        rFonts.set(f"{{{W}}}ascii", "Calibri")
        rFonts.set(f"{{{W}}}hAnsi", "Calibri")
        rFonts.set(f"{{{W}}}cs", "Calibri")
    if size is not None:
        v = str(int(round(size*2)))
        sz = rPr.find("w:sz", NS)  or ET.SubElement(rPr, f"{{{W}}}sz");   sz.set(f"{{{W}}}val", v)
        szcs= rPr.find("w:szCs", NS) or ET.SubElement(rPr, f"{{{W}}}szCs"); szcs.set(f"{{{W}}}val", v)
    if bold is not None:
        b = rPr.find("w:b", NS)
        if bold:
            b = b or ET.SubElement(rPr, f"{{{W}}}b"); b.set(f"{{{W}}}val", "1")
        else:
            if b is not None: rPr.remove(b)
    if italic is not None:
        i = rPr.find("w:i", NS)
        if italic:
            i = i or ET.SubElement(rPr, f"{{{W}}}i"); i.set(f"{{{W}}}val", "1")
        else:
            if i is not None: rPr.remove(i)
    if color is not None:
        c = rPr.find("w:color", NS) or ET.SubElement(rPr, f"{{{W}}}color")
        c.set(f"{{{W}}}val", color)

def redistribute(nodes, new):
    lens = [len(n.text or "") for n in nodes]; pos = 0
    for i,n in enumerate(nodes):
        if i < len(nodes)-1:
            n.text = new[pos:pos+lens[i]]; pos += lens[i]
        else:
            n.text = new[pos:]

def replace_years(root):
    # w:t
    for p in root.findall(".//w:p", NS):
        wts = p.findall(".//w:t", NS)
        if not wts: continue
        txt = "".join(t.text or "" for t in wts); new = txt
        for tgt in TARGETS: new = new.replace(tgt, REPL)
        if new != txt: redistribute(wts, new)
    # a:t textboxes
    ats = root.findall(".//a:txBody//a:t", NS)
    if ats:
        txt = "".join(t.text or "" for t in ats); new = txt
        for tgt in TARGETS: new = new.replace(tgt, REPL)
        if new != txt: redistribute(ats, new)

def force_calibri(root):
    for r in root.findall(".//w:r", NS):
        set_run_props(r, calibri=True)

def red_to_black(root):
    for r in root.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS)
        if rPr is None: continue
        c = rPr.find("w:color", NS)
        if c is not None and (c.get(f"{{{W}}}val","").upper() == "FF0000"):
            c.set(f"{{{W}}}val","000000")

def cover_sizes_cleanup(root):
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]
    def set_size(p, pt):
        for r in p.findall(".//w:r", NS): set_run_props(r, size=pt)
    for i,txt in enumerate(texts):
        low = txt.lower()
        if "fiche de cours" in low:
            set_size(paras[i], 22)
            j=i-1
            while j>=0 and not texts[j].strip(): j-=1
            if j>=0: set_size(paras[j], 18)   # matière
            k=i+1
            while k < len(texts) and not texts[k].strip(): k+=1
            if k < len(texts): set_size(paras[k], 20)  # titre du cours
        if "université" in low and (any(x.replace("\u00A0"," ") in txt.replace("\u00A0"," ") for x in TARGETS+[REPL])):
            set_size(paras[i], 10)
        if txt.strip().upper() == "ACTUALISATION":
            for t in paras[i].findall(".//w:t", NS):
                if t.text: t.text = re.sub(r"(?i)ACTUALISATION", "", t.text)
        if txt.strip().upper() == "PLAN":
            set_size(paras[i], 11)

def tables_and_numbering(root):
    # tableaux
    for tbl in root.findall(".//w:tbl", NS):
        rows = tbl.findall(".//w:tr", NS)
        if not rows: continue
        # 1ère ligne = titre
        for p in rows[0].findall(".//w:p", NS):
            for r in p.findall(".//w:r", NS): set_run_props(r, size=12, bold=True)
        # autres lignes = corps 9
        for tr in rows[1:]:
            for p in tr.findall(".//w:p", NS):
                for r in p.findall(".//w:r", NS): set_run_props(r, size=9)
    # numérotation romaine
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]
    for i,p in enumerate(paras):
        if ROMAN_RE.match(texts[i] or ""):
            for r in p.findall(".//w:r", NS): set_run_props(r, size=10, bold=True, italic=True, color="FFFFFF")
            j=i-1
            while j>=0 and not texts[j].strip(): j-=1
            if j>=0:
                for r in paras[j].findall(".//w:r", NS): set_run_props(r, size=12, bold=True)

def reposition_small_icon(root, left_cm=15.5, top_cm=11.0):
    # déplace la 1ère petite image (<=3cm) : “icône écriture”
    for anchor in root.findall(".//wp:anchor", NS):
        extent = anchor.find("wp:extent", NS)
        if extent is None: continue
        cx = int(extent.get("cx","0")); cy = int(extent.get("cy","0"))
        if anchor.find(".//pic:pic", NS) is not None and cx <= cm_to_emu(3) and cy <= cm_to_emu(3):
            posH = anchor.find("wp:positionH", NS) or ET.SubElement(anchor, f"{{{WP}}}positionH", {"relativeFrom":"page"})
            posH.set("relativeFrom","page"); [posH.remove(ch) for ch in list(posH)]
            ET.SubElement(posH, f"{{{WP}}}posOffset").text = str(cm_to_emu(left_cm))
            posV = anchor.find("wp:positionV", NS) or ET.SubElement(anchor, f"{{{WP}}}positionV", {"relativeFrom":"page"})
            posV.set("relativeFrom","page"); [posV.remove(ch) for ch in list(posV)]
            ET.SubElement(posV, f"{{{WP}}}posOffset").text = str(cm_to_emu(top_cm))
            break

def remove_large_grey_rectangles(root):
    GREYS = {"D9D9D9","E7E7E7","EEEEEE","F2F2F2","EDEDED","EFEFEF","DDDDDD","CCCCCC","F0F0F0"}
    for anchor in root.findall(".//wp:anchor", NS):
        extent = anchor.find("wp:extent", NS)
        if extent is None: continue
        cx = int(extent.get("cx","0")); cy = int(extent.get("cy","0"))
        vals = [el.get("val","").upper() for el in anchor.findall(".//a:srgbClr", NS)]
        if cx >= cm_to_emu(6) and cy >= cm_to_emu(6) and any(v in GREYS for v in vals):
            anchor.clear()  # enlève totalement le rectangle

def build_anchored_image(rId, width_cm, height_cm, left_cm, top_cm, name="Legende"):
    cx = cm_to_emu(width_cm); cy = cm_to_emu(height_cm)
    xoff = cm_to_emu(left_cm); yoff = cm_to_emu(top_cm)
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
    ET.SubElement(anchor, f"{{{WP}}}extent", {"cx":str(cx), "cy":str(cy)})
    ET.SubElement(anchor, f"{{{WP}}}effectExtent", {"l":"0","t":"0","r":"0","b":"0"})
    ET.SubElement(anchor, f"{{{WP}}}wrapNone")
    ET.SubElement(anchor, f"{{{WP}}}docPr", {"id":"10","name":name})
    ET.SubElement(anchor, f"{{{WP}}}cNvGraphicFramePr")   # corrigé
    graphic = ET.SubElement(anchor, f"{{{A}}}graphic")
    gData   = ET.SubElement(graphic, f"{{{A}}}graphicData", {"uri":"http://schemas.openxmlformats.org/drawingml/2006/picture"})
    pic     = ET.SubElement(gData, f"{{{PIC}}}pic")
    nvPicPr = ET.SubElement(pic, f"{{{PIC}}}nvPicPr")
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPr", {"id":"0","name":name+".img"})
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPicPr")
    blipFill= ET.SubElement(pic, f"{{{PIC}}}blipFill")
    ET.SubElement(blipFill, f"{{{A}}}blip", {f"{{{R}}}embed": rId})  # rId injecté directement
    stretch = ET.SubElement(blipFill, f"{{{A}}}stretch"); ET.SubElement(stretch, f"{{{A}}}fillRect")
    spPr    = ET.SubElement(pic, f"{{{PIC}}}spPr")
    xfrm    = ET.SubElement(spPr, f"{{{A}}}xfrm")
    ET.SubElement(xfrm, f"{{{A}}}off", {"x":"0","y":"0"})
    ET.SubElement(xfrm, f"{{{A}}}ext", {"cx":str(cx), "cy":str(cy)})
    prst    = ET.SubElement(spPr, f"{{{A}}}prstGeom", {"prst":"rect"}); ET.SubElement(prst, f"{{{A}}}avLst")
    return drawing

def remove_legend_text(document_xml: bytes) -> bytes:
    root = ET.fromstring(document_xml)
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip().lower() == "légendes":
            for t in p.findall(".//w:t", NS):
                t.text = ""
    LINES = {"Notion nouvelle cette année","Notion hors programme","Notion déjà tombée au concours","Astuces et méthodes"}
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip() in LINES:
            for t in p.findall(".//w:t", NS): t.text = ""
            for d in p.findall(".//w:drawing", NS): d.clear()
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def insert_legend_image(document_xml: bytes, rels_xml: bytes, image_bytes: bytes,
                        left_cm=2.25, top_cm=25.0, width_cm=4.91, height_cm=3.66) -> Tuple[bytes, bytes, Tuple[str, bytes]]:
    root = ET.fromstring(document_xml)
    rels = ET.fromstring(rels_xml)

    paras = root.findall(".//w:p", NS)
    idx = None
    for i,p in enumerate(paras):
        if get_text(p).strip().lower().startswith("légendes"):
            idx = i; break
    if idx is None and paras: idx = 0

    # new rel id
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
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    rel.set("Target", media_name)

    # build drawing using the new rId (no placeholder)
    drawing = build_anchored_image(new_rid, width_cm, height_cm, left_cm, top_cm, "Legende")

    if idx is not None:
        ET.SubElement(paras[idx], f"{{{W}}}r").append(drawing)
    else:
        # no paragraph? append at end
        newp = ET.SubElement(root, f"{{{W}}}p")
        ET.SubElement(newp, f"{{{W}}}r").append(drawing)

    return (
        ET.tostring(root, encoding="utf-8", xml_declaration=True),
        ET.tostring(rels, encoding="utf-8", xml_declaration=True),
        (f"word/{media_name}", image_bytes)
    )

def process_bytes(docx_bytes: bytes,
                  legend_bytes: bytes = None,
                  icon_left=15.5, icon_top=11.0,
                  legend_left=2.25, legend_top=25.0,
                  legend_w=4.91, legend_h=3.66) -> bytes:
    # unzip
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        parts: Dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}

    # pass 1: iterate over all XML parts
    for name, data in list(parts.items()):
        if not name.endswith(".xml"): continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        replace_years(root)
        force_calibri(root)
        red_to_black(root)
        if name == "word/document.xml":
            cover_sizes_cleanup(root)
            tables_and_numbering(root)
            reposition_small_icon(root, icon_left, icon_top)
            remove_large_grey_rectangles(root)
        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # legend: remove textual block + (re)insert image
    if legend_bytes and "word/document.xml" in parts and "word/_rels/document.xml.rels" in parts:
        parts["word/document.xml"] = remove_legend_text(parts["word/document.xml"])
        new_doc, new_rels, media = insert_legend_image(
            parts["word/document.xml"], parts["word/_rels/document.xml.rels"], legend_bytes,
            left_cm=legend_left, top_cm=legend_top, width_cm=legend_w, height_cm=legend_h
        )
        parts["word/document.xml"] = new_doc
        parts["word/_rels/document.xml.rels"] = new_rels
        parts[media[0]] = media[1]

    # rezip
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n, d in parts.items():
            zout.writestr(n, d)
    return out_buf.getvalue()

# -------------------- UI --------------------
st.set_page_config(page_title="Diploma Santé – Phase I", page_icon="🧠", layout="centered")
st.title("🧠 Programme ultime – Phase I")
st.caption("Transforme tes .docx en préservant 100% du design et en appliquant toutes les règles demandées.")

with st.sidebar:
    st.subheader("Paramètres (cm)")
    icon_left  = st.number_input("Icône écriture — gauche", value=15.5, step=0.1)
    icon_top   = st.number_input("Icône écriture — haut",   value=11.0, step=0.1)
    legend_left= st.number_input("Image Légendes — gauche", value=2.25, step=0.1)
    legend_top = st.number_input("Image Légendes — haut",   value=25.0, step=0.1)
    legend_w   = st.number_input("Image Légendes — largeur",value=4.91, step=0.01)
    legend_h   = st.number_input("Image Légendes — hauteur",value=3.66, step=0.01)

st.markdown("**1) Glisse/dépose un ou plusieurs fichiers `.docx`**")
files = st.file_uploader("DOCX à traiter", type=["docx"], accept_multiple_files=True)

st.markdown("**2) Ajoute l’image de la Légende (PNG/JPG)**")
legend_file = st.file_uploader("Image Légendes", type=["png","jpg","jpeg","webp"], accept_multiple_files=False)

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
                    icon_left=icon_left, icon_top=icon_top,
                    legend_left=legend_left, legend_top=legend_top,
                    legend_w=legend_w, legend_h=legend_h
                )
                out_name = up.name.replace(".docx", "") + "_PHASE1_ULTIME.docx"
                st.success(f"✅ Terminé : {up.name}")
                st.download_button(
                    "⬇️ Télécharger " + out_name,
                    data=out_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Échec pour {up.name} : {e}")

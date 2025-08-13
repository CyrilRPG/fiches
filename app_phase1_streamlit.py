# -*- coding: utf-8 -*-
import io
import zipfile
import re
import os
import unicodedata
import xml.etree.ElementTree as ET
from typing import Dict, Tuple, List, Optional
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
    rPr = run.find("w:rPr", NS) or ET.SubElement(run, f"{{{W}}}rPr")
    if calibri:
        rFonts = rPr.find("w:rFonts", NS) or ET.SubElement(rPr, f"{{{W}}}rFonts")
        for k in ("ascii", "hAnsi", "cs"):
            rFonts.set(f"{{{W}}}{k}", "Calibri")
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
def replace_years(root):
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

def strip_actualisation_everywhere(root):
    for t in root.findall(".//w:t", NS) + root.findall(".//a:t", NS):
        if t.text:
            t.text = re.sub(r"(?iu)\b(actualisation|nouvelle\s+fiche)\b", "", t.text)

def force_calibri(root):
    for r in root.findall(".//w:r", NS):
        set_run_props(r, calibri=True)

# ── recoloration texte rouge/bleu → noir (inchangé vs version précédente)
def red_to_black(root):
    RED_HEX = {
        "FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"
    }
    BLUE_HEX = {
        "0000FF","0070C0","2E74B5","1F497D","2F5496","4F81BD","5B9BD5","1F4E79","0F4C81","1E90FF","3399FF","3C78D8"
    }
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        r,g,b = rgb
        return (r >= 170 and g <= 110 and b <= 110)
    def looks_blue(rgb: Tuple[int,int,int]) -> bool:
        r,g,b = rgb
        return (b >= 170 and r <= 110 and g <= 140)

    for run in root.findall(".//w:r", NS):
        rPr = run.find("w:rPr", NS)
        if rPr is None:
            continue
        c = rPr.find("w:color", NS)
        if c is None:
            continue
        val = (c.get(f"{{{W}}}val") or "").strip().upper()
        theme = (c.get(f"{{{W}}}themeColor") or "").strip().lower()
        make_black = False
        if theme in {"hyperlink", "followedHyperlink"}:
            make_black = True
        if not make_black and re.fullmatch(r"[0-9A-F]{6}", val or ""):
            if val in RED_HEX or val in BLUE_HEX:
                make_black = True
            else:
                rgb = _hex_to_rgb(val)
                if rgb and (looks_red(rgb) or looks_blue(rgb)):
                    make_black = True
        if make_black:
            c.set(f"{{{W}}}val", "000000")
            for a in ("themeColor", "themeTint", "themeShade"):
                c.attrib.pop(f"{{{W}}}{a}", None)

# ── NOUVEAU : puces rouges → noires dans numbering.xml
def force_red_bullets_black_in_numbering(root):
    """
    Met en noir les couleurs des symboles de liste définies dans word/numbering.xml
    (w:lvl/w:rPr/w:color) si elles sont rouges.
    """
    RED_HEX = {"FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"}
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        if not rgb: return False
        r,g,b = rgb
        return (r >= 170 and g <= 110 and b <= 110)

    for col in root.findall(".//w:lvl//w:rPr/w:color", NS):
        val = (col.get(f"{{{W}}}val") or "").strip().upper()
        make_black = False
        if re.fullmatch(r"[0-9A-F]{6}", val or ""):
            if val in RED_HEX:
                make_black = True
            else:
                rgb = _hex_to_rgb(val)
                if rgb and looks_red(rgb):
                    make_black = True
        # Si la couleur n'est pas fixée (ex: themeColor), on ne touche pas.
        if make_black:
            col.set(f"{{{W}}}val", "000000")
            for a in ("themeColor","themeTint","themeShade"):
                col.attrib.pop(f"{{{W}}}{a}", None)

# ───────────────────────── Helpers couvertures (formes) ────────────
def holder_pos_cm(holder) -> Tuple[float, float]:
    try:
        x = int((holder.find("wp:positionH/wp:posOffset", NS) or ET.Element("x")).text or "0")
        y = int((holder.find("wp:positionV/wp:posOffset", NS) or ET.Element("y")).text or "0")
        return (emu_to_cm(x), emu_to_cm(y))
    except Exception:
        return (0.0, 0.0)

def get_tx_text(holder) -> str:
    tx = holder.find(".//a:txBody", NS)
    if tx is not None:
        return "".join(t.text or "" for t in tx.findall(".//a:t", NS)).strip()
    txbx = holder.find(".//wps:txbx/w:txbxContent", NS)
    if txbx is not None:
        return "".join(t.text or "" for t in txbx.findall(".//w:t", NS)).strip()
    return ""

def set_tx_size(holder, pt: float):
    tx = holder.find(".//a:txBody", NS)
    if tx is not None:
        set_dml_text_size_in_txbody(tx, pt)
    txbx = holder.find(".//wps:txbx/w:txbxContent", NS)
    if txbx is not None:
        for r in txbx.findall(".//w:r", NS):
            set_run_props(r, size=pt)

# ───────────────────────── Mise en forme couverture (paragraphes) ─
def cover_sizes_cleanup(root):
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]
    def set_size(p, pt):
        for r in p.findall(".//w:r", NS):
            set_run_props(r, size=pt)
    last_was_fiche = False
    for i, txt in enumerate(texts):
        low = txt.lower()
        if txt.strip().upper() in ("ACTUALISATION", "NOUVELLE FICHE"):
            for t in paras[i].findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)\b(actualisation|nouvelle\s+fiche)\b", "", t.text)
            continue
        if "fiche de cours" in low:
            set_size(paras[i], 22)
            last_was_fiche = True
            continue
        if last_was_fiche and txt:
            set_size(paras[i], 20)
            last_was_fiche = False
        if "université" in low and any(
            x.replace("\u00A0", " ") in txt.replace("\u00A0", " ") for x in TARGETS + [REPL]
        ):
            set_size(paras[i], 10)

# ───────────────────────── Couverture : formes DML/WPS (spatial) ──
def tune_cover_shapes_spatial(root):
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
    for _, _, h, txt in holders:
        low = txt.lower()
        if ("universite" in low or "université" in low) and any(
            t.replace("\u00A0", " ") in txt.replace("\u00A0", " ") for t in TARGETS + [REPL]
        ):
            set_tx_size(h, 10.0)
    idx_fiche = None
    for i, (_, _, h, txt) in enumerate(holders):
        if "fiche de cours" in txt.lower():
            set_tx_size(h, 22.0)
            idx_fiche = i
            break
    if idx_fiche is not None:
        for j in range(idx_fiche + 1, len(holders)):
            txt_next = holders[j][3].strip()
            if txt_next and "fiche de cours" not in txt_next.lower():
                set_tx_size(holders[j][2], 20.0)
                break
    for _, _, h, _ in holders:
        tx = h.find(".//a:txBody", NS)
        if tx is not None:
            for t in tx.findall(".//a:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)\b(actualisation|nouvelle\s+fiche)\b", "", t.text)
        txbx = h.find(".//wps:txbx/w:txbxContent", NS)
        if txbx is not None:
            for t in txbx.findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(r"(?iu)\b(actualisation|nouvelle\s+fiche)\b", "", t.text)

# ───────────────────────── Tables & numérotations ──────────────────
def tables_and_numbering(root):
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

# ───────────────────────── Couleurs (helpers) ──────────────────────
def _hex_to_rgb(h: str) -> Optional[Tuple[int, int, int]]:
    h = (h or "").strip().lstrip("#").upper()
    if len(h) != 6 or not re.fullmatch(r"[0-9A-F]{6}", h):
        return None
    return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

def _pct(val: Optional[str]) -> float:
    try:
        return max(0.0, min(1.0, int(val)/100000.0))
    except Exception:
        return 1.0

def _apply_lum(base_rgb: Tuple[int,int,int], lumMod: Optional[str], lumOff: Optional[str]) -> Tuple[int,int,int]:
    mod = _pct(lumMod) if lumMod is not None else 1.0
    off = _pct(lumOff) if lumOff is not None else 0.0
    r,g,b = base_rgb
    def f(x):
        return max(0, min(255, int(round(x*mod + 255*off))))
    return (f(r), f(g), f(b))

def _resolve_solid_fill_color(spPr: ET.Element, theme_colors: Dict[str,str]) -> Optional[Tuple[int,int,int]]:
    if spPr is None:
        return None
    solid = None
    for el in spPr.iter():
        if el.tag == f"{{{A}}}solidFill":
            solid = el
            break
    if solid is None:
        return None
    srgb = solid.find("a:srgbClr", NS)
    if srgb is not None and srgb.get("val"):
        rgb = _hex_to_rgb(srgb.get("val"))
        lm = srgb.find("a:lumMod", NS); lo = srgb.find("a:lumOff", NS)
        if rgb and (lm is not None or lo is not None):
            rgb = _apply_lum(rgb, lm.get("val") if lm is not None else None,
                                  lo.get("val") if lo is not None else None)
        return rgb
    scheme = solid.find("a:schemeClr", NS)
    if scheme is not None:
        base_hex = theme_colors.get((scheme.get("val") or "").lower())
        base_rgb = _hex_to_rgb(base_hex) if base_hex else None
        lm = scheme.find("a:lumMod", NS); lo = scheme.find("a:lumOff", NS)
        if base_rgb:
            return _apply_lum(base_rgb, lm.get("val") if lm is not None else None,
                                         lo.get("val") if lo is not None else None)
    sysc = solid.find("a:sysClr", NS)
    if sysc is not None:
        base_hex = sysc.get("lastClr") or sysc.get("val")
        base_rgb = _hex_to_rgb(base_hex)
        lm = sysc.find("a:lumMod", NS); lo = sysc.find("a:lumOff", NS)
        if base_rgb:
            return _apply_lum(base_rgb, lm.get("val") if lm is not None else None,
                                         lo.get("val") if lo is not None else None)
    return None

def _shape_has_text(holder: ET.Element) -> bool:
    txt = get_tx_text(holder)
    return bool(txt.strip())

# ───────────────────────── Thème ───────────────────────────────────
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
        if srgb is not None and srgb.get("val"):
            colors[tag.lower()] = srgb.get("val", "").upper()
        else:
            sysc = el.find("a:sysClr", NS)
            if sysc is not None and sysc.get("lastClr"):
                colors[tag.lower()] = sysc.get("lastClr", "").upper()
    return colors

# ───────────────────────── Suppression rectangle gris ──────────────
def remove_large_grey_rectangles(root: ET.Element, theme_colors: Dict[str, str]):
    parent_map = {child: parent for parent in root.iter() for child in parent}
    for drawing in root.findall(".//w:drawing", NS):
        holder = drawing.find(".//wp:anchor", NS) or drawing.find(".//wp:inline", NS)
        if holder is None:
            continue
        if holder.find(".//pic:pic", NS) is not None:
            continue
        prst = holder.find(".//a:prstGeom", NS)
        if prst is None or prst.get("prst") not in ("rect", "roundRect"):
            continue
        extent = holder.find("wp:extent", NS)
        if extent is None:
            continue
        try:
            cx = int(extent.get("cx", "0")); cy = int(extent.get("cy", "0"))
        except Exception:
            continue
        width_cm  = emu_to_cm(cx); height_cm = emu_to_cm(cy)
        x_el = holder.find(".//wp:positionH/wp:posOffset", NS)
        try:
            x_cm = emu_to_cm(int(x_el.text)) if x_el is not None else 0.0
        except Exception:
            x_cm = 0.0
        if _shape_has_text(holder):
            continue
        spPr = holder.find(".//a:spPr", NS) or holder.find(".//wps:spPr", NS)
        if spPr is None:
            for el in holder.iter():
                if el.tag.endswith("spPr"):
                    spPr = el; break
        rgb = _resolve_solid_fill_color(spPr, theme_colors)
        looks_gray = rgb is not None and \
                     abs(rgb[0]-0xF2) <= 12 and abs(rgb[1]-0xF2) <= 12 and abs(rgb[2]-0xF2) <= 12
        on_right   = x_cm >= 9.0
        big_enough = (width_cm >= 7.0 and height_cm >= 12.0)
        if looks_gray and on_right and big_enough:
            parent = parent_map.get(drawing)
            if parent is not None:
                parent.remove(drawing)
    for pict in root.findall(".//w:pict", NS):
        for tag in ("rect", "roundrect", "shape"):
            for shape in pict.findall(f".//v:{tag}", NS):
                style = (shape.get("style") or "")
                m_w = re.search(r"width:([0-9.]+)cm", style)
                m_h = re.search(r"height:([0-9.]+)cm", style)
                m_l = re.search(r"left:([0-9.]+)cm", style)
                if not (m_w and m_h):
                    continue
                w = float(m_w.group(1)); h = float(m_h.group(1))
                left_cm = float(m_l.group(1)) if m_l else 0.0
                fill_attr = (shape.get("fillcolor") or "").lstrip("#")
                rgb = _hex_to_rgb(fill_attr.upper())
                looks_gray = rgb is not None and \
                             abs(rgb[0]-0xF2) <= 12 and abs(rgb[1]-0xF2) <= 12 and abs(rgb[2]-0xF2) <= 12
                has_txbx = shape.find(".//w:txbxContent", NS) is not None
                if looks_gray and not has_txbx and left_cm >= 9.0 and w >= 7.0 and h >= 12.0:
                    parent = parent_map.get(pict)
                    if parent is not None:
                        parent.remove(pict)

# ───────────────────────── Légende (optionnelle) ───────────────────
def build_anchored_image(rId, width_cm, height_cm, left_cm, top_cm, name="Legende"):
    cx, cy = cm_to_emu(width_cm), cm_to_emu(height_cm)
    xoff, yoff = cm_to_emu(left_cm), cm_to_emu(top_cm)
    drawing = ET.Element(f"{{{W}}}drawing")
    anchor = ET.SubElement(
        drawing,
        f"{{{WP}}}anchor",
        {
            "distT": "0","distB": "0","distL": "0","distR": "0",
            "simplePos": "0","relativeHeight": "0","behindDoc": "0",
            "locked": "0","layoutInCell": "1","allowOverlap": "1",
        },
    )
    ET.SubElement(anchor, f"{{{WP}}}simplePos", {"x": "0", "y": "0"})
    posH = ET.SubElement(anchor, f"{{{WP}}}positionH", {"relativeFrom": "page"})
    ET.SubElement(posH, f"{{{WP}}}posOffset").text = str(xoff)
    posV = ET.SubElement(anchor, f"{{{WP}}}positionV", {"relativeFrom": "page"})
    ET.SubElement(posV, f"{{{WP}}}posOffset").text = str(yoff)
    ET.SubElement(anchor, f"{{{WP}}}extent", {"cx": str(cx), "cy": str(cy)})
    ET.SubElement(anchor, f"{{{WP}}}effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
    ET.SubElement(anchor, f"{{{WP}}}wrapNone")
    ET.SubElement(anchor, f"{{{WP}}}docPr", {"id": "10", "name": name})
    ET.SubElement(anchor, f"{{{WP}}}cNvGraphicFramePr")
    graphic = ET.SubElement(anchor, f"{{{A}}}graphic")
    gData = ET.SubElement(graphic, f"{{{A}}}graphicData", {"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"})
    pic = ET.SubElement(gData, f"{{{PIC}}}pic")
    nvPicPr = ET.SubElement(pic, f"{{{PIC}}}nvPicPr")
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPr", {"id": "0", "name": name + ".img"})
    ET.SubElement(nvPicPr, f"{{{PIC}}}cNvPicPr")
    blipFill = ET.SubElement(pic, f"{{{PIC}}}blipFill")
    ET.SubElement(blipFill, f"{{{A}}}blip", {f"{{{R}}}embed": rId})
    stretch = ET.SubElement(blipFill, f"{{{A}}}stretch")
    ET.SubElement(stretch, f"{{{A}}}fillRect")
    spPr = ET.SubElement(pic, f"{{{PIC}}}spPr")
    xfrm = ET.SubElement(spPr, f"{{{A}}}xfrm")
    ET.SubElement(xfrm, f"{{{A}}}off", {"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{A}}}ext", {"cx": str(cx), "cy": str(cy)})
    prst = ET.SubElement(spPr, f"{{{A}}}prstGeom", {"prst": "rect"})
    ET.SubElement(prst, f"{{{A}}}avLst")
    return drawing

def remove_legend_text(document_xml: bytes) -> bytes:
    root = ET.fromstring(document_xml)
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip().lower() == "légendes":
            for t in p.findall(".//w:t", NS):
                t.text = ""
    lines = {"Notion nouvelle cette année","Notion hors programme","Notion déjà tombée au concours","Astuces et méthodes"}
    for p in root.findall(".//w:p", NS):
        if get_text(p).strip() in lines:
            for t in p.findall(".//w:t", NS):
                t.text = ""
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def insert_legend_image(
    document_xml: bytes, rels_xml: bytes, image_bytes: bytes,
    left_cm=2.3, top_cm=23.8, width_cm=5.68, height_cm=3.77,
) -> Tuple[bytes, bytes, Tuple[str, bytes]]:
    root = ET.fromstring(document_xml)
    rels = ET.fromstring(rels_xml)
    paras = root.findall(".//w:p", NS)
    idx = None
    for i, p in enumerate(paras):
        if get_text(p).strip().lower().startswith("légendes"):
            idx = i; break
    if idx is None and paras: idx = 0
    nums = []
    for rel in rels.findall(f".//{{{P_REL}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try: nums.append(int(rid[3:]))
            except Exception: pass
    new_rid = f"rId{(max(nums) if nums else 0) + 1}"
    media_name = "media/image_legende.png"
    rel = ET.SubElement(rels, f"{{{P_REL}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    rel.set("Target", media_name)
    drawing = build_anchored_image(new_rid, width_cm, height_cm, left_cm, top_cm, "Legende")
    (ET.SubElement(paras[idx], f"{{{W}}}r") if idx is not None else ET.SubElement(ET.SubElement(root, f"{{{W}}}p"), f"{{{W}}}r")).append(drawing)
    return (
        ET.tostring(root, encoding="utf-8", xml_declaration=True),
        ET.tostring(rels, encoding="utf-8", xml_declaration=True),
        (f"word/{media_name}", image_bytes),
    )

# ───────────────────────── Reposition icône écriture ───────────────
def reposition_small_icon(root, left_cm=15.3, top_cm=11.0):
    cand = []
    for anchor in root.findall(".//wp:anchor", NS):
        extent = anchor.find("wp:extent", NS)
        if extent is None:
            continue
        try:
            cx = int(extent.get("cx", "0")); cy = int(extent.get("cy", "0"))
        except Exception:
            continue
        if anchor.find(".//pic:pic", NS) is None:
            continue
        if cx > cm_to_emu(3.0) or cy > cm_to_emu(3.0):
            continue
        x = anchor.findtext("wp:positionH/wp:posOffset", default="0", namespaces=NS)
        y = anchor.findtext("wp:positionV/wp:posOffset", default="0", namespaces=NS)
        try:
            x_cm = emu_to_cm(int(x)); y_cm = emu_to_cm(int(y))
        except Exception:
            x_cm, y_cm = 0.0, 0.0
        cand.append((x_cm, y_cm, anchor))
    if not cand:
        return
    chosen = max(cand, key=lambda t: t[0])
    anchor = chosen[2]
    posH = anchor.find("wp:positionH", NS) or ET.SubElement(anchor, f"{{{WP}}}positionH")
    for ch in list(posH): posH.remove(ch)
    posH.set("relativeFrom", "page")
    ET.SubElement(posH, f"{{{WP}}}posOffset").text = str(cm_to_emu(left_cm))
    posV = anchor.find("wp:positionV", NS) or ET.SubElement(anchor, f"{{{WP}}}positionV")
    for ch in list(posV): posV.remove(ch)
    posV.set("relativeFrom", "page")
    ET.SubElement(posV, f"{{{WP}}}posOffset").text = str(cm_to_emu(top_cm))

# ───────────────────────── Pieds de page 10 pt ─────────────────────
def set_dml_text_size(root, pt: float):
    val = str(int(round(pt * 100)))
    for r in root.findall(".//a:r", NS):
        rPr = r.find("a:rPr", NS) or ET.SubElement(r, f"{{{A}}}rPr")
        rPr.set("sz", val)

def force_footer_size_10(root):
    for r in root.findall(".//w:r", NS):
        if r.find("w:fldChar", NS) is not None or r.find("w:instrText", NS) is not None:
            continue
        set_run_props(r, size=10)
    set_dml_text_size(root, 10.0)

# ───────────────────────── Processing DOCX ─────────────────────────
def process_bytes(
    docx_bytes: bytes,
    legend_bytes: bytes = None,
    icon_left=15.3,
    icon_top=11.0,
    legend_left=2.3,
    legend_top=23.8,
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

        # traitements texte & formats
        replace_years(root)
        strip_actualisation_everywhere(root)
        force_calibri(root)
        red_to_black(root)  # texte rouge/bleu → noir

        if name == "word/document.xml":
            cover_sizes_cleanup(root)
            tune_cover_shapes_spatial(root)
            tables_and_numbering(root)
            reposition_small_icon(root, icon_left, icon_top)
            remove_large_grey_rectangles(root, theme_colors)

        # NOUVEAU : puces rouges → noires (définies dans numbering.xml)
        if name == "word/numbering.xml":
            force_red_bullets_black_in_numbering(root)

        if name.startswith("word/footer"):
            force_footer_size_10(root)

        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # Légende optionnelle (inchangé)
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
    icon_left  = st.number_input("Icône écriture — gauche", value=15.3, step=0.1)
    icon_top   = st.number_input("Icône écriture — haut",   value=11.0, step=0.1)
    legend_left= st.number_input("Image Légendes — gauche", value=2.3, step=0.1)
    legend_top = st.number_input("Image Légendes — haut",   value=23.8, step=0.1)
    legend_w   = st.number_input("Image Légendes — largeur",value=5.68, step=0.01)
    legend_h   = st.number_input("Image Légendes — hauteur",value=3.77, step=0.01)

st.markdown("**1) Glisse/dépose un ou plusieurs fichiers .docx**")
files = st.file_uploader("DOCX à traiter", type=["docx"], accept_multiple_files=True)
st.markdown("**2) (Optionnel) Ajoute l’image de la Légende (PNG/JPG)**")
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

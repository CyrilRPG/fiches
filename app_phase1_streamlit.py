# -*- coding: utf-8 -*-
import io
import zipfile
import re
import os
import unicodedata
import hashlib
from PIL import Image
import xml.etree.ElementTree as ET
from typing import Dict, Tuple, List, Optional, Set
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
# Paires d'années à transformer vers 2025 - 2026 (espaces / tirets flexibles)
YEAR_PAT = re.compile(
    r"(?:(?:2023|2024)"
    r"[\u00A0\u2007\u202F\s]*[\-\u2010\u2011\u2012\u2013\u2014\u2212][\u00A0\u2007\u202F\s]*"
    r"(?:2024|2025))"
)
# Remplacement standardisé
REPL = "2025 - 2026"

# Fragments de chemin SVG caractéristiques pour différencier Annonce / Cible
ANNONCE_SVG_SNIP = b"M1.98047 8.62184C1.88751 8.46071"
CIBLE_SVG_SNIP   = b"M12.2656 2.73438 12.1094 1.32812"

# Caractères décoratifs / symboles potentiellement affichés comme carrés dans Word
ODD_SYMBOLS_PATTERN = re.compile(r"[\uE000-\uF8FF\u25A0-\u25FF\u2600-\u27BF]")

# Empreintes des bitmaps \"carrés\" issus des annonces (calculées sur assets/fichier_traite.docx)
ANNONCE_SQUARE_BITMAP_HASHES = {
    # Très petites images : icônes/carrés probables
    "818440ed634067aa2598299a8e30c8777af318a3",  # word/media/image3.png (size=1829)
    "68191c53e118952cacde7e0525a31bc2dd328270",  # word/media/image4.png (size=6725)
    "79cea50836da9df36257c4aa1388a399dddf12fe",  # word/media/image6.png (size=11372)
}

ROMAN_TITLE_RE = re.compile(r"^\s*[IVXLC]+\s*[.)]?\s+.+", re.IGNORECASE)

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

def _norm_matchable(s: str) -> str:
    s = s.replace("\u00A0", " ")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s

# ───────────────────────── Remplacements texte ─────────────────────
def replace_years(root):
    for p in root.findall(".//w:p", NS):
        wts = p.findall(".//w:t", NS)
        if not wts:
            continue
        txt = "".join(t.text or "" for t in wts)
        new = YEAR_PAT.sub(REPL, txt)
        # Si le motif '2025 - 2026' est suivi de lettres (UN, P, Paris, etc.),
        # on ne garde que '2025 - 2026'.
        new = re.sub(rf"{re.escape(REPL)}\s*[A-Za-zÀ-ÿ]+", REPL, new)
        if new != txt:
            redistribute(wts, new)
    for tx in root.findall(".//a:txBody", NS):
        ats = tx.findall(".//a:t", NS)
        if not ats:
            continue
        txt = "".join(t.text or "" for t in ats)
        new = YEAR_PAT.sub(REPL, txt)
        new = re.sub(rf"{re.escape(REPL)}\s*[A-Za-zÀ-ÿ]+", REPL, new)
        if new != txt:
            redistribute(ats, new)

def strip_actualisation_everywhere(root):
    PAT = re.compile(
        r"(?iu)\b(actualisation|nouvelle\s+fiche|changements?\s+notables?|nouveau\s+cours|aucun\s+changement)\b"
    )
    for t in root.findall(".//w:t", NS) + root.findall(".//a:t", NS):
        if t.text:
            # Supprimer les mentions de type \"ACTUALISATION\", etc.
            new_txt = PAT.sub("", t.text)
            # Supprimer également les symboles décoratifs susceptibles d'apparaître comme des carrés
            new_txt = ODD_SYMBOLS_PATTERN.sub("", new_txt)
            t.text = new_txt

def force_calibri(root):
    for r in root.findall(".//w:r", NS):
        set_run_props(r, calibri=True)

# ───────────────────────── Couleurs ────────────────────────────────
def _hex_to_rgb(h: str) -> Optional[Tuple[int, int, int]]:
    h = (h or "").strip().lstrip("#").upper()
    if len(h) != 6 or not re.fullmatch(r"[0-9A-F]{6}", h):
        return None
    return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

def red_to_black(root):
    RED_HEX = {
        "FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"
    }
    BLUE_HEX = {
        "0000FF","0070C0","2E74B5","1F497D","2F5496","4F81BD","5B9BD5","1F4E79","0F4C81","1E90FF","3399FF","3C78D8"
    }
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        r,g,b = rgb; return (r >= 170 and g <= 110 and b <= 110)
    def looks_blue(rgb: Tuple[int,int,int]) -> bool:
        r,g,b = rgb; return (b >= 170 and r <= 110 and g <= 140)

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

def force_red_bullets_black_in_numbering(root):
    RED_HEX = {"FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"}
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        if not rgb: return False
        r,g,b = rgb; return (r >= 170 and g <= 110 and b <= 110)
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
        if make_black:
            col.set(f"{{{W}}}val", "000000")
            for a in ("themeColor","themeTint","themeShade"):
                col.attrib.pop(f"{{{W}}}{a}", None)

def force_red_bullets_black_in_styles(root):
    CANDIDATES = {"list","bullet","puce","puces","liste"}
    RED_HEX = {"FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"}
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        if not rgb: return False
        r,g,b = rgb; return (r >= 170 and g <= 110 and b <= 110)
    for st in root.findall(".//w:style[@w:type='paragraph']", NS):
        name_el = st.find("w:name", NS)
        style_id = (st.get(f"{{{W}}}styleId") or "").lower()
        style_name = (name_el.get(f"{{{W}}}val") if name_el is not None else "").lower()
        tag = (style_id + " " + style_name)
        if not any(tok in tag for tok in CANDIDATES):
            continue
        col = st.find(".//w:rPr/w:color", NS)
        if col is None:
            continue
        val = (col.get(f"{{{W}}}val") or "").strip().upper()
        make_black = False
        if re.fullmatch(r"[0-9A-F]{6}", val or ""):
            if val in RED_HEX:
                make_black = True
            else:
                rgb = _hex_to_rgb(val)
                if rgb and looks_red(rgb):
                    make_black = True
        if make_black:
            col.set(f"{{{W}}}val", "000000")
            for a in ("themeColor","themeTint","themeShade"):
                col.attrib.pop(f"{{{W}}}{a}", None)

def force_red_bullets_black_in_paragraphs(root):
    RED_HEX = {"FF0000","C00000","CC0000","E60000","ED1C24","F44336","DC143C","B22222","E74C3C","D0021B"}
    def looks_red(rgb: Tuple[int,int,int]) -> bool:
        if not rgb: return False
        r,g,b = rgb; return (r >= 170 and g <= 110 and b <= 110)
    for p in root.findall(".//w:p", NS):
        pPr = p.find("w:pPr", NS)
        if pPr is None or pPr.find("w:numPr", NS) is None:
            continue
        rPr = pPr.find("w:rPr", NS)
        if rPr is None:
            continue
        col = rPr.find("w:color", NS)
        if col is None:
            continue
        val = (col.get(f"{{{W}}}val") or "").strip().upper()
        make_black = False
        if re.fullmatch(r"[0-9A-F]{6}", val or ""):
            if val in RED_HEX:
                make_black = True
            else:
                rgb = _hex_to_rgb(val)
                if rgb and looks_red(rgb):
                    make_black = True
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

# ───────────────────────── Mise en forme couverture ────────────────
def cover_sizes_cleanup(root):
    paras = root.findall(".//w:p", NS)
    texts = [get_text(p).strip() for p in paras]
    def set_size(p, pt):
        for r in p.findall(".//w:r", NS):
            set_run_props(r, size=pt)
    last_was_fiche = False
    for i, txt in enumerate(texts):
        low = txt.lower()
        if txt.strip().upper() in (
            "ACTUALISATION",
            "NOUVELLE FICHE",
            "CHANGEMENTS NOTABLES",
            "NOUVEAU COURS",
            "AUCUN CHANGEMENT",
        ):
            for t in paras[i].findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(
                        r"(?iu)\b(actualisation|nouvelle\s+fiche|changements?\s+notables?|nouveau\s+cours|aucun\s+changement)\b",
                        "",
                        t.text,
                    )
            continue
        if "fiche de cours" in low:
            # Titre "Fiche de cours" en 20 pt
            set_size(paras[i], 20)
            # Bloc précédent non vide = matière, en 18 pt
            for k in range(i - 1, -1, -1):
                if get_text(paras[k]).strip():
                    set_size(paras[k], 18)
                    break
            last_was_fiche = True
            continue
        if last_was_fiche and txt:
            # Bloc juste après "Fiche de cours" = nom du cours, en 22 pt
            set_size(paras[i], 22)
            last_was_fiche = False
        if "université" in low and (YEAR_PAT.search(txt.replace("\u00A0"," ")) or "universite" in low):
            # Bloc université + année, en 10 pt
            set_size(paras[i], 10)
        # Bloc "PLAN I / II ..." en 11 pt
        if txt.strip().upper().startswith("PLAN"):
            set_size(paras[i], 11)

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
    # Bloc université + année en 10 pt et remplacer l'année
    for _, _, h, txt in holders:
        low = txt.lower()
        if ("universite" in low or "université" in low):
            set_tx_size(h, 10.0)
    # Fiche de cours + matière + nom du cours
    idx_fiche = None
    for i, (_, _, h, txt) in enumerate(holders):
        if "fiche de cours" in txt.lower():
            # Titre "Fiche de cours" en 20 pt
            set_tx_size(h, 20.0)
            idx_fiche = i
            # Bloc précédent non vide = matière, en 18 pt
            for k in range(i - 1, -1, -1):
                prev_txt = holders[k][3].strip()
                if prev_txt and "fiche de cours" not in prev_txt.lower():
                    set_tx_size(holders[k][2], 18.0)
                    break
            break
    if idx_fiche is not None:
        # Bloc suivant non vide = nom du cours, en 22 pt
        for j in range(idx_fiche + 1, len(holders)):
            txt_next = holders[j][3].strip()
            if txt_next and "fiche de cours" not in txt_next.lower():
                set_tx_size(holders[j][2], 22.0)
                break
    # Bloc "PLAN ..." en 11 pt
    for _, _, h, txt in holders:
        if txt.strip().upper().startswith("PLAN"):
            set_tx_size(h, 11.0)
    for _, _, h, _ in holders:
        tx = h.find(".//a:txBody", NS)
        if tx is not None:
            for t in tx.findall(".//a:t", NS):
                if t.text:
                    t.text = re.sub(
                        r"(?iu)\b(actualisation|nouvelle\s+fiche|changements?\s+notables?|nouveau\s+cours|aucun\s+changement)\b",
                        "",
                        t.text,
                    )
        txbx = h.find(".//wps:txbx/w:txbxContent", NS)
        if txbx is not None:
            for t in txbx.findall(".//w:t", NS):
                if t.text:
                    t.text = re.sub(
                        r"(?iu)\b(actualisation|nouvelle\s+fiche|changements?\s+notables?|nouveau\s+cours|aucun\s+changement)\b",
                        "",
                        t.text,
                    )

def force_title_fiche_de_cours_22(root):
    """
    Historiquement : forçaient le titre \"Fiche de cours\" à 22 pt.
    Désormais on aligne avec la nouvelle maquette :
      - \"Fiche de cours\" en 20 pt
      - le bloc suivant (nom du cours) en 22 pt
    """
    for p in root.findall(".//w:p", NS):
        if "fiche de cours" in _norm_matchable(get_text(p)):
            for r in p.findall(".//w:r", NS):
                set_run_props(r, size=20)
    for holder in root.findall(".//wp:anchor", NS) + root.findall(".//wp:inline", NS):
        txt = get_tx_text(holder)
        if txt and "fiche de cours" in _norm_matchable(txt):
            set_tx_size(holder, 20.0)

def force_course_name_after_title_20(root):
    paras = root.findall(".//w:p", NS)
    for i, p in enumerate(paras):
        if "fiche de cours" in _norm_matchable(get_text(p)):
            for j in range(i+1, len(paras)):
                if get_text(paras[j]).strip():
                    # Bloc suivant = nom du cours, en 22 pt
                    for r in paras[j].findall(".//w:r", NS):
                        set_run_props(r, size=22)
                    break
            break

# ───────────────────────── Tables & numérotations ──────────────────
def _is_dark_hex(hexv: Optional[str]) -> bool:
    if not hexv: return False
    rgb = _hex_to_rgb(hexv)
    if not rgb: return False
    r,g,b = rgb
    return (r+g+b) < 200 and b >= max(r, g)

_DARK_BLUE_SET = {"002060","1F4E79","0F4C81","1F497D","2F5496","112F4E","203764","23395D"}

def _para_or_cell_has_dark_bg(p: ET.Element, parent_map: Dict[ET.Element, ET.Element]) -> bool:
    shd = p.find("w:pPr/w:shd", NS)
    if shd is not None:
        fill = (shd.get(f"{{{W}}}fill") or "").upper()
        if fill in _DARK_BLUE_SET or _is_dark_hex(fill):
            return True
    node = p
    while node is not None and node.tag != f"{{{W}}}tc":
        node = parent_map.get(node)
    if node is not None:
        shd2 = node.find("w:tcPr/w:shd", NS)
        if shd2 is not None:
            fill2 = (shd2.get(f"{{{W}}}fill") or "").upper()
            if fill2 in _DARK_BLUE_SET or _is_dark_hex(fill2):
                return True
    return False

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

    parent_map = {child: parent for parent in root.iter() for child in parent}
    for p in root.findall(".//w:p", NS):
        txt = get_text(p).strip()
        if not txt:
            continue
        if not ROMAN_TITLE_RE.match(txt):
            continue
        if not _para_or_cell_has_dark_bg(p, parent_map):
            continue
        for r in p.findall(".//w:r", NS):
            set_run_props(r, size=10, bold=True, italic=True, color="FFFFFF")

# ───────────────────────── Helpers couleurs formes ─────────────────
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
        drawing, f"{{{WP}}}anchor",
        {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"0",
         "behindDoc":"0","locked":"0","layoutInCell":"1","allowOverlap":"1"}
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
    lines = {
        "Notion nouvelle cette année",
        "Notion hors programme",
        "Notion déjà tombée au concours",
        "Astuces et méthodes",
    }
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

# ───────────────────────── Suppression mégaphones ──────────────────
def _sha1(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()

def _ahash(b: bytes, size: int = 8) -> Optional[int]:
    """
    Hash perceptuel très simple (average hash) pour comparer les petites icônes
    même si Word les a légèrement recompressées ou redimensionnées.
    Retourne un entier de size*size bits, ou None en cas d'erreur.
    """
    try:
        with Image.open(io.BytesIO(b)) as im:
            im = im.convert("L").resize((size, size), Image.LANCZOS)
            pixels = list(im.getdata())
    except Exception:
        return None
    avg = sum(pixels) / len(pixels)
    bits = 0
    for i, p in enumerate(pixels):
        if p > avg:
            bits |= (1 << i)
    return bits

def _hamming(a: int, b: int) -> int:
    return bin(a ^ b).count("1")

def is_annonce_square_media(name: str, data: bytes) -> bool:
    """
    Détecte si un media bitmap (PNG/EMF/JPEG) correspond à un carré
    issu d'une annonce, en se basant uniquement sur son hash SHA1.
    
    Les hash sont calculés sur les fichiers réels trouvés dans les documents traités
    (ex: image3.png, image4.png, image6.png dans fichier_traite.docx).
    
    Ne fait plus d'heuristique supplémentaire pour éviter les faux positifs.
    """
    lname = name.lower()
    if not lname.startswith("word/media/"):
        return False
    if not (lname.endswith(".png") or lname.endswith(".emf") or lname.endswith(".jpg") or lname.endswith(".jpeg")):
        return False
    h = _sha1(data)
    return h in ANNONCE_SQUARE_BITMAP_HASHES

# ───────────────────────── SVG modèle (annonce) ──────────────────────
def _load_svg_model_bytes() -> Optional[bytes]:
    """
    Charge le SVG d'annonce de référence depuis assets/annonce.svg.
    Retourne None si le fichier est introuvable.
    """
    candidates = ["annonce.svg", "Annonce.svg"]
    possible_paths = []
    try:
        possible_paths.append(os.path.dirname(__file__))
    except Exception:
        pass
    possible_paths.extend([os.getcwd(), "."])

    for base in possible_paths:
        for fname in candidates:
            p1 = os.path.join(base, "assets", fname)
            p2 = os.path.join(base, fname)
            for path in (p1, p2):
                try:
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            return f.read()
                except OSError:
                    continue
    return None

def _extract_svg_paths(svg_bytes: bytes) -> List[str]:
    """
    Extrait tous les attributs 'd' des éléments <path> d'un SVG.
    Retourne une liste de chemins normalisés (espaces supprimés, nombres normalisés).
    """
    paths = []
    try:
        root = ET.fromstring(svg_bytes)
        for path_el in root.iter():
            # Gérer les namespaces
            tag = path_el.tag.split("}", 1)[-1] if "}" in path_el.tag else path_el.tag
            if tag == "path":
                d_attr = path_el.get("d") or ""
                if d_attr:
                    # Normaliser : supprimer espaces multiples, normaliser nombres
                    normalized = re.sub(r"\s+", " ", d_attr.strip())
                    paths.append(normalized)
    except Exception:
        pass
    return paths

def _svg_content_matches(svg_bytes: bytes, model_bytes: bytes) -> bool:
    """
    Compare deux SVG en extrayant et comparant leurs chemins <path d="...">.
    Retourne True si les SVG correspondent (même contenu géométrique).
    """
    if not svg_bytes or not model_bytes:
        return False
    
    svg_paths = _extract_svg_paths(svg_bytes)
    model_paths = _extract_svg_paths(model_bytes)
    
    if not svg_paths or not model_paths:
        # Si aucun chemin trouvé, comparaison par hash SHA1
        return _sha1(svg_bytes) == _sha1(model_bytes)
    
    # Normaliser et trier les chemins pour comparaison
    svg_paths_sorted = sorted(svg_paths)
    model_paths_sorted = sorted(model_paths)
    
    # Si le nombre de chemins diffère beaucoup, pas de match
    if abs(len(svg_paths_sorted) - len(model_paths_sorted)) > max(1, len(model_paths_sorted) * 0.2):
        return False
    
    # Comparer les chemins : au moins 80% doivent correspondre
    matches = 0
    min_len = min(len(svg_paths_sorted), len(model_paths_sorted))
    for i in range(min_len):
        if svg_paths_sorted[i] == model_paths_sorted[i]:
            matches += 1
    
    # Match si au moins 80% des chemins correspondent
    return matches >= min_len * 0.8

def _normalize_svg(svg_bytes: bytes) -> Optional[bytes]:
    """
    Normalise un SVG en extrayant uniquement les éléments géométriques
    essentiels (paths, formes simples) et en ignorant les métadonnées,
    styles et attributs de position sujets à variation.

    L'objectif est d'obtenir une « signature visuelle » stable pour
    comparer annonce.svg / cible.svg avec leurs équivalents Word.
    """
    try:
        root = ET.fromstring(svg_bytes)
    except Exception:
        return None

    def local_tag(el: ET.Element) -> str:
        return el.tag.split("}", 1)[-1] if "}" in el.tag else el.tag

    # Clés géométriques pertinentes par type de forme
    geom_keys = {
        "path": {"d"},
        "polygon": {"points"},
        "polyline": {"points"},
        "circle": {"cx", "cy", "r"},
        "ellipse": {"cx", "cy", "rx", "ry"},
        "rect": {"x", "y", "width", "height", "rx", "ry"},
        "line": {"x1", "y1", "x2", "y2"},
    }

    # Construire une liste de « signatures » de formes
    shapes = []
    for el in root.iter():
        tag = local_tag(el)
        if tag not in geom_keys:
            continue
        keys = geom_keys[tag]
        attrs = []
        for k, v in el.attrib.items():
            lk = k.split("}", 1)[-1] if "}" in k else k
            if lk in keys:
                attrs.append(f"{lk}={v}")
        if not attrs:
            continue
        attrs.sort()
        shapes.append(f"{tag}|" + "|".join(attrs))

    if not shapes:
        # Repli : si on ne trouve pas de formes géométriques classiques,
        # revenir à l'ancienne normalisation simple.
        try:
            norm = ET.tostring(root, encoding="utf-8", xml_declaration=False)
            norm = re.sub(rb"\s+", b" ", norm)
            norm = re.sub(rb">\s+<", b"><", norm)
            return norm.strip()
        except Exception:
            return None

    # Signature finale : liste triée des signatures de formes
    shapes.sort()
    sig = "\n".join(shapes).encode("utf-8")
    return sig

def _load_cible_svg_model() -> Optional[bytes]:
    """
    Charge le SVG modèle Cible.svg depuis assets/.
    """
    possible_paths = []
    try:
        possible_paths.append(os.path.dirname(__file__))
    except:
        pass
    possible_paths.extend([os.getcwd(), "."])
    
    for base in possible_paths:
        for fname in ["cible.svg", "Cible.svg"]:
            p1 = os.path.join(base, "assets", fname)
            p2 = os.path.join(base, fname)
            for path in (p1, p2):
                try:
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            return f.read()
                except OSError:
                    continue
    return None

def _identify_svg_to_remove(parts: Dict[str, bytes]) -> Set[str]:
    """
    Parcourt TOUS les fichiers word/media/*.svg et identifie ceux à supprimer.
    Règle basée sur CIBLE_SVG_SNIP : on garde UNIQUEMENT les SVG contenant ce fragment,
    on supprime TOUS les autres (annonces + autres SVG).
    
    NOTE: Cette logique fonctionnait mais laissait des carrés vides.
    On la garde pour l'instant et on travaillera sur la suppression des carrés ensuite.
    """
    svg_to_remove: Set[str] = set()

    # Parcourir tous les SVG dans word/media/
    for name, data in parts.items():
        lname = name.lower()
        if not lname.startswith("word/"):
            continue
        if "/media/" not in lname:
            continue
        if not lname.endswith(".svg"):
            continue

        # Règle basée sur le fragment caractéristique de Cible.svg
        # Si le SVG contient CIBLE_SVG_SNIP, on le garde (cible)
        # Sinon, on le supprime (annonce ou autre)
        if CIBLE_SVG_SNIP in data:
            # Cible : on la garde, ne pas ajouter à svg_to_remove
            continue
        else:
            # Annonce ou autre SVG : à supprimer
            svg_to_remove.add(name)
    
    return svg_to_remove

def _find_matching_svg_media(parts: Dict[str, bytes], svg_model: bytes) -> Set[str]:
    """
    Retourne l'ensemble des chemins 'word/media/*.svg' correspondant
    au SVG d'annonce à supprimer. On utilise ici un fragment de chemin
    très spécifique (ANNONCE_SVG_SNIP) pour être robuste aux
    changements de formatage XML.
    """
    matches: Set[str] = set()
    if not svg_model:
        return matches

    for name, data in parts.items():
        lname = name.lower()
        if not lname.startswith("word/"):
            continue
        if "/media/" not in lname:
            continue
        if not lname.endswith(".svg"):
            continue

        # Détection simple par fragment de chemin : si le SVG contient
        # la trace caractéristique d'Annonce, on le marque pour suppression.
        if ANNONCE_SVG_SNIP in data:
            matches.add(name)
    return matches

def _load_default_megaphone_hashes() -> Tuple[Set[str], Set[int]]:
    """
    Charge les icônes 'Annonce' fournies dans le dossier assets comme
    mégaphones à supprimer, sans toucher aux autres icônes de la fiche cible.
    """
    sha_hashes: Set[str] = set()
    ahashes: Set[int] = set()
    # Icônes d'annonce fournies : PNG et SVG
    candidates = ["Annonce1.png", "Annonce2.png", "Annonce.svg"]
    
    # Essayer plusieurs chemins possibles pour trouver assets
    possible_paths = []
    try:
        possible_paths.append(os.path.dirname(__file__))
    except:
        pass
    possible_paths.extend([os.getcwd(), "."])
    
    for base_dir in possible_paths:
        for filename in candidates:
            # Essayer avec assets/ devant
            path1 = os.path.join(base_dir, "assets", filename)
            # Essayer directement dans le dossier
            path2 = os.path.join(base_dir, filename)
            for path in [path1, path2]:
                try:
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            data = f.read()
                            sha_hashes.add(_sha1(data))
                            # Pour les SVG, _ahash renvoie souvent None,
                            # mais pour les PNG on obtient bien un hash perceptuel.
                            ah = _ahash(data)
                            if ah is not None:
                                ahashes.add(ah)
                except OSError:
                    pass
    return sha_hashes, ahashes

def _load_protected_icon_hashes() -> Tuple[Set[str], Set[int]]:
    """
    Charge les icônes qui ne doivent JAMAIS être supprimées (ex: Cible.png).
    """
    sha_hashes: Set[str] = set()
    ahashes: Set[int] = set()
    # Icônes de cible à protéger : PNG et SVG
    candidates = ["Cible.png", "Cible.svg"]
    
    # Essayer plusieurs chemins possibles pour trouver assets
    possible_paths = []
    try:
        possible_paths.append(os.path.dirname(__file__))
    except:
        pass
    possible_paths.extend([os.getcwd(), "."])
    
    for base_dir in possible_paths:
        for filename in candidates:
            # Essayer avec assets/ devant
            path1 = os.path.join(base_dir, "assets", filename)
            # Essayer directement dans le dossier
            path2 = os.path.join(base_dir, filename)
            for path in [path1, path2]:
                try:
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            data = f.read()
                            sha_hashes.add(_sha1(data))
                            ah = _ahash(data)
                            if ah is not None:
                                ahashes.add(ah)
                except OSError:
                    pass
    return sha_hashes, ahashes

def _rels_name_for(part_name: str) -> str:
    d = os.path.dirname(part_name)
    b = os.path.basename(part_name)
    return os.path.join(d, "_rels", b + ".rels")

def _resolve_target_path(base_part: str, target: str) -> str:
    base_dir = os.path.dirname(base_part)
    norm = os.path.normpath(os.path.join(base_dir, target))
    return norm.replace("\\", "/")

def remove_media_references(parts: Dict[str, bytes], media_paths_to_remove: Set[str]) -> None:
    """
    Supprime toutes les références aux media identifiés (SVG annonces + bitmaps carrés).
    Fonction unifiée qui gère à la fois les SVG et les bitmaps.
    
    Étapes :
    1. Construit une map globale rId -> media_path en lisant tous les .rels
    2. Identifie tous les rIds pointant vers les media à supprimer
    3. Supprime tous les blips/drawings/pict dans toutes les parties XML
    4. Supprime toutes les relations dans les .rels
    5. Supprime physiquement les fichiers media de parts
    """
    if not media_paths_to_remove:
        return
    
    # Étape 1 : Construire un map complet de TOUS les rIds -> chemins media
    all_rid_to_media: Dict[str, str] = {}  # rId -> chemin media résolu
    
    for name in list(parts.keys()):
        if not name.endswith(".rels") or "/_rels/" not in name:
            continue
        try:
            rels_root = ET.fromstring(parts[name])
            base_name = name.replace("/_rels/", "/").replace(".rels", "")
            for rel in rels_root.findall(f".//{{{P_REL}}}Relationship"):
                rid = rel.get("Id") or ""
                tgt = rel.get("Target") or ""
                if rid and tgt:
                    media_path = _resolve_target_path(base_name, tgt)
                    all_rid_to_media[rid] = media_path
        except Exception:
            continue
    
    # Étape 2 : Identifier TOUS les rIds qui pointent vers les media à supprimer
    rids_to_remove_all = {rid for rid, path in all_rid_to_media.items() if path in media_paths_to_remove}
    
    # Étape 3 : Supprimer TOUS les blips/drawings avec ces rIds dans TOUTES les parties XML
    for name, data in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except Exception:
            continue
        
        parent_map = {child: parent for parent in root.iter() for child in parent}
        changed = False
        
        # Supprimer TOUS les blips avec les rIds identifiés
        # ET aussi les blips orphelins (qui pointent vers des rIds qui n'existent plus)
        all_valid_rids = set(all_rid_to_media.keys())
        
        for blip in list(root.findall(".//a:blip", NS)):
            rid = blip.get(f"{{{R}}}embed")
            should_remove = False
            
            if rid in rids_to_remove_all:
                # rId identifié pour suppression
                should_remove = True
            elif rid and rid not in all_valid_rids:
                # rId orphelin (la relation a été supprimée mais le blip reste)
                # Vérifier si c'était un SVG (probablement un SVG annonce supprimé)
                # On supprime par précaution si le rId n'existe plus
                should_remove = True
            
            if should_remove:
                # Remonter jusqu'à w:drawing et le supprimer
                node = blip
                drawing = None
                while node is not None:
                    if node.tag == f"{{{W}}}drawing":
                        drawing = node
                        break
                    node = parent_map.get(node)
                
                if drawing is not None:
                    parent = parent_map.get(drawing)
                    if parent is not None:
                        parent.remove(drawing)
                        changed = True
        
        # Supprimer TOUS les imagedata VML avec les rIds identifiés
        for imagedata in list(root.findall(f".//v:imagedata", NS)):
            rid = imagedata.get(f"{{{R}}}id")
            if rid in rids_to_remove_all:
                node = imagedata
                pict = None
                while node is not None:
                    if node.tag == f"{{{W}}}pict":
                        pict = node
                        break
                    node = parent_map.get(node)
                
                if pict is not None:
                    parent = parent_map.get(pict)
                    if parent is not None:
                        parent.remove(pict)
                        changed = True
        
        if changed:
            parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    
    # Étape 4 : Supprimer TOUTES les relations dans TOUS les .rels
    for name in list(parts.keys()):
        if not name.endswith(".rels") or "/_rels/" not in name:
            continue
        try:
            rels_root = ET.fromstring(parts[name])
            changed = False
            for rel in list(rels_root.findall(f".//{{{P_REL}}}Relationship")):
                rid = rel.get("Id") or ""
                if rid in rids_to_remove_all:
                    rels_root.remove(rel)
                    changed = True
            if changed:
                parts[name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
        except Exception:
            continue
    
    # Étape 5 : Supprimer physiquement les fichiers media (SVG + bitmaps)
    for media_path in media_paths_to_remove:
        parts.pop(media_path, None)

def _remove_svg_references(parts: Dict[str, bytes], svg_paths_to_remove: Set[str]) -> None:
    """
    Supprime toutes les références aux SVG identifiés dans toutes les parties du document.
    - Supprime les <a:blip r:embed="rId"> et leurs <w:drawing> parents
    - Supprime les <v:imagedata r:id="rId"> et leurs <w:pict> parents
    - Supprime les relations dans les .rels
    - Supprime physiquement les fichiers SVG de parts
    """
    if not svg_paths_to_remove:
        return
    
    # Construire un map inversé : chemin media -> rId pour toutes les parties
    media_to_rids: Dict[str, Set[str]] = {}  # part_name -> set of rIds
    
    # Parcourir tous les fichiers .rels pour trouver les références
    for name in list(parts.keys()):
        if not name.endswith(".rels"):
            continue
        if "/_rels/" not in name:
            continue
        
        try:
            rels_root = ET.fromstring(parts[name])
        except ET.ParseError:
            continue
        
        # Déterminer le nom de la partie parente
        base_name = name.replace("/_rels/", "/").replace(".rels", "")
        
        for rel in rels_root.findall(f".//{{{P_REL}}}Relationship"):
            rid = rel.get("Id") or ""
            tgt = rel.get("Target") or ""
            if not rid or not tgt:
                continue
            
            # Résoudre le chemin complet du média
            media_path = _resolve_target_path(base_name, tgt)
            
            # Si ce média est un SVG à supprimer, noter le rId
            if media_path in svg_paths_to_remove:
                if base_name not in media_to_rids:
                    media_to_rids[base_name] = set()
                media_to_rids[base_name].add(rid)
    
    # Maintenant, supprimer toutes les références dans TOUTES les parties XML
    # On parcourt toutes les parties XML, pas seulement celles dans media_to_rids
    for name, data in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        
        # Obtenir les rIds à supprimer pour cette partie (peut être vide)
        rids_to_remove = media_to_rids.get(name, set())
        
        # Construire aussi un set de tous les rIds à supprimer (toutes parties confondues)
        # pour être sûr de tout attraper
        all_rids_to_remove = set()
        for rids in media_to_rids.values():
            all_rids_to_remove.update(rids)
        
        parent_map = {child: parent for parent in root.iter() for child in parent}
        changed = False
        
        # Supprimer les <a:blip r:embed="rId"> et leurs <w:drawing> parents
        for blip in root.findall(".//a:blip", NS):
            rid = blip.get(f"{{{R}}}embed")
            if rid and (rid in rids_to_remove or rid in all_rids_to_remove):
                # Remonter jusqu'à w:drawing
                node = blip
                drawing = None
                while node is not None:
                    if node.tag == f"{{{W}}}drawing":
                        drawing = node
                        break
                    node = parent_map.get(node)
                
                if drawing is not None:
                    # Supprimer le drawing, ou remonter jusqu'au run parent si nécessaire
                    parent = parent_map.get(drawing)
                    if parent is not None:
                        parent.remove(drawing)
                        changed = True
                    else:
                        # Si pas de parent direct, essayer de supprimer le run contenant le drawing
                        run = None
                        node2 = drawing
                        while node2 is not None:
                            if node2.tag == f"{{{W}}}r":
                                run = node2
                                break
                            node2 = parent_map.get(node2)
                        if run is not None:
                            run_parent = parent_map.get(run)
                            if run_parent is not None:
                                run_parent.remove(run)
                                changed = True
        
        # Supprimer les <v:imagedata r:id="rId"> et leurs <w:pict> parents
        for imagedata in root.findall(f".//v:imagedata", NS):
            rid = imagedata.get(f"{{{R}}}id")
            if rid and (rid in rids_to_remove or rid in all_rids_to_remove):
                # Remonter jusqu'à w:pict
                node = imagedata
                pict = None
                while node is not None:
                    if node.tag == f"{{{W}}}pict":
                        pict = node
                        break
                    node = parent_map.get(node)
                
                if pict is not None:
                    parent = parent_map.get(pict)
                    if parent is not None:
                        parent.remove(pict)
                        changed = True
        
        # Nettoyer les runs vides après suppression des drawings
        if changed:
            # Supprimer les runs qui ne contiennent plus rien
            for run in root.findall(".//w:r", NS):
                children = list(run)
                if not children or all(child.tag == f"{{{W}}}rPr" for child in children):
                    parent = parent_map.get(run)
                    if parent is not None:
                        parent.remove(run)
                        changed = True
            
            # Supprimer les paragraphes vides
            for para in root.findall(".//w:p", NS):
                children = list(para)
                if not children or all(child.tag in (f"{{{W}}}pPr", f"{{{W}}}rPr") for child in children):
                    # Vérifier qu'il n'y a pas de texte
                    text_content = "".join(t.text or "" for t in para.findall(".//w:t", NS))
                    if not text_content.strip():
                        parent = parent_map.get(para)
                        if parent is not None and parent.tag != f"{{{W}}}body":
                            parent.remove(para)
                            changed = True
        
        if changed:
            parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    
    # Supprimer les relations dans TOUS les .rels
    # Construire un set global de tous les rIds à supprimer
    all_rids_to_remove_global = set()
    for rids in media_to_rids.values():
        all_rids_to_remove_global.update(rids)
    
    for name in list(parts.keys()):
        if not name.endswith(".rels"):
            continue
        if "/_rels/" not in name:
            continue
        
        try:
            rels_root = ET.fromstring(parts[name])
        except ET.ParseError:
            continue
        
        base_name = name.replace("/_rels/", "/").replace(".rels", "")
        rids_to_remove = media_to_rids.get(base_name, set())
        
        # Utiliser aussi le set global pour être sûr de tout attraper
        if not rids_to_remove and not all_rids_to_remove_global:
            continue
        
        changed = False
        for rel in list(rels_root.findall(f".//{{{P_REL}}}Relationship")):
            rid = rel.get("Id") or ""
            tgt = rel.get("Target") or ""
            
            # Supprimer si le rId est dans la liste, OU si la cible est un SVG à supprimer
            should_remove = False
            if rid in rids_to_remove or rid in all_rids_to_remove_global:
                should_remove = True
            elif tgt:
                # Vérifier si la cible résolue est un SVG à supprimer
                resolved = _resolve_target_path(base_name, tgt)
                if resolved in svg_paths_to_remove:
                    should_remove = True
            
            if should_remove:
                rels_root.remove(rel)
                changed = True
        
        if changed:
            parts[name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    
    # Supprimer physiquement les fichiers SVG de parts
    for svg_path in svg_paths_to_remove:
        parts.pop(svg_path, None)

def _remove_specific_svg_in_part(
    parts: Dict[str, bytes],
    part_name: str,
    root: ET.Element,
    svg_media_paths: Set[str],
) -> None:
    """
    Supprime dans une partie donnée toutes les références aux SVG dont
    le chemin figure dans svg_media_paths :
      - <a:blip r:embed="rId"> et le <w:drawing> parent
      - la relation correspondante dans _rels/part.rels
    """
    if not svg_media_paths:
        return

    rels_name = _rels_name_for(part_name)
    if rels_name not in parts:
        return
    try:
        rels_root = ET.fromstring(parts[rels_name])
    except ET.ParseError:
        return

    # Map rId -> chemin media résolu
    rmap: Dict[str, str] = {}
    for rel in rels_root.findall(f".//{{{P_REL}}}Relationship"):
        rid = rel.get("Id") or ""
        tgt = rel.get("Target") or ""
        if not rid or not tgt:
            continue
        rmap[rid] = _resolve_target_path(part_name, tgt)

    # rId à supprimer (pointant vers un SVG modèle)
    rids_to_remove: Set[str] = set()
    for rid, mp in rmap.items():
        if mp in svg_media_paths:
            rids_to_remove.add(rid)

    if not rids_to_remove:
        return

    # Construire la map parent -> enfant pour pouvoir supprimer proprement
    parent_map = {child: parent for parent in root.iter() for child in parent}

    # Supprimer les dessins/blips référencés
    for blip in root.findall(".//a:blip", NS):
        rid = blip.get(f"{{{R}}}embed")
        if not rid or rid not in rids_to_remove:
            continue
        # Remonter jusqu'à w:drawing
        node = blip
        drawing = None
        while node is not None:
            if node.tag == f"{{{W}}}drawing":
                drawing = node
                break
            node = parent_map.get(node)
        if drawing is not None:
            parent = parent_map.get(drawing)
            if parent is not None:
                parent.remove(drawing)

    # Nettoyer les relations correspondantes
    changed = False
    for rel in list(rels_root.findall(f".//{{{P_REL}}}Relationship")):
        rid = rel.get("Id") or ""
        if rid in rids_to_remove:
            rels_root.remove(rel)
            changed = True

    if changed:
        parts[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

def _remove_megaphones_in_part(parts: Dict[str, bytes], part_name: str, root: ET.Element,
                               megaphone_hashes: Set[str], megaphone_ahashes: Set[int],
                               protected_hashes: Set[str], protected_ahashes: Set[int]) -> None:
    rels_name = _rels_name_for(part_name)
    if rels_name not in parts:
        return
    try:
        rels_root = ET.fromstring(parts[rels_name])
    except ET.ParseError:
        return

    rmap: Dict[str, str] = {}
    for rel in rels_root.findall(f".//{{{P_REL}}}Relationship"):
        rid = rel.get("Id") or ""
        tgt = rel.get("Target") or ""
        rmap[rid] = _resolve_target_path(part_name, tgt)

    parent_map = {child: parent for parent in root.iter() for child in parent}
    removed_rids: Set[str] = set()

    # 1) Images DrawingML : <a:blip r:embed="...">
    for blip in root.findall(".//a:blip", NS):
        rid = blip.get(f"{{{R}}}embed")
        if not rid or rid not in rmap:
            continue
        media_path = rmap[rid]
        if media_path not in parts:
            continue
        data = parts[media_path]
        
        # Vérifier si c'est un SVG
        is_svg = media_path.lower().endswith(".svg")
        # IMPORTANT : on laisse désormais TOUT le traitement des SVG
        # à la logique dédiée (_identify_svg_to_remove + remove_media_references)
        # pour éviter de supprimer les cibles ici par erreur.
        if is_svg:
            # Ne rien faire dans _remove_megaphones_in_part pour les SVG
            # (ils seront traités par le pipeline SVG séparé).
            continue
        
        data_hash = _sha1(data)
        data_ah = _ahash(data)

        # Icônes protégées (ex: Cible.png) : on ne les touche jamais.
        if data_hash in protected_hashes:
            continue
        if data_ah is not None and protected_ahashes:
            if min(_hamming(data_ah, ah) for ah in protected_ahashes) <= 5:
                continue

        # Mégaphones bitmap à supprimer : hash exact OU hash perceptuel proche
        match_hash = data_hash in megaphone_hashes if megaphone_hashes else False
        if not match_hash and data_ah is not None and megaphone_ahashes:
            if min(_hamming(data_ah, ah) for ah in megaphone_ahashes) <= 5:
                match_hash = True
        
        holder = None
        node = blip
        while node is not None:
            if node.tag in (f"{{{WP}}}anchor", f"{{{WP}}}inline"):
                holder = node; break
            node = parent_map.get(node)
        drawing = None
        node2 = blip
        while node2 is not None:
            if node2.tag == f"{{{W}}}drawing":
                drawing = node2; break
            node2 = parent_map.get(node2)

        # On supprime si :
        #   - l'empreinte correspond à un mégaphone (bitmap)
        #   - OU si c'est un SVG non-cible
        should_remove = match_hash

        if should_remove and drawing is not None:
            parent = parent_map.get(drawing)
            if parent is not None:
                parent.remove(drawing)
                removed_rids.add(rid)

    # 2) Images VML : <v:imagedata r:id="..."> à l'intérieur de <w:pict>
    for imdata in root.findall(".//v:imagedata", NS):
        rid = imdata.get(f"{{{R}}}id") or imdata.get(f"{{{R}}}embed")
        if not rid or rid not in rmap:
            continue
        media_path = rmap[rid]
        if media_path not in parts:
            continue
        data = parts[media_path]

        # VML porte souvent des bitmap (PNG/EMF) – on applique la même logique de hash
        data_hash = _sha1(data)
        data_ah = _ahash(data)

        if data_hash in protected_hashes:
            continue
        if data_ah is not None and protected_ahashes:
            if min(_hamming(data_ah, ah) for ah in protected_ahashes) <= 5:
                continue

        match_hash = data_hash in megaphone_hashes if megaphone_hashes else False
        if not match_hash and data_ah is not None and megaphone_ahashes:
            if min(_hamming(data_ah, ah) for ah in megaphone_ahashes) <= 5:
                match_hash = True

        if match_hash:
            # Remonter à <w:pict> et le supprimer
            node = imdata
            pict = None
            while node is not None:
                if node.tag == f"{{{W}}}pict":
                    pict = node
                    break
                node = parent_map.get(node)
            if pict is not None:
                parent = parent_map.get(pict)
                if parent is not None:
                    parent.remove(pict)
                    removed_rids.add(rid)

    if removed_rids:
        for rel in list(rels_root.findall(f".//{{{P_REL}}}Relationship")):
            if (rel.get("Id") or "") in removed_rids:
                rels_root.remove(rel)
        parts[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

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
    megaphone_samples: Optional[List[bytes]] = None,
) -> bytes:

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        parts: Dict[str, bytes] = {n: zin.read(n) for n in zin.namelist()}

    # Identifier tous les SVG à supprimer (tous sauf ceux contenant CIBLE_SVG_SNIP)
    # Cette logique fonctionnait mais laissait des carrés vides - on travaillera dessus ensuite
    svg_annonce_paths = _identify_svg_to_remove(parts)

    # Identifier les bitmaps carrés issus des annonces (via is_annonce_square_media)
    bitmap_annonce_media_paths: Set[str] = set()
    for name, data in parts.items():
        if is_annonce_square_media(name, data):
            bitmap_annonce_media_paths.add(name)

    # Union des media à supprimer (SVG annonces + bitmaps carrés)
    media_paths_to_remove = svg_annonce_paths | bitmap_annonce_media_paths

    # Debug détaillé
    total_svg_count = sum(1 for n in parts.keys() if n.lower().endswith(".svg") and "/media/" in n.lower())
    svg_annonce_count = len(svg_annonce_paths)
    svg_cible_count = total_svg_count - svg_annonce_count
    total_bitmap_count = sum(
        1
        for n in parts.keys()
        if n.lower().startswith("word/media/")
        and (n.lower().endswith(".png") or n.lower().endswith(".emf") or n.lower().endswith(".jpg") or n.lower().endswith(".jpeg"))
    )
    try:
        print(f"[DEBUG] Total SVG trouvés : {total_svg_count}")
        print(f"[DEBUG] SVG annonces à supprimer : {svg_annonce_count}")
        print(f"[DEBUG] SVG cibles à conserver : {svg_cible_count}")
        print(f"[DEBUG] Total bitmaps : {total_bitmap_count}")
        print(f"[DEBUG] Bitmaps carrés à supprimer : {len(bitmap_annonce_media_paths)}")
        if svg_annonce_paths:
            print(f"[DEBUG] Chemins SVG annonces : {list(svg_annonce_paths)[:3]}...")
        if bitmap_annonce_media_paths:
            print(f"[DEBUG] Chemins bitmaps carrés : {list(bitmap_annonce_media_paths)[:3]}...")
    except Exception:
        pass

    # Supprimer toutes les références aux media identifiés (SVG annonces + bitmaps carrés)
    # Fonction unifiée qui gère à la fois les SVG et les bitmaps
    remove_media_references(parts, media_paths_to_remove)

    theme_colors = extract_theme_colors(parts)

    # Construire la liste des empreintes d'icônes à supprimer :
    #   - exemples fournis via l'UI (échantillons mégaphone)
    #   - icônes Annonce1/Annonce2 du dossier assets
    default_meg_hashes, default_meg_ahashes = _load_default_megaphone_hashes()
    megaphone_hashes: Set[str] = set(default_meg_hashes)
    megaphone_ahashes: Set[int] = set(default_meg_ahashes)
    if megaphone_samples:
        for b in megaphone_samples:
            try:
                megaphone_hashes.add(_sha1(b))
                ah = _ahash(b)
                if ah is not None:
                    megaphone_ahashes.add(ah)
            except Exception:
                pass

    protected_hashes, protected_ahashes = _load_protected_icon_hashes()

    for name, data in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue

        # Texte & formats
        replace_years(root)
        strip_actualisation_everywhere(root)
        force_calibri(root)
        red_to_black(root)

        if name == "word/document.xml":
            cover_sizes_cleanup(root)
            tune_cover_shapes_spatial(root)
            tables_and_numbering(root)
            force_course_name_after_title_20(root)
            force_title_fiche_de_cours_22(root)
            reposition_small_icon(root, icon_left, icon_top)
            remove_large_grey_rectangles(root, theme_colors)
            force_red_bullets_black_in_paragraphs(root)

        if name == "word/numbering.xml":
            force_red_bullets_black_in_numbering(root)

        if name == "word/styles.xml":
            force_red_bullets_black_in_styles(root)

        if name.startswith("word/footer"):
            force_footer_size_10(root)

        _remove_megaphones_in_part(
            parts, name, root,
            megaphone_hashes, megaphone_ahashes,
            protected_hashes, protected_ahashes,
        )

        parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

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
st.caption("Transforme tes .docx 2024-2025 en 2025-2026 (couleurs, puces, tailles, rectangle gris, légende, mégaphones, etc.).")

with st.sidebar:
    st.subheader("Paramètres (cm)")
    icon_left  = st.number_input("Icône écriture — gauche", value=15.3, step=0.1)
    icon_top   = st.number_input("Icône écriture — haut",   value=11.0, step=0.1)
    legend_left= st.number_input("Image Légendes — gauche", value=2.3, step=0.1)
    legend_top = st.number_input("Image Légendes — haut",   value=23.8, step=0.1)
    legend_w   = st.number_input("Image Légendes — largeur",value=5.68, step=0.01)
    legend_h   = st.number_input("Image Légendes — hauteur",value=3.77, step=0.01)

# Charger l'image de légende par défaut depuis assets
default_legend_path = os.path.join("assets", "Legende.png")
default_legend_bytes = None
if os.path.exists(default_legend_path):
    try:
        with open(default_legend_path, "rb") as f:
            default_legend_bytes = f.read()
    except Exception:
        default_legend_bytes = None

st.markdown("**1) Glisse/dépose un ou plusieurs fichiers .docx**")
files = st.file_uploader("DOCX à traiter", type=["docx"], accept_multiple_files=True)
st.markdown("**2) (Optionnel) Remplace l'image de la Légende (PNG/JPG)**")
if default_legend_bytes:
    st.info("ℹ️ L'image `assets/Legende.png` sera utilisée par défaut si aucune image n'est fournie.")
legend_file = st.file_uploader("Image Légendes (optionnel)", type=["png","jpg","jpeg","webp"], accept_multiple_files=False)
st.markdown("**3) (Optionnel) Fourni 1–2 exemples d'icône mégaphone (PNG/JPG) pour détection par empreinte**")
megaphone_files = st.file_uploader("Icônes mégaphone (exemples)", type=["png","jpg","jpeg","webp"], accept_multiple_files=True)

if st.button("⚙️ Lancer le traitement", type="primary", disabled=not files):
    if not files:
        st.warning("Ajoute au moins un fichier .docx")
    else:
        legend_bytes = legend_file.read() if legend_file else default_legend_bytes
        megaphone_samples = [f.read() for f in megaphone_files] if megaphone_files else None

        processed: List[Tuple[str, bytes]] = []
        errors: List[str] = []

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
                    megaphone_samples=megaphone_samples,
                )
                out_name = cleaned_filename(up.name)
                processed.append((out_name, out_bytes))
                st.success(f"✅ Terminé : {up.name} → {out_name}")
            except Exception as e:
                errors.append(f"{up.name} : {e}")

        if errors:
            st.error("Quelques fichiers ont échoué :\n- " + "\n- ".join(errors))

        if processed:
            # Crée un ZIP avec tous les fichiers modifiés
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for fname, fbytes in processed:
                    z.writestr(fname, fbytes)
            zip_buf.seek(0)
            st.download_button(
                "⬇️ Télécharger le ZIP de tous les fichiers modifiés",
                data=zip_buf,
                file_name="fiches_modifiees.zip",
                mime="application/zip",
            )

"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  DKB POWER ENGINEERING CO., LTD.                                            ║
║  Company Profile Dashboard  v2.0  —  dkb_profile_dashboard.py               ║
║                                                                              ║
║  Run:  streamlit run dkb_profile_dashboard.py                               ║
║  Deploy: push to GitHub → connect to Streamlit Cloud                        ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

# ── Standard library ──────────────────────────────────────────────────────────
import io
import json
import os
import re
import uuid
from copy import deepcopy
from datetime import datetime
from pathlib import Path

# ── Third-party ───────────────────────────────────────────────────────────────
import streamlit as st
from PIL import Image, ImageOps

# ── python-docx ───────────────────────────────────────────────────────────────
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ══════════════════════════════════════════════════════════════════════════════
#  PATHS
# ══════════════════════════════════════════════════════════════════════════════

BASE_DIR   = Path(__file__).parent
PHOTOS_DIR = BASE_DIR / "photos"
OUTPUT_DIR = BASE_DIR / "output"
ASSETS_DIR = BASE_DIR / "assets"
DATA_FILE  = BASE_DIR / "projects.json"
COMPANY_FILE = BASE_DIR / "company.json"

for d in (PHOTOS_DIR, OUTPUT_DIR, ASSETS_DIR):
    d.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
#  DEFAULT COMPANY DATA
# ══════════════════════════════════════════════════════════════════════════════

DEFAULT_COMPANY = {
    "name_th"     : "บริษัท ดีเคบี เพาเวอร์ เอนจิเนียริ่ง จำกัด",
    "name_en"     : "DKB POWER ENGINEERING CO., LTD.",
    "reg_no"      : "0105565009633",
    "reg_date_th" : "17 มกราคม 2565",
    "reg_date_en" : "17 January 2022",
    "capital_th"  : "1,000,000.00 บาท",
    "capital_en"  : "THB 1,000,000.00",
    "address_th"  : "422/635/1 ซอยสุจินทวงศ์ 34 แขวงเสน่แสบ เขตมีนบุรี กรุงเทพมหานคร",
    "address_en"  : "422/635/1 Soi Sujinwong 34, Saen Saeb, Min Buri, Bangkok",
    "tel"         : "082-234-4680",
    "email"       : "kasama.D.K.B@gmail.com",
    "year"        : "2026",
    "history_th"  : (
        "บริษัท ดีเคบี เพาเวอร์ เอนจิเนียริ่ง จำกัด ก่อตั้งขึ้นด้วยเจตนารมณ์อันแน่วแน่ "
        "ในการให้บริการด้านวิศวกรรมระบบไฟฟ้าและสื่อสารอย่างครบวงจร มุ่งเน้นงานติดตั้ง"
        "ระบบไฟฟ้า ระบบเสาอากาศโทรทัศน์ ระบบกล้องวงจรปิด (CCTV) ระบบทีวีดาวเทียม "
        "และระบบทีวีดิจิทัล ให้แก่ลูกค้าทั้งภาครัฐและภาคเอกชน ด้วยทีมช่างผู้ชำนาญการ "
        "ที่มุ่งมั่นส่งมอบงานคุณภาพสูง ตรงเวลา และอยู่ภายในงบประมาณที่ตกลงไว้ "
        "เรามุ่งสร้างความไว้วางใจและความพึงพอใจสูงสุดให้แก่ลูกค้าทุกราย"
    ),
    "history_en"  : (
        "DKB Power Engineering Co., Ltd. was established with an unwavering commitment "
        "to delivering comprehensive electrical and communication engineering services. "
        "Specialising in electrical installations, TV antenna systems, CCTV, satellite TV, "
        "and digital TV, the company serves both public and private sector clients. "
        "Our certified technical team upholds the highest engineering standards, "
        "delivering every project on time and within budget. "
        "We are committed to building lasting trust and maximum satisfaction with every client."
    ),
    "logo_file"   : "",   # filename in assets/
}

DEFAULT_PROJECT = {
    "id"          : "proj_001",
    "no"          : 1,
    "name_th"     : "โครงการเรือนแถวตำรวจ สภ.บางโทรัด จ.สมุทรสาคร",
    "name_en"     : "Police Row House Project, Bang Thorat Police Station, Samut Sakhon",
    "customer_th" : "สภ.บางโทรัด จ.สมุทรสาคร",
    "customer_en" : "Bang Thorat Police Station, Samut Sakhon Province",
    "main_contact": "บ.ณัฐกิจการช่าง",
    "location"    : "จ.สมุทรสาคร / Samut Sakhon",
    "desc_th"     : "งานติดตั้งระบบทีวี",
    "desc_en"     : "TV System Installation",
    "period"      : "2567",
    "value"       : "35,310.00",
    "photos"      : [],
    "created_at"  : "2026-01-01T00:00:00",
}

# ── Colour palette (hex strings for XML + RGBColor for python-docx) ───────────
HEX = {
    "navy"      : "0A3C7A",
    "blue"      : "2979C8",
    "darknavy"  : "062A5A",
    "lightblue" : "EDF4FC",
    "rowalt"    : "F7FAFE",
    "white"     : "FFFFFF",
    "midgray"   : "607A99",
    "textdark"  : "1A2E45",
    "steelblue" : "90B8E0",
}

def rgb(key): return RGBColor.from_string(HEX[key])


# ══════════════════════════════════════════════════════════════════════════════
#  DATA LAYER
# ══════════════════════════════════════════════════════════════════════════════

def load_company() -> dict:
    if COMPANY_FILE.exists():
        try:
            return {**DEFAULT_COMPANY, **json.loads(COMPANY_FILE.read_text("utf-8"))}
        except Exception:
            pass
    return deepcopy(DEFAULT_COMPANY)


def save_company(data: dict) -> None:
    COMPANY_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")


def load_projects() -> list:
    if not DATA_FILE.exists():
        save_projects([DEFAULT_PROJECT])
        return [deepcopy(DEFAULT_PROJECT)]
    try:
        data = json.loads(DATA_FILE.read_text("utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_projects(projects: list) -> None:
    DATA_FILE.write_text(json.dumps(projects, ensure_ascii=False, indent=2), "utf-8")


def next_no(projects: list) -> int:
    return max((p.get("no", 0) for p in projects), default=0) + 1


def safe_fname(original: str) -> str:
    stem = re.sub(r"[^\w\-.]", "_", Path(original).stem)[:40]
    return f"{uuid.uuid4().hex[:8]}_{stem}{Path(original).suffix.lower()}"


def save_photos(files) -> list:
    saved = []
    for f in (files or []):
        name = safe_fname(f.name)
        (PHOTOS_DIR / name).write_bytes(f.read())
        saved.append(name)
    return saved


def save_logo(file) -> str:
    name = safe_fname(file.name)
    (ASSETS_DIR / name).write_bytes(file.read())
    return name


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX — XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _cell_bg(cell, hex6: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex6)
    tcPr.append(shd)


def _cell_borders(cell, color="D8E6F2", size=4, sides=("top","left","bottom","right")):
    tcPr    = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for s in sides:
        el = OxmlElement(f"w:{s}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def _cell_no_border(cell):
    tcPr    = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for s in ("top","left","bottom","right","insideH","insideV"):
        el = OxmlElement(f"w:{s}")
        el.set(qn("w:val"), "none"); el.set(qn("w:sz"), "0"); el.set(qn("w:color"), "auto")
        borders.append(el)
    tcPr.append(borders)


def _cell_margins(cell, top=80, bottom=80, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    mar  = OxmlElement("w:tcMar")
    for s, v in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        el = OxmlElement(f"w:{s}")
        el.set(qn("w:w"), str(v)); el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def _col_width(cell, twips: int):
    tcPr = cell._tc.get_or_add_tcPr()
    w    = OxmlElement("w:tcW")
    w.set(qn("w:w"), str(int(twips))); w.set(qn("w:type"), "dxa")
    tcPr.append(w)


def _para_space(p, before=0, after=0):
    pPr = p._p.get_or_add_pPr()
    spc = OxmlElement("w:spacing")
    spc.set(qn("w:before"), str(before)); spc.set(qn("w:after"), str(after))
    pPr.append(spc)


def _para_indent(p, left=0, right=0):
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    if left:  ind.set(qn("w:left"),  str(left))
    if right: ind.set(qn("w:right"), str(right))
    pPr.append(ind)


def _bottom_rule(p, color="2979C8", size=8):
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), str(size))
    bot.set(qn("w:color"), color); bot.set(qn("w:space"), "4")
    pBdr.append(bot); pPr.append(pBdr)


def _left_bar(p, color="2979C8", size=18):
    """Thick left border = visual section marker."""
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    lft  = OxmlElement("w:left")
    lft.set(qn("w:val"), "single"); lft.set(qn("w:sz"), str(size))
    lft.set(qn("w:color"), color); lft.set(qn("w:space"), "6")
    pBdr.append(lft); pPr.append(pBdr)


def _run(p, text, bold=False, italic=False, size=10,
         color=None, font="TH Sarabun New"):
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.name = font; r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return r


def _cp(cell, idx=0):
    """Return paragraph idx in cell; extend if needed."""
    while len(cell.paragraphs) <= idx:
        cell.add_paragraph()
    return cell.paragraphs[idx]


def _spacer(doc, pts=6):
    p = doc.add_paragraph()
    _para_space(p, before=0, after=0)
    p._p.get_or_add_pPr()
    # set exact line height to pts
    pPr = p._p.get_or_add_pPr()
    spc = OxmlElement("w:spacing")
    spc.set(qn("w:line"), str(int(pts * 20)))
    spc.set(qn("w:lineRule"), "exact")
    pPr.append(spc)


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX — REUSABLE LAYOUT BLOCKS
# ══════════════════════════════════════════════════════════════════════════════

def _section_header(doc, th: str, en: str, rule_color="2979C8"):
    """
    ▌ Thai label  /  ENGLISH  ─────────────────
    """
    p = doc.add_paragraph()
    _para_space(p, before=160, after=60)
    _left_bar(p, color=HEX["navy"], size=20)
    _run(p, f"  {th}", bold=True, size=11, color=rgb("navy"))
    _run(p, f"  /  {en.upper()}", size=8,
         color=rgb("midgray"), font="Calibri")
    _bottom_rule(p, color=rule_color, size=6)


def _page_header_block(doc, co: dict):
    """Thin top accent + running header with company name."""
    # top navy rule
    p = doc.add_paragraph()
    _para_space(p, before=0, after=0)
    _bottom_rule(p, color=HEX["navy"], size=28)

    # running company name line
    p2 = doc.add_paragraph()
    _para_space(p2, before=40, after=80)
    _run(p2, co["name_th"] + "  ", bold=True, size=9, color=rgb("navy"))
    _run(p2, co["name_en"], size=8, color=rgb("midgray"), font="Calibri")
    _bottom_rule(p2, color=HEX["blue"], size=4)


def _page_footer_block(doc, co: dict, page_num: int):
    """Bottom navy bar with contact + page number."""
    _spacer(doc, 10)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    items = [
        f"โทร  {co['tel']}",
        f"อีเมล  {co['email']}",
        f"หน้า / Page  {page_num}",
    ]
    widths = [int(Cm(6.0).twips), int(Cm(8.0).twips), int(Cm(4.0).twips)]
    aligns = [WD_ALIGN_PARAGRAPH.LEFT,
              WD_ALIGN_PARAGRAPH.CENTER,
              WD_ALIGN_PARAGRAPH.RIGHT]
    for j, (txt, w, aln) in enumerate(zip(items, widths, aligns)):
        c = tbl.rows[0].cells[j]
        _cell_bg(c, HEX["darknavy"])
        _cell_borders(c, color=HEX["navy"], size=4)
        _cell_margins(c, top=100, bottom=100, left=120, right=120)
        _col_width(c, w)
        p = _cp(c)
        p.alignment = aln
        _run(p, txt, size=8, color=rgb("white"), font="TH Sarabun New")


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX — PAGE BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _build_cover(doc: Document, co: dict):
    """Page 1 — Professional cover."""

    # ── Thick top rule ──────────────────────────────────────────────────────
    p = doc.add_paragraph()
    _para_space(p, before=0, after=0)
    _bottom_rule(p, color=HEX["navy"], size=32)

    _spacer(doc, 20)

    # ── Logo (if provided) ──────────────────────────────────────────────────
    logo_path = ASSETS_DIR / co.get("logo_file", "") if co.get("logo_file") else None
    if logo_path and logo_path.exists():
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _para_space(lp, before=0, after=80)
        run = lp.add_run()
        run.add_picture(str(logo_path), height=Cm(1.6))

    # ── Company name ────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    _para_space(p, before=0, after=20)
    _run(p, co["name_th"], bold=True, size=15, color=rgb("navy"))

    p2 = doc.add_paragraph()
    _para_space(p2, before=0, after=60)
    _run(p2, co["name_en"], bold=True, size=10,
         color=rgb("blue"), font="Calibri")

    # ── Giant "Company Profile" title ───────────────────────────────────────
    p3 = doc.add_paragraph()
    _para_space(p3, before=200, after=40)
    _run(p3, "Company Profile", bold=True, size=42,
         color=rgb("navy"), font="Calibri")

    p4 = doc.add_paragraph()
    _para_space(p4, before=0, after=30)
    _run(p4, "โปรไฟล์บริษัท", size=16, color=rgb("midgray"))

    # ── Blue rule divider ────────────────────────────────────────────────────
    _bottom_rule(doc.add_paragraph(), color=HEX["blue"], size=12)

    # ── Descriptor ──────────────────────────────────────────────────────────
    _spacer(doc, 12)
    p5 = doc.add_paragraph()
    _para_space(p5, before=0, after=40)
    _run(p5,
         "ผู้เชี่ยวชาญด้านวิศวกรรมระบบไฟฟ้า ระบบสื่อสาร ระบบโทรทัศน์ดิจิทัล "
         "ดาวเทียม และกล้องวงจรปิด\n",
         size=11, color=rgb("textdark"))
    _run(p5,
         "Specialist in Electrical Systems, TV Antenna, Digital TV, "
         "Satellite TV, CCTV & Communication Engineering",
         size=9, color=rgb("midgray"), font="Calibri", italic=True)

    # ── Service tags (text-based pill row) ───────────────────────────────────
    _spacer(doc, 12)
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _para_space(p6, before=0, after=0)
    tags = ["Electrical","TV Antenna","Digital TV","Satellite TV","CCTV","Communication"]
    _run(p6, "  ·  ".join(tags),
         size=9, color=rgb("blue"), font="Calibri")

    # ── Large faded year ────────────────────────────────────────────────────
    _spacer(doc, 80)
    py = doc.add_paragraph()
    _para_space(py, before=0, after=0)
    _run(py, co["year"], bold=True, size=52,
         color=RGBColor(0xD8, 0xE8, 0xF5), font="Calibri")

    # ── Footer contact bar ───────────────────────────────────────────────────
    _spacer(doc, 20)
    _page_footer_block(doc, co, 1)

    doc.add_page_break()


def _build_company_info(doc: Document, co: dict):
    """Page 2 — Company information."""
    _page_header_block(doc, co)

    # ── Page title ───────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    _para_space(p, before=20, after=60)
    _run(p, "ข้อมูลบริษัท  ", bold=True, size=15, color=rgb("navy"))
    _run(p, "COMPANY INFORMATION",
         size=9, color=rgb("midgray"), font="Calibri")
    _bottom_rule(p, color=HEX["blue"], size=12)

    # ── Company name card ────────────────────────────────────────────────────
    nb = doc.add_table(rows=1, cols=1)
    nb.alignment = WD_TABLE_ALIGNMENT.LEFT
    c = nb.rows[0].cells[0]
    _cell_bg(c, HEX["lightblue"])
    _cell_borders(c, color=HEX["blue"], size=8,
                  sides=("top","left","bottom","right"))
    _cell_margins(c, top=120, bottom=120, left=200, right=200)
    _col_width(c, int(Cm(18.0).twips))
    lp = _cp(c)
    _run(lp, co["name_th"] + "\n", bold=True, size=14, color=rgb("navy"))
    _run(lp, co["name_en"], bold=True, size=10,
         color=rgb("blue"), font="Calibri")

    _spacer(doc, 14)

    # ── 2-col layout: Registration (left) | History (right) ─────────────────
    body = doc.add_table(rows=1, cols=2)
    body.alignment = WD_TABLE_ALIGNMENT.LEFT
    lc = body.rows[0].cells[0]   # registration
    rc = body.rows[0].cells[1]   # history + services
    _cell_no_border(lc); _cell_no_border(rc)
    _cell_margins(lc, top=0, bottom=0, left=0, right=120)
    _cell_margins(rc, top=0, bottom=0, left=120, right=0)
    _col_width(lc, int(Cm(8.5).twips))
    _col_width(rc, int(Cm(9.0).twips))

    # ── Left: Registration table ─────────────────────────────────────────────
    rh = lc.add_paragraph()
    _para_space(rh, before=0, after=40)
    _run(rh, "ข้อมูลการจดทะเบียน  ", bold=True, size=9, color=rgb("navy"))
    _run(rh, "REGISTRATION", size=7.5, color=rgb("midgray"), font="Calibri")
    _bottom_rule(rh, color=HEX["blue"], size=6)

    fields = [
        ("เลขทะเบียนนิติบุคคล\nRegistration No.",
         co["reg_no"]),
        ("วันที่จดทะเบียน\nDate of Incorporation",
         co["reg_date_th"] + "\n" + co["reg_date_en"]),
        ("ทุนจดทะเบียน\nRegistered Capital",
         co["capital_th"] + "\n" + co["capital_en"]),
        ("ที่ตั้งสำนักงาน\nOffice Address",
         co["address_th"] + "\n" + co["address_en"]),
        ("โทรศัพท์ / Tel",  co["tel"]),
        ("อีเมล / Email",   co["email"]),
    ]
    reg = lc.add_table(rows=len(fields), cols=2)
    reg.style = "Table Grid"
    for i, (lbl, val) in enumerate(fields):
        lcc = reg.rows[i].cells[0]
        vcc = reg.rows[i].cells[1]
        _cell_bg(lcc, HEX["lightblue"] if i % 2 == 0 else "F4F8FC")
        _cell_bg(vcc, HEX["white"])
        _cell_borders(lcc, color="D8E6F2"); _cell_borders(vcc, color="D8E6F2")
        _cell_margins(lcc, top=60, bottom=60, left=80, right=60)
        _cell_margins(vcc, top=60, bottom=60, left=80, right=60)
        _col_width(lcc, int(Cm(3.2).twips))
        _col_width(vcc, int(Cm(5.3).twips))
        lp2 = _cp(lcc)
        for k, ln in enumerate(lbl.split("\n")):
            r = lp2.add_run(ln + ("\n" if k == 0 and "\n" in lbl else ""))
            r.font.name = "TH Sarabun New"
            r.font.size = Pt(8.5 if k == 0 else 7)
            r.font.color.rgb = rgb("navy") if k == 0 else rgb("midgray")
            r.bold = (k == 0)
        vp2 = _cp(vcc)
        for k, ln in enumerate(val.split("\n")):
            r = vp2.add_run(ln + ("\n" if k < len(val.split("\n")) - 1 else ""))
            r.font.name = "TH Sarabun New"
            r.font.size = Pt(9 if k == 0 else 7.5)
            r.font.color.rgb = rgb("textdark") if k == 0 else rgb("midgray")
            r.bold = (k == 0)

    # ── Right: History ───────────────────────────────────────────────────────
    hh = rc.add_paragraph()
    _para_space(hh, before=0, after=40)
    _run(hh, "ประวัติและพันธกิจ  ", bold=True, size=9, color=rgb("navy"))
    _run(hh, "HISTORY & MISSION", size=7.5, color=rgb("midgray"), font="Calibri")
    _bottom_rule(hh, color=HEX["blue"], size=6)

    ht = rc.add_table(rows=1, cols=1)
    hc = ht.rows[0].cells[0]
    _cell_bg(hc, "F4F8FC")
    _cell_borders(hc, color=HEX["blue"], size=6)
    _cell_margins(hc, top=100, bottom=100, left=140, right=140)
    _col_width(hc, int(Cm(9.0).twips))
    hp = _cp(hc)
    _run(hp, co["history_th"] + "\n\n",
         size=9, color=rgb("textdark"))
    _run(hp, co["history_en"],
         size=8, color=rgb("midgray"), font="Calibri", italic=True)

    # ── Services grid (below history) ────────────────────────────────────────
    _spacer_in_cell = rc.add_paragraph()
    _para_space(_spacer_in_cell, before=80, after=40)

    sh = rc.add_paragraph()
    _para_space(sh, before=0, after=40)
    _run(sh, "ขอบเขตงานบริการ  ", bold=True, size=9, color=rgb("navy"))
    _run(sh, "SCOPE OF SERVICES", size=7.5, color=rgb("midgray"), font="Calibri")
    _bottom_rule(sh, color=HEX["blue"], size=6)

    services = [
        ("ระบบไฟฟ้า",        "Electrical Installation"),
        ("ระบบกล้อง CCTV",   "CCTV System"),
        ("ระบบทีวีดิจิทัล", "Digital TV System"),
        ("ทีวีดาวเทียม",     "Satellite TV"),
        ("เสาอากาศทีวี",     "TV Antenna System"),
        ("งานวิศวกรรมอื่นๆ","Other Eng. Works"),
    ]
    sg = rc.add_table(rows=3, cols=2)
    sg.alignment = WD_TABLE_ALIGNMENT.LEFT
    idx = 0
    for ri in range(3):
        for ci in range(2):
            if idx >= len(services): break
            sc = sg.rows[ri].cells[ci]
            _cell_bg(sc, HEX["lightblue"])
            _cell_borders(sc, color="D8E6F2")
            _cell_margins(sc, top=70, bottom=70, left=90, right=70)
            _col_width(sc, int(Cm(4.5).twips))
            sp = _cp(sc)
            _run(sp, services[idx][0] + "\n",
                 bold=True, size=8.5, color=rgb("navy"))
            _run(sp, services[idx][1],
                 size=7.5, color=rgb("midgray"), font="Calibri")
            idx += 1

    _page_footer_block(doc, co, 2)
    doc.add_page_break()


def _build_project_reference(doc: Document, co: dict, projects: list):
    """Page 3 — Project reference table."""
    _page_header_block(doc, co)

    p = doc.add_paragraph()
    _para_space(p, before=20, after=60)
    _run(p, "รายการอ้างอิงโครงการ  ", bold=True, size=15, color=rgb("navy"))
    _run(p, "PROJECT REFERENCE LIST",
         size=9, color=rgb("midgray"), font="Calibri")
    _bottom_rule(p, color=HEX["blue"], size=12)
    _spacer(doc, 8)

    # Column config: (th_label, en_label, width_cm)
    cols = [
        ("ที่",              "No.",              0.9),
        ("ชื่อโครงการ",      "Project Name",     4.2),
        ("ผู้ว่าจ้าง/ลูกค้า","Customer / Owner",  3.4),
        ("ผู้รับช่วง",       "Main Contact",      2.8),
        ("สถานที่",          "Location",          2.6),
        ("รายละเอียดงาน",    "Job Description",   3.1),
        ("มูลค่า (บาท)",     "Value (THB)",       2.0),
    ]

    tbl = doc.add_table(rows=1 + len(projects), cols=len(cols))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header
    for j, (th, en, wcm) in enumerate(cols):
        cell = tbl.rows[0].cells[j]
        _cell_bg(cell, HEX["navy"])
        _cell_borders(cell, color=HEX["white"], size=4)
        _cell_margins(cell, top=80, bottom=80, left=90, right=90)
        _col_width(cell, int(Cm(wcm).twips))
        hp = _cp(cell)
        hp.alignment = (WD_ALIGN_PARAGRAPH.CENTER if j in (0, 6)
                        else WD_ALIGN_PARAGRAPH.LEFT)
        _run(hp, th + "\n", bold=True, size=8.5, color=rgb("white"))
        _run(hp, en, size=7, color=RGBColor(0x90, 0xB8, 0xE0), font="Calibri")

    # Data rows
    aligns = [WD_ALIGN_PARAGRAPH.CENTER] + [WD_ALIGN_PARAGRAPH.LEFT] * 5 + [WD_ALIGN_PARAGRAPH.RIGHT]
    for i, proj in enumerate(projects):
        bg = HEX["rowalt"] if i % 2 == 0 else HEX["white"]
        vals = [
            str(proj.get("no", i + 1)),
            proj.get("name_th","") + "\n" + proj.get("name_en",""),
            proj.get("customer_th","") + "\n" + proj.get("customer_en",""),
            proj.get("main_contact","—"),
            proj.get("location","—"),
            proj.get("desc_th","") + "\n" + proj.get("desc_en",""),
            "฿ " + proj.get("value","—"),
        ]
        for j, (val, aln, (_, _, wcm)) in enumerate(zip(vals, aligns, cols)):
            cell = tbl.rows[1+i].cells[j]
            _cell_bg(cell, bg)
            _cell_borders(cell, color="D8E6F2")
            _cell_margins(cell, top=70, bottom=70, left=90, right=90)
            _col_width(cell, int(Cm(wcm).twips))
            dp = _cp(cell)
            dp.alignment = aln
            lines = val.split("\n")
            for k, ln in enumerate(lines):
                r = dp.add_run(ln + ("\n" if k < len(lines)-1 else ""))
                r.font.name  = "TH Sarabun New"
                r.font.size  = Pt(9.5 if k == 0 else 7.5)
                r.font.color.rgb = (rgb("navy")     if j == 1 and k == 0 else
                                    rgb("blue")     if j == 0 else
                                    rgb("textdark") if k == 0 else rgb("midgray"))
                r.bold = (k == 0 and j in (0, 1, 6))

    _page_footer_block(doc, co, 3)
    doc.add_page_break()


def _build_case_study(doc: Document, co: dict, proj: dict, page_num: int):
    """One case study page per project."""
    _page_header_block(doc, co)

    # Section title
    p = doc.add_paragraph()
    _para_space(p, before=20, after=60)
    _run(p, "กรณีศึกษาโครงการ  ", bold=True, size=15, color=rgb("navy"))
    _run(p, "CASE STUDY FOR PROJECTS REFERENCE",
         size=9, color=rgb("midgray"), font="Calibri")
    _bottom_rule(p, color=HEX["blue"], size=12)
    _spacer(doc, 8)

    # Project banner card
    bt = doc.add_table(rows=1, cols=1)
    bt.alignment = WD_TABLE_ALIGNMENT.LEFT
    bc = bt.rows[0].cells[0]
    _cell_bg(bc, HEX["lightblue"])
    _cell_borders(bc, color=HEX["navy"], size=12)
    _cell_margins(bc, top=110, bottom=110, left=200, right=140)
    _col_width(bc, int(Cm(18.0).twips))
    lp3 = _cp(bc)
    _run(lp3,
         f"โครงการที่ {proj.get('no','—')}  ·  {proj.get('name_th','')}",
         bold=True, size=12, color=rgb("navy"))
    lp3.add_run("\n")
    _run(lp3,
         f"Project No.{proj.get('no','—')}  ·  {proj.get('name_en','')}",
         size=9, color=rgb("midgray"), font="Calibri", italic=True)

    _spacer(doc, 12)

    # Summary table
    _section_header(doc, "สรุปข้อมูลโครงการ", "Project Summary")
    _spacer(doc, 6)

    rows_data = [
        ("ชื่อโครงการ\nProject Name",
         proj.get("name_th","") + "\n" + proj.get("name_en","")),
        ("ลูกค้า / เจ้าของโครงการ\nCustomer / Owner",
         proj.get("customer_th","") + "\n" + proj.get("customer_en","")),
        ("ผู้รับช่วงงาน\nMain Contact",
         proj.get("main_contact","—")),
        ("รายละเอียดงาน\nJob Description",
         proj.get("desc_th","") + "\n" + proj.get("desc_en","")),
        ("สถานที่ดำเนินงาน\nSite Location",
         proj.get("location","—")),
        ("ระยะเวลาดำเนินงาน\nProject Period",
         proj.get("period","—")),
        ("มูลค่างาน\nContract Value (THB)",
         f"฿  {proj.get('value','—')}  บาท"),
    ]
    lw = int(Cm(4.5).twips)
    vw = int(Cm(13.5).twips)

    st2 = doc.add_table(rows=len(rows_data), cols=2)
    st2.style = "Table Grid"
    st2.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, (lbl, val) in enumerate(rows_data):
        lc2, vc2 = st2.rows[i].cells[0], st2.rows[i].cells[1]
        is_value = lbl.startswith("มูลค่า")
        _cell_bg(lc2, "092F63" if is_value else HEX["navy"])
        _cell_bg(vc2, "EDF5FF" if is_value else
                 (HEX["rowalt"] if i % 2 == 0 else HEX["white"]))
        _cell_borders(lc2, color=HEX["white"], size=4)
        _cell_borders(vc2, color="D8E6F2")
        _cell_margins(lc2, top=80, bottom=80, left=120, right=80)
        _cell_margins(vc2, top=80, bottom=80, left=120, right=80)
        _col_width(lc2, lw); _col_width(vc2, vw)

        lp4 = _cp(lc2)
        for k, ln in enumerate(lbl.split("\n")):
            r = lp4.add_run(ln + ("\n" if k == 0 and "\n" in lbl else ""))
            r.font.name = "TH Sarabun New"
            r.font.size = Pt(9 if k == 0 else 7.5)
            r.font.color.rgb = rgb("white") if k == 0 else RGBColor(0x90,0xB8,0xE0)
            r.bold = (k == 0)

        vp3 = _cp(vc2)
        for k, ln in enumerate(val.split("\n")):
            r = vp3.add_run(ln + ("\n" if k < len(val.split("\n"))-1 else ""))
            r.font.name  = "TH Sarabun New"
            r.font.size  = Pt(12 if is_value and k == 0 else
                              (9.5 if k == 0 else 8))
            r.font.color.rgb = rgb("navy") if is_value else (
                rgb("textdark") if k == 0 else rgb("midgray"))
            r.bold = is_value and k == 0

    _spacer(doc, 14)

    # Photo gallery
    _section_header(doc, "ภาพถ่ายผลงานโครงการ", "Project Photography")
    _spacer(doc, 8)

    photos = [PHOTOS_DIR / fn for fn in proj.get("photos", [])
              if (PHOTOS_DIR / fn).exists()]
    COLS   = 4
    IMG_W  = Cm(4.0)

    if photos:
        chunks = [photos[i:i+COLS] for i in range(0, len(photos), COLS)]
        for chunk in chunks:
            pt = doc.add_table(rows=1, cols=COLS)
            pt.alignment = WD_TABLE_ALIGNMENT.LEFT
            for j in range(COLS):
                cell = pt.rows[0].cells[j]
                _cell_no_border(cell)
                _cell_margins(cell, top=40, bottom=40, left=40, right=40)
                _col_width(cell, int(Cm(4.3).twips))
                pp = _cp(cell)
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if j < len(chunk):
                    try:
                        # Resize image to avoid oversized embeds
                        img = Image.open(str(chunk[j]))
                        img = ImageOps.exif_transpose(img)  # auto-rotate
                        max_px = 800
                        if max(img.size) > max_px:
                            img.thumbnail((max_px, max_px), Image.LANCZOS)
                        buf = io.BytesIO()
                        img.save(buf, format="JPEG", quality=85)
                        buf.seek(0)
                        run = pp.add_run()
                        run.add_picture(buf, width=IMG_W)
                    except Exception:
                        _run(pp, f"[Photo {j+1}]", size=8,
                             color=rgb("midgray"), font="Calibri", italic=True)
                else:
                    _cell_bg(cell, "F4F8FC")
            _spacer(doc, 6)
    else:
        # Placeholder grid when no photos
        for _ in range(2):
            pt = doc.add_table(rows=1, cols=COLS)
            pt.style = "Table Grid"
            pt.alignment = WD_TABLE_ALIGNMENT.LEFT
            for j in range(COLS):
                cell = pt.rows[0].cells[j]
                _cell_bg(cell, HEX["lightblue"])
                _cell_borders(cell, color="C8D8E8")
                _cell_margins(cell)
                _col_width(cell, int(Cm(4.3).twips))
                pp = _cp(cell)
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(pp, "[ รูปภาพ / Photo ]",
                     size=8, color=RGBColor(0xAA, 0xBC, 0xCC),
                     font="Calibri", italic=True)
            _spacer(doc, 6)

    _page_footer_block(doc, co, page_num)
    doc.add_page_break()


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX — MAIN GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_docx(co: dict, projects: list) -> bytes:
    """
    Build the full Company Profile .docx.
    Returns raw bytes for st.download_button.
    Saves a timestamped copy to output/.
    """
    doc = Document()

    # A4 page setup
    for sec in doc.sections:
        sec.page_width    = Cm(21.0)
        sec.page_height   = Cm(29.7)
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(2.0)
        sec.right_margin  = Cm(2.0)

    # Default paragraph style
    style = doc.styles["Normal"]
    style.font.name = "TH Sarabun New"
    style.font.size = Pt(10)

    _build_cover(doc, co)
    _build_company_info(doc, co)
    _build_project_reference(doc, co, projects)
    for i, proj in enumerate(projects):
        _build_case_study(doc, co, proj, 4 + i)

    buf = io.BytesIO()
    doc.save(buf)
    raw = buf.getvalue()

    # Save timestamped copy
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    out  = OUTPUT_DIR / f"DKB_Profile_{ts}.docx"
    out.write_bytes(raw)

    return raw


# ══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT  UI
# ══════════════════════════════════════════════════════════════════════════════

# ─── CSS ─────────────────────────────────────────────────────────────────────
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Sarabun',sans-serif!important;}

/* Sidebar */
[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#062A5A 0%,#0A3C7A 60%,#0D4A8E 100%)!important;
}
[data-testid="stSidebar"] *{color:#E8F0FA!important;}
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3{color:#FFF!important;font-weight:700!important;}
[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.2)!important;}
[data-testid="stSidebar"] small{color:#90B8E0!important;}
[data-testid="stSidebar"] .stButton button{
  background:rgba(255,255,255,.12)!important;color:#fff!important;
  border:1px solid rgba(255,255,255,.25)!important;border-radius:6px!important;
  width:100%!important;font-size:12px!important;}

/* Main container */
.main .block-container{max-width:1000px;padding-top:1.5rem;}

/* Header banner */
.app-header{
  background:linear-gradient(135deg,#0A3C7A 0%,#2979C8 100%);
  border-radius:12px;padding:22px 28px;margin-bottom:20px;
}
.app-header .tag{font-size:10px;letter-spacing:3px;color:#90C8F0;
  text-transform:uppercase;margin-bottom:4px;}
.app-header h1{font-size:22px;font-weight:700;color:#fff;margin:0;line-height:1.3;}
.app-header p{font-size:12px;color:#B3D1F0;margin:4px 0 0;}

/* Cards */
.proj-card{
  background:#fff;border:1px solid #D8E6F2;border-radius:10px;
  padding:16px 20px;margin:8px 0;
  box-shadow:0 2px 8px rgba(10,60,122,.06);
}
.proj-card h4{color:#0A3C7A;font-size:15px;margin-bottom:6px;}
.val-badge{
  display:inline-block;background:#0A3C7A;color:#fff;
  border-radius:4px;padding:2px 10px;font-size:12px;font-weight:600;
}
.info-row{display:flex;gap:6px;align-items:baseline;margin:3px 0;}
.info-label{font-size:11px;color:#8AA0B8;min-width:110px;}
.info-value{font-size:12px;color:#1A2E45;font-weight:500;}

/* Generate button */
.stButton>button[kind="primary"]{
  background:linear-gradient(135deg,#0A3C7A,#2979C8)!important;
  color:#fff!important;font-size:15px!important;font-weight:700!important;
  padding:14px 28px!important;border-radius:8px!important;
  border:none!important;width:100%!important;letter-spacing:.3px!important;
}
.stButton>button[kind="primary"]:hover{
  box-shadow:0 4px 20px rgba(10,60,122,.4)!important;
}
.stDownloadButton>button{
  background:linear-gradient(135deg,#1A6A50,#2EB382)!important;
  color:#fff!important;font-size:14px!important;font-weight:600!important;
  padding:12px 24px!important;border-radius:8px!important;
  border:none!important;width:100%!important;
}

/* Form */
div[data-testid="stForm"]{
  background:#F9FBFD;border:1px solid #D8E6F2;
  border-radius:10px;padding:20px;
}
.stTextInput>label,.stTextArea>label,.stFileUploader>label{
  font-weight:600!important;color:#0A3C7A!important;font-size:13px!important;
}

/* Section dividers */
.section-title{
  font-size:18px;font-weight:700;color:#0A3C7A;
  border-bottom:2px solid #2979C8;padding-bottom:6px;margin-bottom:14px;
}

/* Tab custom */
.stTabs [data-baseweb="tab"]{font-size:14px;font-weight:600;}
.stTabs [aria-selected="true"]{color:#0A3C7A!important;}
</style>
"""


def _sidebar(co: dict) -> dict:
    """
    Sidebar: company info display + logo upload + editable fields.
    Returns possibly updated company dict.
    """
    st.sidebar.markdown("### ⚡ DKB POWER ENGINEERING")
    st.sidebar.markdown(f"**{co['name_th']}**")
    st.sidebar.markdown("---")

    info = [
        ("📋 เลขทะเบียน",      co["reg_no"]),
        ("📅 วันที่จดทะเบียน",  co["reg_date_th"]),
        ("💰 ทุนจดทะเบียน",     co["capital_th"]),
        ("📍 ที่อยู่",          co["address_th"][:60] + "…"),
        ("📞 โทรศัพท์",         co["tel"]),
        ("✉️ อีเมล",            co["email"]),
    ]
    for lbl, val in info:
        st.sidebar.markdown(f"**{lbl}**")
        st.sidebar.markdown(f"<small>{val}</small>", unsafe_allow_html=True)

    st.sidebar.markdown("---")

    # Edit company info toggle
    if st.sidebar.button("✏️  แก้ไขข้อมูลบริษัท"):
        st.session_state["edit_company"] = not st.session_state.get("edit_company", False)

    if st.session_state.get("edit_company"):
        with st.sidebar.expander("📝 ข้อมูลบริษัท", expanded=True):
            co2 = dict(co)
            co2["name_th"]    = st.text_input("ชื่อ (ไทย)",    co["name_th"])
            co2["name_en"]    = st.text_input("Name (EN)",      co["name_en"])
            co2["reg_no"]     = st.text_input("เลขทะเบียน",    co["reg_no"])
            co2["reg_date_th"]= st.text_input("วันที่จดทะเบียน (ไทย)", co["reg_date_th"])
            co2["reg_date_en"]= st.text_input("Date (EN)",     co["reg_date_en"])
            co2["capital_th"] = st.text_input("ทุน (ไทย)",     co["capital_th"])
            co2["capital_en"] = st.text_input("Capital (EN)",  co["capital_en"])
            co2["address_th"] = st.text_area("ที่อยู่ (ไทย)", co["address_th"], height=70)
            co2["address_en"] = st.text_area("Address (EN)",  co["address_en"], height=70)
            co2["tel"]        = st.text_input("โทรศัพท์",      co["tel"])
            co2["email"]      = st.text_input("อีเมล",         co["email"])
            co2["year"]       = st.text_input("ปี / Year",     co["year"])
            co2["history_th"] = st.text_area("ประวัติ (ไทย)", co["history_th"], height=120)
            co2["history_en"] = st.text_area("History (EN)",  co["history_en"], height=120)

            logo_up = st.file_uploader("โลโก้บริษัท (PNG/JPG)",
                                       type=["png","jpg","jpeg"],
                                       key="logo_upload")
            if logo_up:
                co2["logo_file"] = save_logo(logo_up)
                st.success("บันทึกโลโก้เรียบร้อย")

            if st.button("💾  บันทึกข้อมูลบริษัท", key="save_co"):
                save_company(co2)
                st.session_state["company"] = co2
                st.success("✅  บันทึกแล้ว")
                st.rerun()

            return co2

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "<small style='opacity:.5'>DKB Profile Dashboard v2.0<br>"
        f"Year {co['year']}</small>",
        unsafe_allow_html=True)
    return co


def _tab_add_project(projects: list) -> list | None:
    """Tab 1 — Add new project form."""
    st.markdown('<div class="section-title">➕ เพิ่มโครงการใหม่ / Add New Project</div>',
                unsafe_allow_html=True)

    with st.form("add_proj", clear_on_submit=True):
        st.markdown("#### 📋 รายละเอียดโครงการ / Project Details")
        c1, c2 = st.columns(2)

        with c1:
            name_th      = st.text_input("ชื่อโครงการ (ไทย) *",
                placeholder="เช่น โครงการติดตั้งระบบทีวี อาคาร A")
            customer_th  = st.text_input("ลูกค้า / ผู้ว่าจ้าง (ไทย) *",
                placeholder="เช่น สภ.บางโทรัด จ.สมุทรสาคร")
            main_contact = st.text_input("ผู้รับช่วงงาน / Main Contact",
                placeholder="เช่น บ.ณัฐกิจการช่าง")
            location     = st.text_input("สถานที่ / Location",
                placeholder="เช่น จ.สมุทรสาคร / Samut Sakhon")
            desc_th      = st.text_area("รายละเอียดงาน (ไทย) *", height=80,
                placeholder="เช่น งานติดตั้งระบบทีวีดิจิทัลและดาวเทียม")

        with c2:
            name_en      = st.text_input("Project Name (EN) *",
                placeholder="e.g. TV System Installation Project")
            customer_en  = st.text_input("Customer / Owner (EN) *",
                placeholder="e.g. Bang Thorat Police Station, Samut Sakhon")
            period       = st.text_input("ช่วงเวลา / Period",
                placeholder="เช่น 2567 หรือ Jan 2026")
            value        = st.text_input("มูลค่างาน (THB) *",
                placeholder="เช่น 35,310.00")
            desc_en      = st.text_area("Job Description (EN) *", height=80,
                placeholder="e.g. Digital TV and satellite TV system installation")

        st.markdown("#### 📸 อัพโหลดรูปภาพโครงการ / Upload Project Photos")
        uploaded = st.file_uploader(
            "ลากไฟล์มาวางหรือคลิกเลือก  ·  รองรับ JPG PNG WEBP",
            accept_multiple_files=True,
            type=["jpg","jpeg","png","webp"],
            label_visibility="collapsed",
            key="photo_upload"
        )

        # Preview strip
        if uploaded:
            st.markdown(f"**เลือก {len(uploaded)} ไฟล์**")
            tcols = st.columns(min(len(uploaded), 4))
            for i, uf in enumerate(uploaded[:4]):
                with tcols[i]:
                    st.image(Image.open(uf), use_container_width=True,
                             caption=uf.name[:18])
                    uf.seek(0)
            if len(uploaded) > 4:
                st.caption(f"+ {len(uploaded)-4} ไฟล์เพิ่มเติม")

        submitted = st.form_submit_button(
            "💾  บันทึกโครงการ / Save Project",
            use_container_width=True)

    if submitted:
        errs = []
        if not name_th.strip():    errs.append("ชื่อโครงการ (ไทย)")
        if not name_en.strip():    errs.append("Project Name (EN)")
        if not customer_th.strip():errs.append("ลูกค้า (ไทย)")
        if not customer_en.strip():errs.append("Customer (EN)")
        if not desc_th.strip():    errs.append("รายละเอียดงาน (ไทย)")
        if not desc_en.strip():    errs.append("Job Description (EN)")
        if not value.strip():      errs.append("มูลค่างาน")
        if errs:
            st.error(f"⚠️  กรุณากรอก: {', '.join(errs)}")
            return None

        photos  = save_photos(uploaded)
        new     = {
            "id"         : f"proj_{uuid.uuid4().hex[:8]}",
            "no"         : next_no(projects),
            "name_th"    : name_th.strip(),
            "name_en"    : name_en.strip(),
            "customer_th": customer_th.strip(),
            "customer_en": customer_en.strip(),
            "main_contact": main_contact.strip() or "—",
            "location"   : location.strip(),
            "desc_th"    : desc_th.strip(),
            "desc_en"    : desc_en.strip(),
            "period"     : period.strip(),
            "value"      : value.strip(),
            "photos"     : photos,
            "created_at" : datetime.now().isoformat(),
        }
        updated = projects + [new]
        save_projects(updated)
        st.success(f"✅  บันทึกโครงการที่ {new['no']}: {new['name_th']}")
        if photos: st.info(f"📸  บันทึก {len(photos)} รูปภาพ")
        st.balloons()
        return updated
    return None


def _tab_view_projects(projects: list) -> list | None:
    """Tab 2 — View and manage all projects."""
    st.markdown(
        f'<div class="section-title">📋 โครงการทั้งหมด / All Projects '
        f'<span style="font-size:14px;color:#2979C8;">({len(projects)} โครงการ)</span></div>',
        unsafe_allow_html=True)

    if not projects:
        st.info("ยังไม่มีโครงการ — เพิ่มที่แท็บแรก")
        return None

    for proj in projects:
        with st.expander(
            f"  #{proj.get('no','—')}  ·  {proj.get('name_th','(ไม่มีชื่อ)')}",
            expanded=False):

            col_info, col_photo = st.columns([3, 1])
            with col_info:
                rows = [
                    ("ชื่อ EN",          proj.get("name_en","—")),
                    ("ลูกค้า",           f"{proj.get('customer_th','—')}  ·  {proj.get('customer_en','—')}"),
                    ("ผู้รับช่วงงาน",    proj.get("main_contact","—")),
                    ("สถานที่",          proj.get("location","—")),
                    ("รายละเอียด",       f"{proj.get('desc_th','—')}  /  {proj.get('desc_en','—')}"),
                    ("ช่วงเวลา",         proj.get("period","—")),
                ]
                for lbl, val in rows:
                    st.markdown(
                        f'<div class="info-row">'
                        f'<span class="info-label">{lbl}</span>'
                        f'<span class="info-value">{val}</span></div>',
                        unsafe_allow_html=True)
                st.markdown(
                    f'มูลค่างาน: <span class="val-badge">฿ {proj.get("value","—")} บาท</span>',
                    unsafe_allow_html=True)

            with col_photo:
                photos = proj.get("photos", [])
                st.caption(f"📸 {len(photos)} รูป")
                if photos:
                    fp = PHOTOS_DIR / photos[0]
                    if fp.exists():
                        st.image(str(fp), use_container_width=True)

            # Photo strip
            if len(photos) > 1:
                tc = st.columns(min(len(photos), 4))
                for k, fn in enumerate(photos[:8]):
                    fp2 = PHOTOS_DIR / fn
                    if fp2.exists():
                        with tc[k % 4]:
                            st.image(str(fp2), use_container_width=True)

            # Add more photos to existing project
            with st.expander("➕ เพิ่มรูปภาพเพิ่มเติม"):
                add_key = f"add_ph_{proj['id']}"
                more_ph = st.file_uploader(
                    "เลือกรูปเพิ่ม",
                    accept_multiple_files=True,
                    type=["jpg","jpeg","png","webp"],
                    key=add_key)
                if st.button("บันทึกรูปเพิ่ม", key=f"saveph_{proj['id']}"):
                    if more_ph:
                        new_fns = save_photos(more_ph)
                        for p2 in projects:
                            if p2["id"] == proj["id"]:
                                p2["photos"] = p2.get("photos",[]) + new_fns
                        save_projects(projects)
                        st.success(f"เพิ่ม {len(new_fns)} รูปแล้ว")
                        st.rerun()

            # Delete
            if st.button(f"🗑️  ลบโครงการนี้", key=f"del_{proj['id']}", type="secondary"):
                remaining = [p for p in projects if p["id"] != proj["id"]]
                for k, p in enumerate(remaining): p["no"] = k + 1
                save_projects(remaining)
                st.warning(f"ลบโครงการ '{proj['name_th']}' แล้ว")
                st.session_state.projects = remaining
                st.rerun()

    return None


def _tab_company(co: dict) -> dict | None:
    """Tab 3 — Edit company information in full."""
    st.markdown('<div class="section-title">🏢 ข้อมูลบริษัท / Company Information</div>',
                unsafe_allow_html=True)

    with st.form("edit_company_full"):
        st.markdown("#### ชื่อและทะเบียน / Identity")
        c1, c2 = st.columns(2)
        with c1:
            name_th    = st.text_input("ชื่อบริษัท (ไทย)", co["name_th"])
            reg_no     = st.text_input("เลขทะเบียนนิติบุคคล", co["reg_no"])
            reg_th     = st.text_input("วันที่จดทะเบียน (ไทย)", co["reg_date_th"])
            capital_th = st.text_input("ทุนจดทะเบียน (ไทย)", co["capital_th"])
        with c2:
            name_en    = st.text_input("Company Name (EN)", co["name_en"])
            year       = st.text_input("ปี / Year", co["year"])
            reg_en     = st.text_input("Date of Incorporation (EN)", co["reg_date_en"])
            capital_en = st.text_input("Registered Capital (EN)", co["capital_en"])

        st.markdown("#### ที่อยู่และติดต่อ / Address & Contact")
        c3, c4 = st.columns(2)
        with c3:
            addr_th = st.text_area("ที่อยู่ (ไทย)", co["address_th"], height=70)
            tel     = st.text_input("โทรศัพท์", co["tel"])
        with c4:
            addr_en = st.text_area("Address (EN)", co["address_en"], height=70)
            email   = st.text_input("อีเมล", co["email"])

        st.markdown("#### ประวัติ / History & Mission")
        hist_th = st.text_area("ประวัติบริษัท (ไทย)", co["history_th"], height=120)
        hist_en = st.text_area("Company History (EN)", co["history_en"], height=120)

        st.markdown("#### โลโก้บริษัท / Company Logo")
        logo_file = st.file_uploader(
            "อัพโหลดโลโก้ (PNG/JPG, แนะนำขนาด 300×100 px)",
            type=["png","jpg","jpeg"], key="logo_full")
        if co.get("logo_file"):
            lp = ASSETS_DIR / co["logo_file"]
            if lp.exists():
                st.image(str(lp), width=180, caption="โลโก้ปัจจุบัน")

        saved = st.form_submit_button("💾  บันทึกข้อมูลบริษัท", use_container_width=True)

    if saved:
        logo_fn = co.get("logo_file","")
        if logo_file:
            logo_fn = save_logo(logo_file)
        updated = {**co,
            "name_th": name_th, "name_en": name_en,
            "reg_no": reg_no, "reg_date_th": reg_th, "reg_date_en": reg_en,
            "capital_th": capital_th, "capital_en": capital_en,
            "address_th": addr_th, "address_en": addr_en,
            "tel": tel, "email": email, "year": year,
            "history_th": hist_th, "history_en": hist_en,
            "logo_file": logo_fn,
        }
        save_company(updated)
        st.success("✅  บันทึกข้อมูลบริษัทสำเร็จ")
        return updated
    return None


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="DKB Power Engineering — Profile Dashboard",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CSS, unsafe_allow_html=True)

    # ── State init ────────────────────────────────────────────────────────────
    if "company"   not in st.session_state:
        st.session_state.company  = load_company()
    if "projects"  not in st.session_state:
        st.session_state.projects = load_projects()
    if "docx_bytes" not in st.session_state:
        st.session_state.docx_bytes = None

    co       = st.session_state.company
    projects = st.session_state.projects

    # ── Sidebar ───────────────────────────────────────────────────────────────
    _sidebar(co)

    # ── App header ────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="app-header">
      <div class="tag">Company Profile Dashboard  ·  v2.0</div>
      <h1>⚡ DKB POWER ENGINEERING CO., LTD.</h1>
      <p>บริษัท ดีเคบี เพาเวอร์ เอนจิเนียริ่ง จำกัด  ·  ระบบสร้างเอกสารโปรไฟล์บริษัทอัตโนมัติ</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "➕  เพิ่มโครงการใหม่",
        "📋  ดูโครงการทั้งหมด",
        "🏢  ข้อมูลบริษัท",
    ])

    with tab1:
        result = _tab_add_project(projects)
        if result is not None:
            st.session_state.projects = result
            projects = result

    with tab2:
        _tab_view_projects(projects)

    with tab3:
        result3 = _tab_company(co)
        if result3 is not None:
            st.session_state.company = result3
            co = result3
            st.rerun()

    # ── Generate section ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        f'<div style="text-align:center;font-size:16px;font-weight:700;'
        f'color:#0A3C7A;margin-bottom:10px;">'
        f"📄  พร้อมสร้างเอกสาร — {len(projects)} โครงการ</div>",
        unsafe_allow_html=True)

    gc, dc = st.columns([1, 1])
    with gc:
        if st.button("⚙️  สร้าง Company Profile (.docx)",
                     use_container_width=True, type="primary"):
            with st.spinner("กำลังสร้างเอกสาร…"):
                try:
                    raw = generate_docx(co, projects)
                    st.session_state.docx_bytes = raw
                    st.success("✅  สร้างเอกสารสำเร็จ!")
                except Exception as exc:
                    st.error(f"❌  ข้อผิดพลาด: {exc}")
                    raise

    with dc:
        if st.session_state.docx_bytes:
            fname = f"DKB_Power_Engineering_Company_Profile_{co['year']}.docx"
            st.download_button(
                label="⬇️  Download Company Profile (.docx)",
                data=st.session_state.docx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document",
                use_container_width=True,
            )

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown(
        f"<div style='text-align:center;margin-top:32px;font-size:11px;"
        f"color:#AAB8C8;'>DKB Power Engineering Co., Ltd.  ·  "
        f"Company Profile Dashboard v2.0  ·  {co['year']}</div>",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()

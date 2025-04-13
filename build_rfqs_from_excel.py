"""
build_rfqs_from_excel.py  –  final with folder export
====================================================
Generates one RFQ per equipment row and saves it to:
out/<Item‑name>/RFQ/RFQ_<cid>_<Item‑name>.docx
"""

from __future__ import annotations
import os, json, pathlib, textwrap
from typing import List
import re

import pandas as pd
from tqdm import tqdm
from dotenv import load_dotenv
import openai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls

# ── ENV / KEY ───────────────────────────────────────────────
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY") or ""
if not openai.api_key:
    raise RuntimeError("OPENAI_API_KEY not set")

# ── CONFIG ─────────────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_ROOT     = pathlib.Path("C:/Users/PRANAY-RES/OneDrive - Renewable Energy Systems Limited/RES/Capital Equipment 2025")           # top‑level output dir
MODEL        = "gpt-4o-mini"                # or gpt-4o
TEMPERATURE  = 0.2
OUT_ROOT.mkdir(exist_ok=True)

# ── 1. Load Excel ──────────────────────────────────────────
raw = pd.read_excel(EXCEL_PATH, header=2).iloc[1:]
raw = raw.rename(columns={
    "Unnamed: 1": "cid",
    "Unnamed: 2": "Item",
    "Unnamed: 3": "Priority",
    "Unnamed: 4": "Assignee",
    "Unnamed: 5": "Description",
})
df = (
    raw[["cid", "Item", "Priority", "Assignee", "Description"]]
    .dropna(subset=["Item"])
    .reset_index(drop=True)
)

# ── 2. GPT prompts (intro, scope, tech, docs) ──────────────
SYSTEM_PROMPT = textwrap.dedent("""\
You are a senior procurement engineer at Renewable Energy Systems Limited
(AS9100D, lithium‑thermal battery manufacturer). Draft concise RFQ sections
aligned with MIL‑STD‑810H, MIL‑STD‑1580, ASTM, IEC, etc.

Return **valid JSON only** with keys:
  introduction, scope, tech_table, docs_required

• tech_table must contain **at least 8** relevant {parameter, requirement} pairs.
""")

USER_TMPL = (
    "Draft the RFQ sections for equipment '{item}'. "
    "{desc_clause}"
    "Return JSON only."
)

def gpt_sections(item: str, desc: str | None) -> dict:
    clause = f"Include these user‑specified criteria: {desc}. " if desc else ""
    prompt = USER_TMPL.format(item=item, desc_clause=clause)

    resp = openai.chat.completions.create(
        model=MODEL,
        temperature=TEMPERATURE,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
    )
    return json.loads(resp.choices[0].message.content)

# ── 3. DOCX helpers ────────────────────────────────────────
def style_body(paragraph):
    for run in paragraph.runs:
        run.font.bold = False
        run.font.underline = False
        run.font.size = Pt(11)

def bullet(doc: Document, entry, bullet_ok: bool):
    if isinstance(entry, (list, tuple)):
        for part in entry:
            bullet(doc, part, bullet_ok)
        return
    p = (
        doc.add_paragraph(entry, style="List Bullet")
        if bullet_ok
        else doc.add_paragraph(f"• {entry}")
    )
    style_body(p)

def box(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcPr.append(
        parse_xml(
            r'<w:tcBorders %s>'
            r'<w:top w:val="single" w:sz="4" w:color="000000"/>'
            r'<w:left w:val="single" w:sz="4" w:color="000000"/>'
            r'<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
            r'<w:right w:val="single" w:sz="4" w:color="000000"/>'
            r'</w:tcBorders>' % nsdecls("w")
        )
    )

# Hard‑coded Commercial Requirements rows
COMM_ROWS = [
    ("Quotation Validity",  "Minimum 90 days"),
    ("Delivery Time",       "Specify lead time"),
    ("Pricing Terms",       "Provide EXW, FOB, and CIF pricing (as applicable)"),
    ("Payment Terms",       "To be negotiated"),
    ("Installation and Training",
                             "Specify installation and operator training charges if applicable"),
    ("After‑Sales Support",
                             "Provide details of service support, spares availability, and annual maintenance contracts"),
]

# ── 4. Build DOCX ──────────────────────────────────────────
def build_doc(sec: dict, item: str, cid: int, bullet_ok: bool) -> pathlib.Path:
    doc = Document(TEMPLATE_HDR)

    # Remove leading blank paragraph if template adds one
    if doc.paragraphs and not doc.paragraphs[0].text.strip():
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)

    # Title
    title = doc.add_paragraph(f"RFQ for {item}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size = Pt(16)
    title.style.font.bold = True
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(0)

    def H(num, txt):
        h = doc.add_heading(f"{num}. {txt}", level=2)
        for r in h.runs:
            r.font.size = Pt(14); r.font.bold = True

    # 1 Introduction & 2 Scope
    H(1, "Introduction"); style_body(doc.add_paragraph(sec["introduction"]))
    H(2, "Scope of Supply"); style_body(doc.add_paragraph(sec["scope"]))

    # 3 Technical Requirements
    H(3, "Technical Requirements")
    tbl = doc.add_table(rows=1, cols=2); tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl.rows[0].cells
    hdr[0].text = "Parameter"; hdr[1].text = "Requirement"
    for c in hdr:
        r=c.paragraphs[0].runs[0]; r.font.bold=True; r.font.size=Pt(12); box(c)
    for row in sec["tech_table"]:
        cells = tbl.add_row().cells
        cells[0].text, cells[1].text = row["parameter"], row["requirement"]
        for c in cells: style_body(c.paragraphs[0]); box(c)

    # 4 Commercial Requirements (fixed table)
    H(4, "Commercial Requirements")
    ct = doc.add_table(rows=1, cols=2); ct.alignment = WD_TABLE_ALIGNMENT.CENTER
    ch = ct.rows[0].cells
    ch[0].text, ch[1].text = "Parameter", "Requirement"
    for c in ch:
        r=c.paragraphs[0].runs[0]; r.font.bold=True; r.font.size=Pt(12); box(c)
    for p, req in COMM_ROWS:
        cells = ct.add_row().cells
        cells[0].text, cells[1].text = p, req
        for c in cells: style_body(c.paragraphs[0]); box(c)

    # 5 Documentation Requirements
    H(5, "Documentation Requirements")
    bullet(doc, sec["docs_required"], bullet_ok)

    # 6 Submission Guidelines
    H(6, "Submission Guidelines")
    style_body(doc.add_paragraph(
        f"Please submit your quotations via email with the subject line:\n"
        f"“Quotation for {item} - RESL”."))
    bullet(doc, "Contact Person: P. Pranay Kiran, A. Sai Nithin", bullet_ok)
    bullet(doc, "Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in", bullet_ok)

    # 7 Confidentiality Clause
    H(7, "Confidentiality Clause")
    style_body(doc.add_paragraph(
        "All quotations and related documents submitted in response to this RFQ "
        "will be treated as confidential and used solely for evaluation purposes."))

    # ── Export path logic ────────────────────────────────
    safe_name = re.sub(r"[^A-Za-z0-9_]+", "_", item.strip())
    rfq_dir   = OUT_ROOT / safe_name / "RFQ"
    rfq_dir.mkdir(parents=True, exist_ok=True)

    out_path = rfq_dir / f"RFQ_{cid:02d}_{safe_name}.docx"
    doc.save(out_path)
    return out_path

# Bullet style availability
try:
    Document(TEMPLATE_HDR).styles["List Bullet"]; BULLET_OK=True
except KeyError:
    BULLET_OK=False; print("ℹ️  'List Bullet' style not found – plain bullets used.")

# ── 5. Main loop ───────────────────────────────────────────
created: List[pathlib.Path] = []
for rec in tqdm(df.to_dict(orient="records"), desc="Generating RFQs"):
    try: sec = gpt_sections(rec["Item"], rec.get("Description"))
    except Exception as e:
        print(f"\n⚠️  GPT error on '{rec['Item']}': {e}"); continue
    if not sec:
        print(f"\n⚠️  '{rec['Item']}' skipped: GPT did not return valid JSON."); continue
    try: created.append(build_doc(sec, rec["Item"], int(rec["cid"]), BULLET_OK))
    except Exception as e:
        print(f"\n⚠️  DOCX error on '{rec['Item']}': {e}")

print(f"\n✓ Created {len(created)} RFQ file(s) under {OUT_ROOT.resolve()}")

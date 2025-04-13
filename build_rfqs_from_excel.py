"""
build_rfqs_from_excel.py  –  stable release
===========================================
Generate one RFQ DOCX per row in data/data_cp_1.xlsx using templates/u1.docx.

Features
--------
• New OpenAI SDK (>=1.0) with response_format="json_object"
• Optional Description column included in GPT prompt
• Table header bold, body normal, 0.5 pt black borders on every cell
• Removes any existing <w:tcBorders> before adding new ⇒ no '{http' crash
• Bullet fallback if 'List Bullet' style absent
"""

from __future__ import annotations
import os, json, pathlib, textwrap
from typing import List

import pandas as pd
from tqdm import tqdm
from dotenv import load_dotenv
import openai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ── 0. Environment / API key ─────────────────────────────────
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")
if not openai.api_key:
    raise RuntimeError("OPENAI_API_KEY not set (env var or .env file)")

# ── 1. Config paths ─────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_DIR      = pathlib.Path("out/rfq_docs")
MODEL        = "gpt-4o-mini"      # or "gpt-4o"
TEMPERATURE  = 0.2
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 2. Load Excel (cid, Item, Priority, Assignee, Description) ─
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

# ── 3. GPT prompt templates ─────────────────────────────────
SYSTEM_PROMPT = textwrap.dedent(
    """\
    You are a senior procurement engineer at Renewable Energy Systems Limited
    (AS9100D‑compliant, lithium‑thermal battery manufacturer).
    Draft concise RFQ sections aligned with MIL‑STD‑810H, MIL‑STD‑1580,
    ASTM, IEC, or other relevant standards.

    Return **valid JSON only** with keys:
      introduction       (string)
      scope              (string)
      tech_table         (list of {parameter, requirement})
      commercial_terms   (list of strings)
      docs_required      (list of strings)
    """
)

USER_TMPL = (
    "Draft the RFQ sections for capital equipment '{item}'. "
    "{desc_clause}"
    "Use lowercase keys 'parameter' and 'requirement' in tech_table. "
    "Return JSON only."
)

def gpt_sections(item: str, desc: str | None) -> dict | None:
    desc_clause = f"Include these user‑specified criteria: {desc}. " if desc else ""
    user_prompt = USER_TMPL.format(item=item, desc_clause=desc_clause)

    resp = openai.chat.completions.create(
        model=MODEL,
        temperature=TEMPERATURE,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    )
    return json.loads(resp.choices[0].message.content)

# ── 4. DOCX helper functions ────────────────────────────────
def unbold(paragraph):
    for run in paragraph.runs:
        run.font.bold = False

def add_bullet(doc: Document, text: str, bullet_ok: bool):
    p = (
        doc.add_paragraph(text, style="List Bullet")
        if bullet_ok
        else doc.add_paragraph(f"• {text}")
    )
    unbold(p)

def box(cell):
    """
    Give *cell* a thin black border on all four edges.
    Removes any existing <w:tcBorders> first to avoid duplicate-tag errors.
    """
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:tcBorders")):
        tc_pr.remove(old)

    borders = parse_xml(
        r'<w:tcBorders %s>'
        r'<w:top    w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left   w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:right  w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tcBorders>' % nsdecls("w")
    )
    tc_pr.append(borders)

def build_doc(sec: dict, item: str, cid: int, bullet_ok: bool) -> pathlib.Path:
    doc = Document(TEMPLATE_HDR)

    # Title
    title = doc.add_paragraph(f"RFQ for {item}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size = Pt(16)
    title.style.font.bold = True
    doc.add_paragraph()

    def H(num, txt): doc.add_heading(f"{num}. {txt}", level=2)

    # 1. Introduction
    H(1, "Introduction")
    unbold(doc.add_paragraph(sec["introduction"]))

    # 2. Scope of Supply
    H(2, "Scope of Supply")
    unbold(doc.add_paragraph(sec["scope"]))

    # 3. Technical Requirements
    H(3, "Technical Requirements")
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Parameter"
    hdr_cells[1].text = "Requirement"
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].font.bold = True
        box(cell)

    for row in sec["tech_table"]:
        cells = table.add_row().cells
        cells[0].text = row["parameter"]
        cells[1].text = row["requirement"]
        for c in cells:
            unbold(c.paragraphs[0])
            box(c)

    # 4. Commercial Requirements
    H(4, "Commercial Requirements")
    for bullet in sec["commercial_terms"]:
        add_bullet(doc, bullet, bullet_ok)

    # 5. Documentation Requirements
    H(5, "Documentation Requirements")
    for bullet in sec["docs_required"]:
        add_bullet(doc, bullet, bullet_ok)

    # 6. Submission Guidelines
    H(6, "Submission Guidelines")
    unbold(
        doc.add_paragraph(
            f"Please submit your quotations via email with the subject line:\n"
            f"“Quotation for {item} - RESL”."
        )
    )
    add_bullet(doc, "Contact Person: P. Pranay Kiran, A. Sai Nithin", bullet_ok)
    add_bullet(
        doc,
        "Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in",
        bullet_ok,
    )

    # 7. Confidentiality Clause
    H(7, "Confidentiality Clause")
    unbold(
        doc.add_paragraph(
            "All quotations and related documents submitted in response to this RFQ "
            "will be treated as confidential and used solely for evaluation purposes."
        )
    )

    out_path = OUT_DIR / f"RFQ_{cid:02d}_{item.replace(' ', '_')}.docx"
    doc.save(out_path)
    return out_path

# Bullet style availability
try:
    Document(TEMPLATE_HDR).styles["List Bullet"]
    BULLET_OK = True
except KeyError:
    BULLET_OK = False
    print("ℹ️  'List Bullet' style not found – plain bullets will be used.")

# ── 5. Main loop ───────────────────────────────────────────
created: List[pathlib.Path] = []

for rec in tqdm(df.to_dict(orient="records"), desc="Generating RFQs"):
    try:
        sec = gpt_sections(rec["Item"], rec.get("Description"))
    except Exception as e:
        print(f"\n⚠️  GPT error on '{rec['Item']}': {e}")
        continue

    if not sec:
        print(f"\n⚠️  '{rec['Item']}' skipped: GPT did not return valid JSON.")
        continue

    try:
        created.append(build_doc(sec, rec["Item"], int(rec["cid"]), BULLET_OK))
    except Exception as e:
        print(f"\n⚠️  DOCX error on '{rec['Item']}': {e}")

print(f"\n✓ Created {len(created)} RFQ file(s) in {OUT_DIR.resolve()}")

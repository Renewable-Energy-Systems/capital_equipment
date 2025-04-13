"""
build_rfqs_from_excel.py – bullet‑fix final
===========================================
• New OpenAI SDK (>=1.0) with response_format=json_object
• Description column supported
• Headings 14 pt bold, table header 12 pt bold, body 11 pt regular
• Bullet helper treats a string as ONE bullet (no alphabet soup)
• Table fully boxed (0.5 pt black), no duplicate-border crash
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
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls

# ── ENV / API KEY ───────────────────────────────────────────
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY") or ""
if not openai.api_key:
    raise RuntimeError("OPENAI_API_KEY not set")

# ── CONFIG ─────────────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_DIR      = pathlib.Path("out/rfq_docs")
MODEL        = "gpt-4o-mini"      # or "gpt-4o"
TEMPERATURE  = 0.2
OUT_DIR.mkdir(parents=True, exist_ok=True)

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

# ── 2. GPT prompts ─────────────────────────────────────────
SYSTEM_PROMPT = textwrap.dedent(
    """\
    You are a senior procurement engineer at Renewable Energy Systems Limited
    (AS9100D, lithium‑thermal battery manufacturer). Draft concise RFQ sections
    aligned with MIL‑STD‑810H, MIL‑STD‑1580, ASTM, IEC, etc.

    Return **valid JSON only** with keys:
      introduction, scope, tech_table, commercial_terms, docs_required

    • tech_table must contain **at least 8** relevant {parameter, requirement} pairs.
    """
)

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
    """
    Add ONE bullet line.  If *entry* is a list/tuple, recurse for each item.
    """
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
    """Add 0.5 pt black border to all sides of *cell* (safe)."""
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

def build_doc(sec: dict, item: str, cid: int, bullet_ok: bool) -> pathlib.Path:
    doc = Document(TEMPLATE_HDR)

    # Title
    title = doc.add_paragraph(f"RFQ for {item}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size = Pt(16)
    title.style.font.bold = True
    doc.add_paragraph()

    def H(num, txt):
        h = doc.add_heading(f"{num}. {txt}", level=2)
        for r in h.runs:
            r.font.size = Pt(14)
            r.font.bold = True

    # 1. Introduction & 2. Scope
    H(1, "Introduction")
    style_body(doc.add_paragraph(sec["introduction"]))
    H(2, "Scope of Supply")
    style_body(doc.add_paragraph(sec["scope"]))

    # 3. Technical Requirements
    H(3, "Technical Requirements")
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text = "Parameter", "Requirement"
    for c in hdr:
        r = c.paragraphs[0].runs[0]
        r.font.bold = True
        r.font.size = Pt(12)
        box(c)

    for row in sec["tech_table"]:
        cells = table.add_row().cells
        cells[0].text, cells[1].text = row["parameter"], row["requirement"]
        for c in cells:
            style_body(c.paragraphs[0])
            box(c)

    # 4. Commercial Requirements
    H(4, "Commercial Requirements")
    bullet(doc, sec["commercial_terms"], bullet_ok)

    # 5. Documentation Requirements
    H(5, "Documentation Requirements")
    bullet(doc, sec["docs_required"], bullet_ok)

    # 6. Submission Guidelines
    H(6, "Submission Guidelines")
    p = doc.add_paragraph(
        f"Please submit your quotations via email with the subject line:\n"
        f"“Quotation for {item} - RESL”."
    )
    style_body(p)
    bullet(doc, "Contact Person: P. Pranay Kiran, A. Sai Nithin", bullet_ok)
    bullet(
        doc,
        "Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in",
        bullet_ok,
    )

    # 7. Confidentiality Clause
    H(7, "Confidentiality Clause")
    style_body(
        doc.add_paragraph(
            "All quotations and related documents submitted in response to this RFQ "
            "will be treated as confidential and used solely for evaluation purposes."
        )
    )

    out = OUT_DIR / f"RFQ_{cid:02d}_{item.replace(' ', '_')}.docx"
    doc.save(out)
    return out

# Bullet style availability
try:
    Document(TEMPLATE_HDR).styles["List Bullet"]
    BULLET_OK = True
except KeyError:
    BULLET_OK = False
    print("ℹ️  'List Bullet' style not found – plain bullets will be used.")

# ── 4. Main loop ───────────────────────────────────────────
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

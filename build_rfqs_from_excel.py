"""
build_rfqs_from_excel.py – formatting & richer‑spec release
===========================================================
• Removes underlines in table body
• Distinct font sizes: headings 14 pt, table header 12 pt, body 11 pt
• GPT asked for ≥8 tech parameters
• Safe border handling (no '{http' crash)
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

# ── ENV / KEY ───────────────────────────────────────────────
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY") or ""
if not openai.api_key:
    raise RuntimeError("OPENAI_API_KEY not set")

# ── CONFIG ─────────────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_DIR      = pathlib.Path("out/rfq_docs")
MODEL        = "gpt-4o-mini"     # or "gpt-4o"
TEMPERATURE  = 0.2
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 1. Load Excel ──────────────────────────────────────────
raw = pd.read_excel(EXCEL_PATH, header=2).iloc[1:]
raw = raw.rename(columns={
    "Unnamed: 1":"cid","Unnamed: 2":"Item","Unnamed: 3":"Priority",
    "Unnamed: 4":"Assignee","Unnamed: 5":"Description"})
df = (raw[["cid","Item","Priority","Assignee","Description"]]
      .dropna(subset=["Item"]).reset_index(drop=True))

# ── 2. GPT prompts ─────────────────────────────────────────
SYSTEM = textwrap.dedent("""\
You are a senior procurement engineer at Renewable Energy Systems Limited
(lithium‑thermal battery manufacturer, AS9100D). Draft concise RFQ sections
aligned to MIL‑STD‑810H, MIL‑STD‑1580, ASTM, IEC, etc.

Return **valid JSON only**:
  introduction, scope, tech_table, commercial_terms, docs_required
tech_table must contain **at least 8** {parameter, requirement} pairs that are
truly relevant to the equipment.
""")

USER_TMPL = ("Draft the RFQ sections for equipment '{item}'. "
             "{desc_clause}"
             "Return JSON only.")

def gpt_sections(item:str, desc:str|None)->dict:
    clause = f"Include these user‑specified criteria: {desc}. " if desc else ""
    prompt = USER_TMPL.format(item=item, desc_clause=clause)
    resp = openai.chat.completions.create(
        model=MODEL, temperature=TEMPERATURE,
        response_format={"type":"json_object"},
        messages=[{"role":"system","content":SYSTEM},
                  {"role":"user","content":prompt}]
    )
    return json.loads(resp.choices[0].message.content)

# ── 3. DOCX helpers ────────────────────────────────────────
def unbold_underline(par):
    for run in par.runs:
        run.font.bold      = False
        run.font.underline = False

def bullet(doc,text,ok):
    p = doc.add_paragraph(text,style="List Bullet") if ok else doc.add_paragraph(f"• {text}")
    unbold_underline(p)

def box(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")): tcPr.remove(old)
    tcPr.append(parse_xml(
        r'<w:tcBorders %s>'
        r'<w:top w:val="single" w:sz="4" w:color="000000"/>'
        r'<w:left w:val="single" w:sz="4" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
        r'<w:right w:val="single" w:sz="4" w:color="000000"/>'
        r'</w:tcBorders>' % nsdecls('w')))

def build_doc(sec:dict,item:str,cid:int,bullet_ok:bool)->pathlib.Path:
    doc=Document(TEMPLATE_HDR)
    # Title
    title=doc.add_paragraph(f"RFQ for {item}")
    title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size=Pt(16); title.style.font.bold=True
    doc.add_paragraph()

    def H(n,t):
        h=doc.add_heading(f"{n}. {t}",level=2)
        for r in h.runs: r.font.size=Pt(14); r.font.bold=True

    # 1 Intro & 2 Scope
    H(1,"Introduction"); unbold_underline(doc.add_paragraph(sec["introduction"]))
    H(2,"Scope of Supply"); unbold_underline(doc.add_paragraph(sec["scope"]))

    # 3 Technical Requirements
    H(3,"Technical Requirements")
    tbl=doc.add_table(rows=1,cols=2); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    hdr=tbl.rows[0].cells
    hdr[0].text="Parameter"; hdr[1].text="Requirement"
    for c in hdr:
        r=c.paragraphs[0].runs[0]
        r.font.bold=True; r.font.size=Pt(12)
        box(c)
    for row in sec["tech_table"]:
        cells=tbl.add_row().cells
        cells[0].text=row["parameter"]; cells[1].text=row["requirement"]
        for c in cells:
            r=c.paragraphs[0].runs[0]
            r.font.size=Pt(11); r.font.bold=False; r.font.underline=False
            box(c)

    # 4 Commercial & 5 Docs
    H(4,"Commercial Requirements")
    for b in sec["commercial_terms"]: bullet(doc,b,bullet_ok)
    H(5,"Documentation Requirements")
    for b in sec["docs_required"]:   bullet(doc,b,bullet_ok)

    # 6 Submission
    H(6,"Submission Guidelines")
    unbold_underline(doc.add_paragraph(
        f"Please submit your quotations via email with the subject line:\n"
        f"“Quotation for {item} - RESL”."))
    bullet(doc,"Contact Person: P. Pranay Kiran, A. Sai Nithin",bullet_ok)
    bullet(doc,"Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in",bullet_ok)

    # 7 Confidentiality
    H(7,"Confidentiality Clause")
    unbold_underline(doc.add_paragraph(
        "All quotations and related documents submitted in response to this RFQ "
        "will be treated as confidential and used solely for evaluation purposes."))

    out=OUT_DIR/f"RFQ_{cid:02d}_{item.replace(' ','_')}.docx"
    doc.save(out); return out

# Bullet style availability
try: Document(TEMPLATE_HDR).styles["List Bullet"]; BULLET_OK=True
except KeyError: BULLET_OK=False; print("ℹ️  'List Bullet' style not found – plain bullets used.")

# ── 4. Main loop ───────────────────────────────────────────
created: List[pathlib.Path]=[]
for rec in tqdm(df.to_dict(orient="records"),desc="Generating RFQs"):
    try: sec=gpt_sections(rec["Item"], rec.get("Description"))
    except Exception as e:
        print(f"\n⚠️  GPT error on '{rec['Item']}': {e}"); continue
    if not sec:
        print(f"\n⚠️  '{rec['Item']}' skipped: no JSON"); continue
    try: created.append(build_doc(sec, rec["Item"], int(rec["cid"]), BULLET_OK))
    except Exception as e:
        print(f"\n⚠️  DOCX error on '{rec['Item']}': {e}")

print(f"\n✓ Created {len(created)} RFQ file(s) in {OUT_DIR.resolve()}")

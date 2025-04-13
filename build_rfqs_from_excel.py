
"""
build_rfqs_from_excel.py – v2.1
Fixes: removes unintended bold formatting from body text & bullet points.
"""

from __future__ import annotations
import os, json, textwrap, pathlib
from typing import Tuple, List

import pandas as pd
from tqdm import tqdm
from dotenv import load_dotenv
import openai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

# ── ENV / API KEY ────────────────────────────────────────────
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")
if not openai.api_key:
    raise RuntimeError("OPENAI_API_KEY not set")

# ── CONFIG ───────────────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_DIR      = pathlib.Path("out/rfq_docs")
MODEL        = "gpt-4o-mini"
TEMPERATURE  = 0.2

OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 1. LOAD LIST ─────────────────────────────────────────────
df = (
    pd.read_excel(EXCEL_PATH, header=2)
      .iloc[1:]
      .rename(columns={
          "Unnamed: 1": "cid",
          "Unnamed: 2": "Item",
          "Unnamed: 3": "Priority",
          "Unnamed: 4": "Assignee",
      })
      .loc[:, ["cid", "Item", "Priority", "Assignee"]]
      .dropna(subset=["Item"])
      .reset_index(drop=True)
)

# ── 2. GPT PROMPTS ───────────────────────────────────────────
SYSTEM_PROMPT = textwrap.dedent("""You are a senior procurement engineer at Renewable Energy Systems Limited
(defence lithium‑thermal battery manufacturer). Draft RFQ sections with
practical, measurable specifications aligned to MIL‑STD, ASTM, IEC where
relevant. Output JSON with keys:
  introduction, scope, tech_table(list[{parameter,requirement}]),
  commercial_terms(list[str]), docs_required(list[str]).
""")

def draft_sections(name:str)->dict:
    user = (f"Draft RFQ sections for capital equipment '{name}'. "
            "Use lowercase keys parameter & requirement inside tech_table.")
    resp = openai.chat.completions.create(
        model=MODEL, temperature=TEMPERATURE,
        messages=[{"role":"system","content":SYSTEM_PROMPT},
                  {"role":"user","content":user}],
        response_format={"type":"json_object"}
    )
    return json.loads(resp.choices[0].message.content)

# ── 3. WORD HELPERS ─────────────────────────────────────────-
def normalise(row:dict)->Tuple[str,str]:
    low={k.lower():str(v) for k,v in row.items()}
    return low.get("parameter",""), low.get("requirement","")

def unbold(paragraph):
    for run in paragraph.runs:
        run.font.bold=False

def add_bullet(doc:Document,text:str,style_available:bool):
    if style_available:
        p=doc.add_paragraph(text,style="List Bullet")
    else:
        p=doc.add_paragraph(f"• {text}")
    unbold(p)

def build_doc(header_tpl:pathlib.Path,sections:dict,name:str,cid:int,bullet_style:bool)->pathlib.Path:
    doc=Document(header_tpl)
    # Title
    title=doc.add_paragraph(f"RFQ for {name}")
    title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size=Pt(16)
    title.style.font.bold=True
    doc.add_paragraph()

    # Sections
    def add_heading(num, text):
        doc.add_heading(f"{num}. {text}", level=2)

    add_heading(1,"Introduction")
    p=doc.add_paragraph(sections.get("introduction",""))
    unbold(p)

    add_heading(2,"Scope of Supply")
    p=doc.add_paragraph(sections.get("scope",""))
    unbold(p)

    add_heading(3,"Technical Requirements")
    table=doc.add_table(rows=1, cols=2)
    table.alignment=WD_TABLE_ALIGNMENT.CENTER
    hdr=table.rows[0].cells
    hdr[0].text="Parameter"; hdr[1].text="Requirement"
    for c in hdr:
        c.paragraphs[0].runs[0].font.bold=True
        c.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    for r in sections.get("tech_table",[]):
        param,req=normalise(r)
        cells=table.add_row().cells
        cells[0].text=param
        cells[1].text=req

    add_heading(4,"Commercial Requirements")
    for b in sections.get("commercial_terms",[]):
        add_bullet(doc,b,bullet_style)

    add_heading(5,"Documentation Requirements")
    for b in sections.get("docs_required",[]):
        add_bullet(doc,b,bullet_style)

    add_heading(6,"Submission Guidelines")
    p=doc.add_paragraph(
        f"Please submit your quotations via email with the subject line:\n"
        f"“Quotation for {name} - RESL”.")
    unbold(p)
    add_bullet(doc,"Contact Person: P. Pranay Kiran, A. Sai Nithin",bullet_style)
    add_bullet(doc,"Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in",bullet_style)

    add_heading(7,"Confidentiality Clause")
    p=doc.add_paragraph(
        "All quotations and related documents submitted in response to this RFQ "
        "will be treated as confidential and used solely for evaluation purposes.")
    unbold(p)

    out=OUT_DIR/f"RFQ_{cid:02d}_{name.replace(' ','_')}.docx"
    doc.save(out)
    return out

# bullet style availability
try:
    Document(TEMPLATE_HDR).styles["List Bullet"]
    BULLET=True
except KeyError:
    BULLET=False
    print("ℹ️  'List Bullet' style missing – plain bullets used.")

# ── 4. MAIN LOOP ────────────────────────────────────────────
generated=[]
for rec in tqdm(df.to_dict(orient="records"),desc="RFQs"):
    try:
        secs=draft_sections(rec["Item"])
        path=build_doc(TEMPLATE_HDR,secs,rec["Item"],int(rec["cid"]),BULLET)
        generated.append(path)
    except Exception as e:
        print(f"⚠️  {rec['Item']} skipped: {e}")

print(f"✓ Created {len(generated)} RFQ files in {OUT_DIR.resolve()}")

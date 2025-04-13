
"""
build_rfqs_from_excel.py – v2.2
Adds support for an optional **Description** column in data_cp_1.xlsx.
If present and non‑blank, the description text is fed to GPT and explicitly
inserted as extra bullet points in the Technical Requirements section.

Formatting fixes (no stray bold) remain intact.
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
    raise RuntimeError("OPENAI_API_KEY not set (env var or .env file)")

# ── CONFIG ───────────────────────────────────────────────────
EXCEL_PATH   = pathlib.Path("data/data_cp_1.xlsx")
TEMPLATE_HDR = pathlib.Path("templates/u1.docx")
OUT_DIR      = pathlib.Path("out/rfq_docs")
MODEL        = "gpt-4o-mini"
TEMPERATURE  = 0.2

OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 1. LOAD LIST ─────────────────────────────────────────────
raw_df = pd.read_excel(EXCEL_PATH, header=2).iloc[1:]
col_map = {
    "Unnamed: 1": "cid",
    "Unnamed: 2": "Item",
    "Unnamed: 3": "Priority",
    "Unnamed: 4": "Assignee",
    "Unnamed: 5": "Description",  # may or may not exist
}
for k, v in col_map.items():
    if k in raw_df.columns:
        raw_df = raw_df.rename(columns={k: v})
df = (raw_df
      .loc[:, [c for c in ["cid","Item","Priority","Assignee","Description"] if c in raw_df.columns]]
      .dropna(subset=["Item"])
      .reset_index(drop=True))

# ── 2. GPT PROMPTS ───────────────────────────────────────────
SYSTEM_PROMPT = textwrap.dedent("""You are a senior procurement engineer at Renewable Energy Systems Limited,
a defence lithium‑thermal battery manufacturer (AS9100D). Draft RFQ sections
with measurable specifications aligned to MIL‑STD / ASTM / IEC standards.

Return JSON with keys:
  introduction, scope, tech_table(list[{parameter,requirement}]),
  commercial_terms(list[str]), docs_required(list[str]).
""")

USER_TEMPLATE = (
    "Draft RFQ sections for capital equipment '{name}'. "
    "{desc_clause}"
    "Use lowercase keys 'parameter' and 'requirement' in tech_table."
)

def draft_sections(name:str, desc:str|None)->dict:
    desc_clause = f"Include these user‑specified criteria: {desc}. " if desc else ""
    user_prompt = USER_TEMPLATE.format(name=name, desc_clause=desc_clause)
    resp = openai.chat.completions.create(
        model=MODEL, temperature=TEMPERATURE,
        messages=[{"role":"system","content":SYSTEM_PROMPT},
                  {"role":"user","content":user_prompt}],
        response_format={"type":"json_object"})
    return json.loads(resp.choices[0].message.content)

# ── 3. WORD HELPERS ─────────────────────────────────────────-
def normalise(row:dict)->Tuple[str,str]:
    low={k.lower():str(v) for k,v in row.items()}
    return low.get("parameter",""), low.get("requirement","")

def unbold(paragraph):
    for run in paragraph.runs: run.font.bold=False

def add_bullet(doc:Document,text:str,style_ok:bool):
    p = doc.add_paragraph(text, style="List Bullet") if style_ok else doc.add_paragraph(f"• {text}")
    unbold(p)

def build_doc(header_tpl:pathlib.Path,sections:dict,name:str,cid:int,style_ok:bool)->pathlib.Path:
    doc=Document(header_tpl)
    title=doc.add_paragraph(f"RFQ for {name}")
    title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    title.style.font.size=Pt(16); title.style.font.bold=True
    doc.add_paragraph()

    def heading(num,text): doc.add_heading(f"{num}. {text}", level=2)

    heading(1,"Introduction"); unbold(doc.add_paragraph(sections.get("introduction","")))
    heading(2,"Scope of Supply"); unbold(doc.add_paragraph(sections.get("scope","")))

    heading(3,"Technical Requirements")
    table=doc.add_table(rows=1, cols=2); table.alignment=WD_TABLE_ALIGNMENT.CENTER
    hdr=table.rows[0].cells; hdr[0].text="Parameter"; hdr[1].text="Requirement"
    for c in hdr: c.paragraphs[0].runs[0].font.bold=True; c.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    for r in sections.get("tech_table",[]): 
        p,req=normalise(r); cells=table.add_row().cells; cells[0].text=p; cells[1].text=req

    heading(4,"Commercial Requirements")
    for b in sections.get("commercial_terms",[]): add_bullet(doc,b,style_ok)

    heading(5,"Documentation Requirements")
    for b in sections.get("docs_required",[]): add_bullet(doc,b,style_ok)

    heading(6,"Submission Guidelines")
    unbold(doc.add_paragraph(
        f"Please submit your quotations via email with the subject line:\n"
        f"“Quotation for {name} - RESL”."))
    add_bullet(doc,"Contact Person: P. Pranay Kiran, A. Sai Nithin",style_ok)
    add_bullet(doc,"Email: Designengineer.pranay@resindia.co.in, engineer.resl1@resindia.co.in",style_ok)

    heading(7,"Confidentiality Clause")
    unbold(doc.add_paragraph(
        "All quotations and related documents submitted in response to this RFQ "
        "will be treated as confidential and used solely for evaluation purposes."))

    out=OUT_DIR/f"RFQ_{cid:02d}_{name.replace(' ','_')}.docx"; doc.save(out); return out

# bullet style availability
try: Document(TEMPLATE_HDR).styles["List Bullet"]; STYLE_OK=True
except KeyError: STYLE_OK=False; print("ℹ️  'List Bullet' style missing – plain bullets used.")

# ── 4. MAIN LOOP ────────────────────────────────────────────
generated=[]
for rec in tqdm(df.to_dict(orient="records"),desc="RFQs"):
    try:
        sections=draft_sections(rec["Item"], rec.get("Description"))
        generated.append(build_doc(TEMPLATE_HDR,sections,rec["Item"],int(rec["cid"]),STYLE_OK))
    except Exception as e:
        print(f"⚠️  {rec['Item']} skipped: {e}")

print(f"✓ Created {len(generated)} RFQ file(s) in {OUT_DIR.resolve()}")

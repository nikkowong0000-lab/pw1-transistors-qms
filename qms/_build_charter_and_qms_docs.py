# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

BASE = r"C:\Users\nikko\Downloads\REA-Prompts-main\REA-Prompts-main\pw1-transistors\qms"


def set_normal(doc):
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


def add_h(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    return p


def add_p(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    return p


def add_bullets(doc, items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


# ═══════════════════════════════════════════
# 1) Team Charter
# ═══════════════════════════════════════════
doc = Document()
set_normal(doc)
t = doc.add_paragraph()
r = t.add_run("University of Bath — Year 2 Project Week")
r.bold = True
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0x1A, 0x6B, 0x56)

doc.add_heading("Team Charter — PW1 Transistors", 0)
add_p(doc, "Project: The secret life of transistors (Week 6)  |  Group: _______________  |  Date: _______________")
add_p(doc, "All members must read, agree, and sign. Store in QMS → Human resources.")

add_h(doc, "1. Team objectives and expectations", 1)
add_bullets(doc, [
    "Deliver a working half adder by Thursday evening and join another team for the Friday 10:15 full-adder challenge.",
    "Complete characterisation, logic-gate evidence, and a non-technical group slideshow.",
    "Maintain a QMS ready for the Thursday 11:15 audit (find any named document in ≤30 seconds).",
    "Each member submits an individual ≤3-minute video (Thursday evening ready; Moodle by Friday 12:00).",
    "Behave professionally: on time, honest about blockers, respectful disagreement (Controversy with Civility).",
])

add_h(doc, "2. Roles and fair distribution of responsibilities", 1)
add_p(doc, "Fill names after the skills matrix. One primary owner per stream; others support.")
table = doc.add_table(rows=5, cols=3)
table.style = "Table Grid"
hdr = ["Stream / area", "Primary owner(s)", "Support"]
for i, h in enumerate(hdr):
    table.rows[0].cells[i].text = h
rows = [
    ("A — Characterisation (I/O curves)", "", ""),
    ("B — Logic gates & half adder", "", ""),
    ("C — QMS / documentation", "", ""),
    ("Slideshow + individual videos (coord.)", "", ""),
]
for r_i, row in enumerate(rows, 1):
    for c_i, val in enumerate(row):
        table.rows[r_i].cells[c_i].text = val

add_h(doc, "3. Missed deadlines or uneven contribution", 1)
add_bullets(doc, [
    "Raise a blocker in the group chat as soon as it appears (same half-day).",
    "If a deadline will be missed: message the group ≥2 hours before, propose a catch-up plan.",
    "First uneven-load issue: private reminder from the day’s facilitator.",
    "Repeated no-shows / silent free-riding: escalate to the whole team in the next sync; reassign tasks; note in QMS Human resources record.",
    "Critical-path tasks (half adder, QMS audit prep, videos) take priority over non-critical writing.",
])

add_h(doc, "4. Communication methods", 1)
add_bullets(doc, [
    "Primary: Microsoft Teams group chat.",
    "Daily stand-up: 10 minutes at an agreed time (state: done / doing / blocker).",
    "File store: shared QMS folder (OneDrive/Teams Files) mirroring the local QMS tree.",
    "Urgent lab issues: call/message the stream owner immediately.",
])

add_h(doc, "5. Decision-making when opinions differ", 1)
add_bullets(doc, [
    "Try to reach consensus in ≤10 minutes.",
    "If still split: vote (majority). Tie → critical-path owner decides for technical issues; QMS owner for document issues.",
    "Use Controversy with Civility: state the idea, evidence, listen, no personal attacks; record the decision in chat.",
])

add_h(doc, "6. Signatures — all members agree", 1)
add_p(doc, "By signing, I agree to this charter and the skills-matrix-based task plan.")
sig = doc.add_table(rows=7, cols=3)
sig.style = "Table Grid"
sig.rows[0].cells[0].text = "Name"
sig.rows[0].cells[1].text = "Signature"
sig.rows[0].cells[2].text = "Date"
for i in range(1, 7):
    sig.rows[i].cells[0].text = f"Member {i}: _______________"
    sig.rows[i].cells[1].text = ""
    sig.rows[i].cells[2].text = ""

charter_path = os.path.join(BASE, "Human resources", "PW1_Team_Charter.docx")
doc.save(charter_path)
print("CHARTER", charter_path)


# ═══════════════════════════════════════════
# 2) Naming procedure
# ═══════════════════════════════════════════
proc = Document()
set_normal(proc)
proc.add_heading("QMS File Naming Procedure — PW1 Transistors", 0)
add_p(proc, "Location: Quality documents / Procedure  |  Version: v1.0  |  Status: Controlled")
add_p(proc, "Purpose: Everyone names files the same way so documents are easy to find (including during the 30-second audit).")

add_h(proc, "1. Naming convention", 1)
add_p(proc, "Format:", bold=True)
add_p(proc, "PW1_Transistors_<Area>_<DocumentName>_v<Major.Minor>.<ext>")
add_bullets(proc, [
    "Area codes: QM (quality manual), PROC (procedure), FORM, QC, ENG, OPS, HR, SALES, GAL",
    "DocumentName: short English CamelCase or underscores, no spaces if possible (spaces allowed if clear).",
    "Version: v0.1 draft → v1.0 first approved → v1.1 small edit → v2.0 major change.",
    "Date optional suffix if helpful: _YYYYMMDD",
])
add_p(proc, "Examples:", bold=True)
add_bullets(proc, [
    "PW1_Transistors_PROC_FileNaming_v1.0.docx",
    "PW1_Transistors_HR_TeamCharter_v1.0.docx",
    "PW1_Transistors_HR_SkillsMatrix_v1.0.xlsx",
    "PW1_Transistors_OPS_GanttChart_v1.0.png",
    "PW1_Transistors_ENG_TraceabilityMatrix_v0.1.xlsx",
    "PW1_Transistors_GAL_DocumentRegister_v1.0.xlsx",
])

add_h(proc, "2. Version control rules", 1)
add_bullets(proc, [
    "Do not overwrite history blindly: keep the previous major file or use Git for Engineering.",
    "When content is approved by the team, bump to the next vX.0 and update the Documents gallery.",
    "Put the version in BOTH the filename and the document header.",
    "Engineering circuit/code changes: use Git branch → merge to main (evidence for change management).",
])

add_h(proc, "3. Where files live", 1)
add_bullets(proc, [
    "Quality manual → Quality documents/Quality manual",
    "This naming procedure → Quality documents/Procedure",
    "Forms → Quality documents/Forms",
    "Test plans/results → Quality control",
    "Schematics, code, traceability → Engineering",
    "Gantt / time plan → Operations",
    "Purchasing → empty this week",
    "Client/audience materials → Sales_customer",
    "Charter, skills matrix → Human resources",
    "Master list of all files → Documents gallery ONLY (register spreadsheet)",
])

add_h(proc, "4. Easy to find / everyone understands", 1)
add_bullets(proc, [
    "English names only in the QMS.",
    "No ‘final_final2’ — use version numbers.",
    "After any add/move/rename: update Documents gallery the same day.",
    "Before Thursday audit: practise finding 5 random files in ≤30 seconds.",
])

proc_path = os.path.join(BASE, "Quality documents", "Procedure", "PW1_Transistors_PROC_FileNaming_v1.0.docx")
proc.save(proc_path)
print("PROC", proc_path)


# ═══════════════════════════════════════════
# 3) Quality manual (short commitment)
# ═══════════════════════════════════════════
qm = Document()
set_normal(qm)
qm.add_heading("Quality Manual — PW1 Transistors (Commitment to Quality)", 0)
add_p(qm, "Version: v1.0  |  Group: _______________  |  Date: _______________")
add_h(qm, "1. Purpose", 1)
add_p(qm, "This project week uses a lightweight ISO 9001-inspired QMS so technical work, evidence, and skills submissions are controlled, findable, and auditable.")
add_h(qm, "2. Quality policy (our commitment)", 1)
add_bullets(qm, [
    "We plan work with a Gantt chart and critical paths.",
    "We record who owns what via the team charter and skills matrix.",
    "We name and store files using the controlled naming procedure.",
    "We verify circuits against truth tables / measurement plans before claiming completion.",
    "We improve continuously (short daily review of blockers and document gaps).",
])
add_h(qm, "3. Scope", 1)
add_p(qm, "Applies to all PW1 Transistors deliverables during Week 6: lab evidence, Mahara skills evidence, QMS audit pack, and Friday kit return.")
add_h(qm, "4. Responsibilities", 1)
add_bullets(qm, [
    "All members: follow naming rules; update gallery when adding files.",
    "QMS owner (Stream C): keep tree intact; lead audit drill.",
    "Stream owners: ensure technical evidence is filed under Engineering / Quality control.",
])
qm_path = os.path.join(BASE, "Quality documents", "Quality manual", "PW1_Transistors_QM_QualityManual_v1.0.docx")
qm.save(qm_path)
print("QM", qm_path)


# ═══════════════════════════════════════════
# 4) Documents gallery Excel
# ═══════════════════════════════════════════
wb = Workbook()
ws = wb.active
ws.title = "Document_Register"
thin = Border(
    left=Side(style="thin", color="CFD8DC"),
    right=Side(style="thin", color="CFD8DC"),
    top=Side(style="thin", color="CFD8DC"),
    bottom=Side(style="thin", color="CFD8DC"),
)
header_fill = PatternFill("solid", fgColor="1A6B56")
edit = PatternFill("solid", fgColor="FFF8E1")
ws["A1"] = "PW1 Transistors — Documents Gallery (list of ALL QMS files and locations)"
ws["A1"].font = Font(bold=True, size=13, color="1A6B56")
ws.merge_cells("A1:G1")
ws["A2"] = "Update this sheet whenever a file is added, renamed, or moved. This folder should contain ONLY this register."
ws["A2"].fill = edit
ws.merge_cells("A2:G2")

headers = ["File name", "Folder path (from qms/)", "Area", "Version", "Owner", "Date added", "Notes"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(4, c, h)
    cell.fill = header_fill
    cell.font = Font(bold=True, color="FFFFFF")
    cell.border = thin
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

seed = [
    ("PW1_Transistors_QM_QualityManual_v1.0.docx", "Quality documents/Quality manual/", "QM", "v1.0", "", "", "Quality commitment"),
    ("PW1_Transistors_PROC_FileNaming_v1.0.docx", "Quality documents/Procedure/", "PROC", "v1.0", "", "", "Naming procedure"),
    ("PW1_Team_Charter.docx", "Human resources/", "HR", "v1.0", "", "", "Sign all members"),
    ("PW1_Skills_Matrix.xlsx", "Human resources/", "HR", "v1.0", "", "", "Fill ratings 1-5"),
    ("PW1_GANTT_SUBMIT.html", "Operations/", "OPS", "v1.0", "", "", "Save PNG for Mahara"),
    ("PW1_Gantt_ProjectLibre.xml", "Operations/", "OPS", "v1.0", "", "", "Optional ProjectLibre"),
    ("PW1_Transistors_GAL_DocumentRegister_v1.0.xlsx", "Documents gallery/", "GAL", "v1.0", "", "", "THIS file"),
]
for r, row in enumerate(seed, 5):
    for c, v in enumerate(row, 1):
        cell = ws.cell(r, c, v)
        cell.border = thin
        if c in (5, 6):
            cell.fill = edit

# empty rows for future files
for r in range(5 + len(seed), 5 + len(seed) + 15):
    for c in range(1, 8):
        cell = ws.cell(r, c, "")
        cell.border = thin
        if c in (5, 6):
            cell.fill = edit

widths = [48, 36, 8, 10, 14, 12, 24]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

gal_path = os.path.join(BASE, "Documents gallery", "PW1_Transistors_GAL_DocumentRegister_v1.0.xlsx")
wb.save(gal_path)
print("GAL", gal_path)

placeholders = {
    os.path.join(BASE, "Purchasing", "README_EMPTY.txt"): "Purchasing is empty for this project week.\n",
    os.path.join(BASE, "Quality documents", "Forms", "README_EMPTY.txt"): "Forms folder may remain empty unless you create quality forms.\n",
    os.path.join(BASE, "Engineering", "README_PUT_TRACEABILITY_HERE.txt"): "Put schematics, code, and the D&D traceability matrix (complete to verification columns) here.\n",
    os.path.join(BASE, "Quality control", "README_PUT_TEST_RECORDS_HERE.txt"): "Put measurement plans and truth-table / characterisation results here.\n",
    os.path.join(BASE, "Sales_customer", "README_PUT_CLIENT_MATERIALS_HERE.txt"): "Put non-technical audience / patient-facing communication materials here.\n",
}
for path, text in placeholders.items():
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

print("DONE")

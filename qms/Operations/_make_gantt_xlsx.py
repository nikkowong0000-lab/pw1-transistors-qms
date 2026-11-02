# -*- coding: utf-8 -*-
"""Build a submission-ready PW1 Time Management Gantt in Excel."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.page import PageMargins

wb = Workbook()

# ── palette ──────────────────────────────────────────────
C = {
    "ink": "142028",
    "mint": "1A6B56",
    "mint_lt": "E8F5F0",
    "sand": "D9E4DC",
    "line": "CFD8DC",
    "cp_a": "C62828",
    "cp_a_lt": "FFCDD2",
    "cp_b": "1565C0",
    "cp_b_lt": "BBDEFB",
    "cp_c": "EF6C00",
    "cp_c_lt": "FFE0B2",
    "buf": "7B1FA2",
    "buf_lt": "E1BEE7",
    "normal": "546E7A",
    "normal_lt": "CFD8DC",
    "mile": "212121",
    "edit": "FFF8E1",
    "white": "FFFFFF",
    "am": "F5F7F6",
    "pm": "EEF2F0",
}

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=None, name="Calibri"):
    return Font(bold=bold, size=size, color=color or C["ink"], name=name)

thin = Border(
    left=Side(style="thin", color=C["line"]),
    right=Side(style="thin", color=C["line"]),
    top=Side(style="thin", color=C["line"]),
    bottom=Side(style="thin", color=C["line"]),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_al = Alignment(horizontal="left", vertical="center", wrap_text=True)

TYPE_FILL = {
    "cp-a": fill(C["cp_a"]),
    "cp-b": fill(C["cp_b"]),
    "cp-c": fill(C["cp_c"]),
    "buf": fill(C["buf"]),
    "normal": fill(C["normal"]),
    "mile": fill(C["mile"]),
}
TYPE_FONT = {
    "cp-a": font(True, 9, C["white"]),
    "cp-b": font(True, 9, C["white"]),
    "cp-c": font(True, 9, C["white"]),
    "buf": font(True, 9, C["white"]),
    "normal": font(True, 9, C["white"]),
    "mile": font(True, 9, C["white"]),
}

# Half-day slots Mon–Fri
SLOTS = [
    ("Mon", "AM", "09:30"),
    ("Mon", "PM", "13:30"),
    ("Tue", "AM", "09:30"),
    ("Tue", "PM", "13:30"),
    ("Wed", "AM", "09:30"),
    ("Wed", "PM", "13:30"),
    ("Thu", "AM", "09:30"),
    ("Thu", "PM", "13:30"),
    ("Fri", "AM", "09:30"),
    ("Fri", "PM", "13:30"),
]

# name_zh, name_en, stream, cp_label, type, start, end, bar_label
TASKS = [
    ("Kit 領取 11:15", "Kit handout", "All", "CP-A/B", "mile", 0, 0, "KIT"),
    ("研究 BJT/MOSFET", "Research", "A: 改名1, 改名2", "CP-C", "cp-c", 0, 2, "Research"),
    ("I/O 特性量測", "I/O char.", "A: 改名1, 改名2", "—", "normal", 1, 4, "I/O"),
    ("邏輯閘", "Logic gates", "B: 改名3, 改名4", "CP-A", "cp-a", 1, 3, "Gates"),
    ("半加器", "Half adder", "B: 改名3, 改名4", "CP-A", "cp-a", 4, 5, "HA"),
    ("緩衝（半加器後）", "Buffer HA", "B", "CP-A", "buf", 5, 5, "BUF"),
    ("QMS 建置", "QMS setup", "C: 改名5, 改名6", "CP-B", "cp-b", 0, 5, "QMS"),
    ("緩衝（QMS 後）", "Buffer QMS", "C", "CP-B", "buf", 5, 5, "BUF"),
    ("稽核演練 30s", "Audit prep", "C + All", "CP-B", "cp-b", 5, 6, "Drill"),
    ("QMS 稽核 11:15", "Audit", "All", "CP-B", "mile", 6, 6, "AUDIT"),
    ("團隊簡報", "Slideshow", "C + A/B", "CP-C", "cp-c", 3, 7, "Slides"),
    ("個人 3 分鐘影片", "Video", "每人自己", "CP-C", "cp-c", 3, 7, "Video"),
    ("緩衝（影片／簡報）", "Buffer media", "C / Individual", "CP-C", "buf", 7, 7, "BUF"),
    ("全加器演練", "FA rehearse", "B + A", "CP-A", "cp-a", 7, 8, "FA"),
    ("緩衝（挑戰前）", "Buffer FA", "All", "CP-A", "buf", 8, 8, "BUF"),
    ("全加器挑戰 10:15", "FA challenge", "All", "CP-A", "mile", 8, 8, "FA★"),
    ("影片 Moodle 12:00", "Video submit", "每人自己", "CP-C", "mile", 8, 8, "VID★"),
    ("器材歸還 12:15", "Kit return", "All", "—", "normal", 8, 9, "Return"),
    ("Leadership 文件（並行）", "Leadership", "C + All", "—", "normal", 0, 4, "Lead"),
    ("循環經濟寫作（可砍）", "Circular econ.", "C: 改名5, 改名6", "—", "normal", 2, 6, "CE"),
]


# ═══════════════════════════════════════════════════════════
# Sheet 1 — GANTT (print / screenshot this)
# ═══════════════════════════════════════════════════════════
ws = wb.active
ws.title = "GANTT"
ws.sheet_view.showGridLines = False
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
ws.print_title_rows = "1:6"
ws.freeze_panes = "D7"

# Title block
ws.merge_cells("A1:N1")
ws["A1"] = "University of Bath · Y2 PW1 Transistors · Advanced Time Management"
ws["A1"].font = font(True, 16, C["mint"])
ws["A1"].alignment = left_al

ws.merge_cells("A2:N2")
ws["A2"] = "GANTT Chart — Week 6 ｜ 截圖／列印本頁交 Mahara（含圖例）"
ws["A2"].font = font(True, 12, C["ink"])

ws.merge_cells("A3:N3")
ws["A3"] = (
    "【改黃色人名】見 Tasks_EDIT 工作表 → 再貼回本表 B 欄。"
    " 紅=CP-A 全加器｜藍=CP-B 稽核｜橙=CP-C 影片｜紫=緩衝｜灰=並行｜黑◆=硬節點"
)
ws["A3"].font = font(False, 9, "4A5C66")
ws["A3"].fill = fill(C["edit"])

# Meta row
ws["A4"] = "Group:"
ws["B4"] = "PW1 Transistors Group __"
ws["B4"].fill = fill(C["edit"])
ws["B4"].font = font(True, 11)
ws["C4"] = "Hours:"
ws.merge_cells("D4:G4")
ws["D4"] = "Mon–Fri 09:30–17:00（午休 12:30–13:30）· 6人 · 淨工時≈6.5h/人/日"
ws["D4"].fill = fill(C["edit"])
ws["H4"] = "Streams:"
ws.merge_cells("I4:N4")
ws["I4"] = "A 量測｜B 閘/加器｜C QMS/文件（可並行）"
ws["I4"].fill = fill(C["edit"])

for col in range(1, 15):
    ws.cell(4, col).border = thin
    ws.cell(4, col).alignment = left_al if col != 1 else center

# Day header row 5 (merged Mon..Fri)
ws.cell(5, 1, "").fill = fill(C["mint"])
ws.cell(5, 2, "").fill = fill(C["mint"])
ws.cell(5, 3, "").fill = fill(C["mint"])
days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
for i, d in enumerate(days):
    c0 = 4 + i * 2
    ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 1)
    cell = ws.cell(5, c0, d)
    cell.fill = fill(C["mint"])
    cell.font = font(True, 12, C["white"])
    cell.alignment = center
    ws.cell(5, c0 + 1).fill = fill(C["mint"])
    for c in (c0, c0 + 1):
        ws.cell(5, c).border = thin

# Sub-header row 6
sub = ["任務 Task", "負責人 Owner 【改】", "CP"]
for col, h in enumerate(sub, 1):
    cell = ws.cell(6, col, h)
    cell.fill = fill(C["sand"])
    cell.font = font(True, 10)
    cell.alignment = center
    cell.border = thin

for i, (day, ap, t) in enumerate(SLOTS):
    col = 4 + i
    cell = ws.cell(6, col, f"{ap}\n{t}")
    cell.fill = fill(C["am"] if ap == "AM" else C["pm"])
    cell.font = font(True, 9)
    cell.alignment = center
    cell.border = thin

# Task rows
for r, (zh, en, stream, cp, typ, s, e, label) in enumerate(TASKS, 7):
    # Task name
    c1 = ws.cell(r, 1, f"{zh}\n{en}")
    c1.alignment = left_al
    c1.font = font(True, 9) if typ == "mile" else font(False, 9)
    c1.border = thin
    if typ == "mile":
        c1.fill = fill("ECEFF1")

    # Owner (editable look)
    c2 = ws.cell(r, 2, stream)
    c2.alignment = left_al
    c2.font = font(False, 9)
    c2.border = thin
    c2.fill = fill(C["edit"])

    # CP tag
    c3 = ws.cell(r, 3, cp)
    c3.alignment = center
    c3.border = thin
    c3.font = font(True, 8, C["white"] if cp.startswith("CP") else C["ink"])
    if cp.startswith("CP-A"):
        c3.fill = fill(C["cp_a"])
    elif cp.startswith("CP-B"):
        c3.fill = fill(C["cp_b"])
    elif cp.startswith("CP-C"):
        c3.fill = fill(C["cp_c"])
    else:
        c3.fill = fill(C["sand"])

    # Bars
    for c in range(4, 14):
        cell = ws.cell(r, c, "")
        cell.border = thin
        cell.alignment = center
        slot = c - 4
        # subtle zebra for empty
        if slot % 2 == 0:
            cell.fill = fill("FAFBFA")
        else:
            cell.fill = fill("F3F5F4")
        if s <= slot <= e:
            cell.fill = TYPE_FILL[typ]
            cell.font = TYPE_FONT[typ]
            # label only on first cell of bar (or single cell)
            if slot == s:
                cell.value = label

    ws.row_dimensions[r].height = 28

# Column widths
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 7
for c in range(4, 14):
    ws.column_dimensions[get_column_letter(c)].width = 8

ws.row_dimensions[1].height = 24
ws.row_dimensions[2].height = 20
ws.row_dimensions[3].height = 28
ws.row_dimensions[4].height = 22
ws.row_dimensions[5].height = 22
ws.row_dimensions[6].height = 32

# Legend block
leg0 = 7 + len(TASKS) + 1
ws.merge_cells(start_row=leg0, start_column=1, end_row=leg0, end_column=3)
ws.cell(leg0, 1, "圖例 LEGEND").font = font(True, 11, C["mint"])

legend_items = [
    (C["cp_a"], "CP-A 技術路徑 → Fri 10:15 Full adder challenge"),
    (C["cp_b"], "CP-B QMS 路徑 → Thu 11:15 Audit（30 秒找檔）"),
    (C["cp_c"], "CP-C 溝通路徑 → Fri 12:00 Moodle video"),
    (C["buf"], "Buffer 緩衝（只吸延遲，不加新範圍）"),
    (C["normal"], "並行／非關鍵任務（可延後或縮短）"),
    (C["mile"], "◆ 硬性 milestone（deadline）"),
]
for i, (color, text) in enumerate(legend_items):
    r = leg0 + 1 + i
    ws.cell(r, 1, "").fill = fill(color)
    ws.cell(r, 1).border = thin
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(r, 2, text).font = font(False, 10)
    ws.cell(r, 2).alignment = left_al

# Critical path summary
sum0 = leg0 + 8
ws.merge_cells(start_row=sum0, start_column=1, end_row=sum0, end_column=10)
ws.cell(sum0, 1, "Critical paths（摘要）").font = font(True, 11, C["mint"])
summaries = [
    "CP-A：Kit → Logic gates → Half adder →（BUF）→ FA rehearse → Fri 10:15 challenge",
    "CP-B：QMS setup →（BUF）→ Audit prep → Thu 11:15 audit",
    "CP-C：Research → Slideshow/Video →（BUF Thu PM）→ Fri 12:00 Moodle",
    "若來不及：把 CE／文件人力調去半加器；週三凍結 QMS；影片腳本早定、緩衝只留給剪輯。",
]
for i, t in enumerate(summaries):
    r = sum0 + 1 + i
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    ws.cell(r, 1, t).font = font(False, 9, "37474F")

# Print area
last_row = sum0 + 4
ws.print_area = f"A1:N{last_row}"


# ═══════════════════════════════════════════════════════════
# Sheet 2 — editable task list
# ═══════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Tasks_EDIT")
ws2["A1"] = "✏️ 只改黃色「Owner」欄的名字，然後複製到 GANTT 工作表 B 欄對應列"
ws2["A1"].font = font(True, 12, C["mint"])
ws2.merge_cells("A1:H1")
ws2["A2"] = "PERT: te = (O + 4M + P) / 6  ｜ 單位：小時"
ws2["A2"].font = font(False, 9, "4A5C66")

headers2 = [
    "ID", "任務", "Owner 【改這裡】", "開始", "結束",
    "O", "M", "P", "te=(O+4M+P)/6", "Critical Path", "備註",
]
for c, h in enumerate(headers2, 1):
    cell = ws2.cell(4, c, h)
    cell.fill = fill(C["mint"])
    cell.font = font(True, 10, C["white"])
    cell.alignment = center
    cell.border = thin

# Detailed rows with PERT
detail = [
    ("T01", "Kit handout", "All", "Mon", "Mon", 0.3, 0.5, 1.0, "CP-A|CP-B", "11:15 hard"),
    ("T02", "Research BJT/MOSFET", "A: 改名1, 改名2", "Mon", "Tue", 1.5, 2.5, 4.0, "CP-C", "Feeds slides+video"),
    ("T03", "I/O characterisation", "A: 改名1, 改名2", "Mon", "Wed", 2.0, 3.5, 5.5, "", "Parallel with gates"),
    ("T04", "Logic gates", "B: 改名3, 改名4", "Mon", "Tue", 2.0, 3.0, 5.0, "CP-A", ""),
    ("T05", "Half adder", "B: 改名3, 改名4", "Wed", "Wed", 1.5, 2.5, 4.0, "CP-A", "Highest technical risk"),
    ("T05b", "Buffer after HA", "B", "Wed", "Wed", 1.0, 1.0, 1.0, "CP-A", "Task buffer 1h"),
    ("T09", "QMS setup", "C: 改名5, 改名6", "Mon", "Wed", 1.5, 2.5, 4.0, "CP-B", ""),
    ("T09b", "Buffer after QMS", "C", "Wed", "Wed", 1.0, 1.0, 1.0, "CP-B", "Task buffer 1h"),
    ("T10", "Audit prep", "C + All", "Wed", "Thu", 0.5, 1.0, 1.5, "CP-B", "30s find drill"),
    ("M-AUD", "QMS Audit", "All", "Thu", "Thu", 0.75, 0.75, 0.75, "CP-B", "11:15 hard"),
    ("T07", "Team slideshow", "C + A/B", "Tue", "Thu", 1.5, 2.5, 4.0, "CP-C", ""),
    ("T08", "Individual video", "每人自己", "Tue", "Thu", 1.5, 2.5, 4.5, "CP-C", "Done Thu evening"),
    ("T08b", "Buffer video/slides", "C / Individual", "Thu", "Thu", 1.0, 1.0, 1.0, "CP-C", "End buffer 1h"),
    ("T06", "Full adder rehearse", "B + A", "Thu", "Fri", 1.0, 2.0, 3.5, "CP-A", ""),
    ("T06b", "Buffer pre-challenge", "All", "Fri", "Fri", 0.5, 0.5, 0.5, "CP-A", "0.5h"),
    ("M-FA", "Full adder challenge", "All", "Fri", "Fri", 1.75, 1.75, 1.75, "CP-A", "10:15 hard"),
    ("M-VID", "Video Moodle", "每人自己", "Fri", "Fri", 0.5, 0.5, 0.5, "CP-C", "12:00 hard"),
    ("T13", "Kit return", "All", "Fri", "Fri", 0.3, 0.5, 1.0, "", "12:15"),
    ("T12", "Leadership docs", "C + All", "Mon", "Wed", 1.0, 2.0, 3.0, "", "Parallel"),
    ("T11", "Circular economy", "C: 改名5, 改名6", "Tue", "Thu", 1.0, 1.5, 2.5, "", "Cut first if late"),
]

for r, row in enumerate(detail, 5):
    tid, task, owner, ds, de, o, m, p, cp, note = row
    vals = [tid, task, owner, ds, de, o, m, p, None, cp, note]
    for c, v in enumerate(vals, 1):
        cell = ws2.cell(r, c, v if v is not None else "")
        cell.border = thin
        cell.alignment = center if c not in (2, 3, 11) else left_al
        cell.font = font(False, 10)
        if c == 3:
            cell.fill = fill(C["edit"])
            cell.font = font(True, 10)
    # PERT formula
    formula_cell = ws2.cell(r, 9)
    formula_cell.value = f"=ROUND((F{r}+4*G{r}+H{r})/6,2)"
    formula_cell.border = thin
    formula_cell.alignment = center
    formula_cell.font = font(True, 10, C["mint"])
    # CP colour
    cp_cell = ws2.cell(r, 10)
    if "CP-A" in cp:
        cp_cell.fill = fill(C["cp_a_lt"])
    elif "CP-B" in cp:
        cp_cell.fill = fill(C["cp_b_lt"])
    elif "CP-C" in cp:
        cp_cell.fill = fill(C["cp_c_lt"])

widths2 = [8, 24, 22, 8, 8, 6, 6, 6, 14, 12, 22]
for i, w in enumerate(widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = "A5"
ws2.auto_filter.ref = f"A4:K{4 + len(detail)}"


# ═══════════════════════════════════════════════════════════
# Sheet 3 — Mahara paste + risks
# ═══════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Mahara_Text")
ws3["A1"] = "Mahara 可貼文字（改黃色格）"
ws3["A1"].font = font(True, 14, C["mint"])
ws3.merge_cells("A1:B1")

blocks = [
    (3, "Working hours / resources",
     "6 people; Mon–Fri 09:30–17:00 (lunch 12:30–13:30); ~6.5 net h/person/day. "
     "Three parallel streams after kit handout: A characterisation, B gates/adder, C QMS/docs."),
    (5, "GANTT / PERT note",
     "Tasks were broken down and timed with PERT te=(O+4M+P)/6 (see Tasks_EDIT). "
     "The coloured GANTT on sheet GANTT marks critical paths and buffers. "
     "Parallel work is visible where A/B/C bars overlap on the same days."),
    (7, "Critical paths",
     "CP-A (red): Kit → gates → half adder → rehearsal → Fri 10:15 full-adder challenge. "
     "CP-B (blue): QMS setup → audit prep → Thu 11:15 audit. "
     "CP-C (orange): Research → slides/video → Thu evening freeze → Fri 12:00 Moodle. "
     "If late: move people from circular-economy writing onto the half adder; freeze QMS Wednesday; lock video script early."),
    (9, "Buffers",
     "Task buffers: 1 h after half adder; 1 h after QMS build. "
     "End-of-project buffers: 1 h Thu 16:00–17:00 (video/slides); 0.5 h Fri 09:30–10:00 before challenge. "
     "Buffers absorb delay only — not new scope."),
    (11, "Top 3 scheduling risks（50–100 字）【改】",
     "半加器除錯逾時會直接擠壓週五全加器演練；QMS 檔名未凍結會讓週四 11:15 稽核找檔失敗；"
     "個人影片常低估剪輯而拖到週五中午。對策：週三設半加器檢查點並配置任務緩衝、"
     "週三凍結 QMS 結構並演練 30 秒找檔、週四晚強制影片定稿再於週五 12:00 前提交 Moodle。"),
]

for row, title, body in blocks:
    ws3.cell(row, 1, title).font = font(True, 11)
    ws3.cell(row, 1).fill = fill(C["mint_lt"])
    ws3.cell(row, 1).alignment = left_al
    ws3.cell(row, 1).border = thin
    ws3.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=2)
    cell = ws3.cell(row + 1, 1, body)
    cell.fill = fill(C["edit"])
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.border = thin
    ws3.row_dimensions[row + 1].height = 70

ws3.column_dimensions["A"].width = 36
ws3.column_dimensions["B"].width = 80

ws3["A14"] = "繳交步驟"
ws3["A14"].font = font(True, 11, C["mint"])
ws3["A15"] = (
    "1) Tasks_EDIT 改黃色人名  2) 同步到 GANTT 的 Owner 欄  "
    "3) ShareX／Excel 截 GANTT 整頁（含圖例）  "
    "4) 風險短文貼 Mahara  5) 本 xlsx 放進 QMS/Operations/ 並寫入 Documents gallery"
)
ws3.merge_cells("A15:B15")
ws3["A15"].alignment = Alignment(wrap_text=True)
ws3.row_dimensions[15].height = 40

import os
base = r"C:\Users\nikko\Downloads\REA-Prompts-main\REA-Prompts-main\pw1-transistors\qms\Operations"
out = os.path.join(base, "PW1_GANTT.xlsx")
try:
    wb.save(out)
except PermissionError:
    out = os.path.join(base, "PW1_GANTT_v2.xlsx")
    wb.save(out)
print("Saved", out)

#!/usr/bin/env python3
"""
Step 3: Build a multi-tab, formula-driven Excel grading workbook from a student roster CSV
and a reflections audit JSON.
Author: Gautam Gupta (https://www.linkedin.com/in/higautam/)
"""

import argparse
import csv
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, color="1F4E78")
BODY_FONT = Font(name="Calibri", size=10, color="000000")
GREEN_FILL = PatternFill("solid", fgColor="E2EFDA")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
GRAY_FILL = PatternFill("solid", fgColor="EDEDED")
BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build formula-driven grading workbook.")
    parser.add_argument("--roster", required=True, help="Input CSV roster path")
    parser.add_argument("--audit", required=True, help="Input reflections audit JSON path")
    parser.add_argument("--output", required=True, help="Output Excel .xlsx path")
    args = parser.parse_args()

    with open(args.roster, "r", encoding="utf-8") as f:
        roster = list(csv.DictReader(f))
    with open(args.audit, "r", encoding="utf-8") as f:
        audit = json.load(f)

    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Executive Summary & Metrics"
    ws_ind = wb.create_sheet("Individual Status & Marks")
    ws_mat = wb.create_sheet("Class Reflections Matrix")

    # 1. Individual Status & Marks
    ind_headers = [
        "S.No", "Student ID", "Full Name", "Team", "Product",
        "Presented in Class", "Project Uploaded", "Reflections (/12)",
        "Reflections Status", "Reflections Marks (/10)",
        "Combined Status", "Professor Marks (Presentation)", "Total Marks"
    ]
    for c_idx, h in enumerate(ind_headers, start=1):
        cell = ws_ind.cell(1, c_idx, h)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    stats = {"all_12": 0, "partial": 0, "placeholder": 0, "none": 0, "pres_refl": 0, "pres_no_refl": 0, "up_only": 0, "neither": 0}
    for idx, row in enumerate(roster, start=2):
        sid = row["student_id"].strip().upper()
        pres = row["presented"].strip()
        proj = row["project_uploaded"].strip()
        info = audit.get(sid, {"classes": [], "count": 0, "has_any_file": False})
        cnt = info["count"]

        if cnt == 12:
            r_status, r_fill = "All 12 Complete", GREEN_FILL
            stats["all_12"] += 1
        elif cnt > 0:
            r_status, r_fill = f"Partial ({cnt}/12)", YELLOW_FILL
            stats["partial"] += 1
        elif info["has_any_file"]:
            r_status, r_fill = "Placeholder Only (0/12)", RED_FILL
            stats["placeholder"] += 1
        else:
            r_status, r_fill = "Not Submitted (0/12)", RED_FILL
            stats["none"] += 1

        if pres == "Yes" and cnt > 0:
            comb = "Presented + Reflections"
            stats["pres_refl"] += 1
        elif pres == "Yes":
            comb = "Presented + No Reflections"
            stats["pres_no_refl"] += 1
        elif proj == "Yes":
            comb = "Uploaded Only (Not Presented)"
            stats["up_only"] += 1
        else:
            comb = "Not Submitted"
            stats["neither"] += 1

        vals = [
            idx - 1, sid, row["full_name"], row["team"], row["product"],
            pres, proj, cnt, r_status, f"=ROUND((H{idx}/12)*10, 2)",
            comb, None, f"=J{idx}+L{idx}"
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws_ind.cell(idx, c_idx, val)
            cell.font, cell.border = BODY_FONT, BORDER
            cell.alignment = Alignment(horizontal="center" if c_idx not in (3, 5, 11) else "left", vertical="center")
        ws_ind.cell(idx, 9).fill = r_fill
        ws_ind.cell(idx, 10).fill = r_fill

    # 2. Class 1..12 Matrix
    mat_headers = ["Student ID", "Full Name"] + [f"C{i}" for i in range(1, 13)] + ["Total (/12)", "Marks (/10)"]
    for c_idx, h in enumerate(mat_headers, start=1):
        cell = ws_mat.cell(1, c_idx, h)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, row in enumerate(roster, start=2):
        sid = row["student_id"].strip().upper()
        cset = set(audit.get(sid, {}).get("classes", []))
        ws_mat.cell(idx, 1, sid).border = BORDER
        ws_mat.cell(idx, 2, row["full_name"]).border = BORDER
        for cnum in range(1, 13):
            v = 1 if cnum in cset else 0
            c = ws_mat.cell(idx, 2 + cnum, v)
            c.fill = GREEN_FILL if v else GRAY_FILL
            c.border = BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws_mat.cell(idx, 15, f"=SUM(C{idx}:N{idx})").border = BORDER
        ws_mat.cell(idx, 16, f"=ROUND((O{idx}/12)*10, 2)").border = BORDER

    tot_r = len(roster) + 2
    ws_mat.cell(tot_r, 1, "TOTAL").font = SUBHEADER_FONT
    ws_mat.cell(tot_r, 2, f"{len(roster)} Students").font = SUBHEADER_FONT
    for cnum in range(1, 13):
        col_l = get_column_letter(2 + cnum)
        c = ws_mat.cell(tot_r, 2 + cnum, f"=SUM({col_l}2:{col_l}{tot_r-1})")
        c.fill, c.font, c.border = SUBHEADER_FILL, SUBHEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

    # 3. Executive Summary
    ws_sum["A1"] = "GitHub Cohort Evaluation — Executive Summary & Metrics"
    ws_sum["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    summary_rows = [
        ("Category", "Count"),
        ("Total Students Audited", len(roster)),
        ("All 12 Reflections Complete (10/10)", stats["all_12"]),
        ("Partial Reflections (1–11/12)", stats["partial"]),
        ("Placeholder Only (0/10)", stats["placeholder"]),
        ("Not Submitted (0/10)", stats["none"]),
        ("Presented + Reflections", stats["pres_refl"]),
        ("Presented + No Reflections", stats["pres_no_refl"]),
        ("Uploaded Only (Not Presented)", stats["up_only"]),
        ("Not Submitted Anywhere", stats["neither"]),
    ]
    for r_idx, (k, v) in enumerate(summary_rows, start=3):
        c1, c2 = ws_sum.cell(r_idx, 1, k), ws_sum.cell(r_idx, 2, v)
        c1.border, c2.border = BORDER, BORDER
        if r_idx == 3:
            c1.fill, c2.fill = HEADER_FILL, HEADER_FILL
            c1.font, c2.font = HEADER_FONT, HEADER_FONT

    for ws in (ws_sum, ws_ind, ws_mat):
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 3, 10), 42)

    wb.save(args.output)
    print(f"Generated grading workbook: {args.output}")


if __name__ == "__main__":
    main()

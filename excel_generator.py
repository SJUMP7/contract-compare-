"""
excel_generator.py — Generate a styled Excel workbook from the AI JSON data.
Layout matches the user's Google Sheets template (yellow header / blue period rows / red condition cells).
"""
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# ─── Style constants ──────────────────────────────────────────────────────────
_YELLOW = PatternFill("solid", fgColor="FFFF00")
_BLUE   = PatternFill("solid", fgColor="4A86E8")
_RED    = PatternFill("solid", fgColor="CC0000")
_LGRAY  = PatternFill("solid", fgColor="F3F4F6")

_WHITE_BOLD  = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
_BLACK_BOLD  = Font(bold=True, name="Calibri", size=10)
_BLACK_NORM  = Font(name="Calibri", size=10)
_RED_BOLD_BIG = Font(color="CC0000", bold=True, name="Calibri", size=12)

_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

_THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _cell(ws, row, col, value="", fill=None, font=None, align=None, border=None):
    """Write a cell and optionally apply styles. Returns the cell."""
    c = ws.cell(row=row, column=col, value=value)
    if fill:   c.fill      = fill
    if font:   c.font      = font
    if align:  c.alignment = align
    if border: c.border    = border
    return c


def _border_row(ws, row, num_cols=8):
    """Apply thin border to every cell in a row."""
    for col in range(1, num_cols + 1):
        ws.cell(row=row, column=col).border = _THIN


def generate_comparison_excel(data: dict) -> bytes:
    """
    Build an Excel workbook from the AI result dict and return raw bytes.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparison Report"

    # Column widths (A–H)
    widths = [38, 28, 28, 10, 12, 55, 4, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    row = 1

    # ── WARNING ─────────────────────────────────────────────────────────────
    warning = str(data.get("warning") or "").strip()
    if warning and warning.lower() not in ("none", "n/a", "null", "false", "no warning"):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _cell(ws, row, 1, f"⚠ WARNING: {warning}",
              fill=PatternFill("solid", fgColor="FEF3C7"),
              font=_RED_BOLD_BIG, align=_CENTER, border=_THIN)
        ws.row_dimensions[row].height = 28
        row += 2

    # ── SECTION 1: Promotions ────────────────────────────────────────────────
    promotions = data.get("section1_promotions") or []
    if promotions:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        _cell(ws, row, 1, "PROMOTIONS — Book by / Stay Period",
              fill=_LGRAY, font=_BLACK_BOLD, align=_LEFT)
        row += 1

        headers = ["Condition", "Contract 1", "Contract 2", "Analysis"]
        for col_i, h in enumerate(headers, start=1):
            _cell(ws, row, col_i, h, fill=_LGRAY, font=_BLACK_BOLD,
                  align=_CENTER, border=_THIN)
        row += 1

        for p in promotions:
            _cell(ws, row, 1, p.get("condition", ""), font=_BLACK_BOLD, align=_LEFT, border=_THIN)
            _cell(ws, row, 2, p.get("promo1", ""),    font=_BLACK_NORM, align=_CENTER, border=_THIN)
            _cell(ws, row, 3, p.get("promo2", ""),    font=_BLACK_NORM, align=_CENTER, border=_THIN)
            _cell(ws, row, 4, p.get("diff", ""),      font=_BLACK_NORM, align=_LEFT, border=_THIN)
            row += 1
        row += 1

    # ── SECTION 2: Hotel / Periods / Prices ──────────────────────────────────
    hotel_name = data.get("hotel_name", "Unknown Hotel")
    year_1     = data.get("year_1", "Yr 1")
    year_2     = data.get("year_2", "Yr 2")

    # Hotel name header row  (A=name, B=yr1, C=yr2, D=%, E=DIFF, F=Conditions, H=Changes)
    _cell(ws, row, 1, hotel_name, fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 2, year_1,     fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 3, year_2,     fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 4, "%",        fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 5, "DIFF",     fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 6, "Conditions / Notes",
          fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    _cell(ws, row, 8, "Contract Changes",
          fill=_YELLOW, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
    row += 1

    # "Room Name" sub-header
    _cell(ws, row, 1, "Room Name", fill=_LGRAY, font=_BLACK_BOLD, align=_LEFT, border=_THIN)
    for col_i in [2, 3, 4, 5, 6, 8]:
        _cell(ws, row, col_i, "", fill=_LGRAY, border=_THIN)
    row += 1

    periods = data.get("section2_periods") or []
    for period in periods:
        p1      = (period.get("period_1") or "").strip()
        p2      = (period.get("period_2") or "").strip()
        season  = (period.get("season_name") or "").strip()
        conds   = period.get("conditions") or []
        change  = (period.get("change_note") or "").strip()
        rooms   = period.get("rooms") or []

        # Period header row (blue)
        season_label = f"{season}  " if season else ""
        _cell(ws, row, 1, season_label,  fill=_BLUE, font=_WHITE_BOLD, align=_LEFT, border=_THIN)
        _cell(ws, row, 2, p1,            fill=_BLUE, font=_WHITE_BOLD, align=_CENTER, border=_THIN)
        _cell(ws, row, 3, p2,            fill=_BLUE, font=_WHITE_BOLD, align=_CENTER, border=_THIN)
        _cell(ws, row, 4, "%",           fill=_BLUE, font=_WHITE_BOLD, align=_CENTER, border=_THIN)
        _cell(ws, row, 5, "DIFF",        fill=_BLUE, font=_WHITE_BOLD, align=_CENTER, border=_THIN)

        # Conditions in col F (red background)
        if conds:
            cond_text = "\n".join(f"• {c}" for c in conds)
            cond_cell = _cell(ws, row, 6, cond_text,
                              fill=_RED, font=_WHITE_BOLD,
                              align=Alignment(horizontal="left", vertical="center", wrap_text=True),
                              border=_THIN)
            # Row height: estimate based on number of condition lines
            ws.row_dimensions[row].height = max(20, 16 * len(conds))

        # Change note in col H
        if change:
            _cell(ws, row, 8, change, font=Font(color="CC0000", bold=True, name="Calibri", size=10),
                  align=_LEFT, border=_THIN)

        row += 1

        # Room rows
        for rm in rooms:
            name       = rm.get("room_name") or ""
            price_1    = rm.get("price_1")
            price_2    = rm.get("price_2")
            diff_pct   = rm.get("diff_percent") or ""
            diff_amt   = rm.get("diff_amount")

            _cell(ws, row, 1, name,     font=_BLACK_NORM, align=_LEFT, border=_THIN)
            _cell(ws, row, 2, price_1,  font=_BLACK_NORM, align=_CENTER, border=_THIN)
            _cell(ws, row, 3, price_2,  font=_BLACK_NORM, align=_CENTER, border=_THIN)

            # Colour % cell: green if negative (cheaper), red if positive (more expensive)
            pct_font = _BLACK_NORM
            if diff_pct:
                if diff_pct.startswith("+"):
                    pct_font = Font(color="CC0000", bold=True, name="Calibri", size=10)
                elif diff_pct.startswith("-") or diff_pct.startswith("−"):
                    pct_font = Font(color="16A34A", bold=True, name="Calibri", size=10)
            _cell(ws, row, 4, diff_pct, font=pct_font, align=_CENTER, border=_THIN)

            # DIFF amount
            diff_font = _BLACK_NORM
            if isinstance(diff_amt, (int, float)):
                if diff_amt > 0:
                    diff_font = Font(color="CC0000", bold=True, name="Calibri", size=10)
                    diff_val  = f"+{diff_amt:,.0f}"
                elif diff_amt < 0:
                    diff_font = Font(color="16A34A", bold=True, name="Calibri", size=10)
                    diff_val  = f"{diff_amt:,.0f}"
                else:
                    diff_val = "0"
                _cell(ws, row, 5, diff_val, font=diff_font, align=_CENTER, border=_THIN)
            else:
                _cell(ws, row, 5, diff_amt or "", font=_BLACK_NORM, align=_CENTER, border=_THIN)

            row += 1

    row += 1  # spacer

    # ── SECTION 3: Conditions / Child Policy / E.B ───────────────────────────
    conditions = data.get("section3_conditions") or []
    if conditions:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        _cell(ws, row, 1, "CONDITIONS / CHILD POLICY / EARLY BIRD / BENEFITS",
              fill=_LGRAY, font=_BLACK_BOLD, align=_LEFT, border=_THIN)
        row += 1

        # Contract name labels
        _cell(ws, row, 1, "Topic",       fill=_LGRAY, font=_BLACK_BOLD, align=_LEFT,   border=_THIN)
        _cell(ws, row, 2, f"Contract {year_1}",
              fill=_LGRAY, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
        _cell(ws, row, 3, f"Contract {year_2}",
              fill=_LGRAY, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
        _cell(ws, row, 4, "Changes",     fill=_LGRAY, font=_BLACK_BOLD, align=_CENTER, border=_THIN)
        row += 1

        for cond in conditions:
            topic = cond.get("condition_type") or ""
            c1    = cond.get("contract_1") or ""
            c2    = cond.get("contract_2") or ""
            diff  = cond.get("diff") or ""

            # Colour the diff cell
            if diff.upper() == "SAME":
                diff_font = Font(color="16A34A", bold=True, name="Calibri", size=10)
            elif diff:
                diff_font = Font(color="CC0000", bold=True, name="Calibri", size=10)
            else:
                diff_font = _BLACK_NORM

            _cell(ws, row, 1, topic, font=_BLACK_BOLD, align=_LEFT,   border=_THIN)
            _cell(ws, row, 2, c1,    font=_BLACK_NORM, align=_LEFT,   border=_THIN)
            _cell(ws, row, 3, c2,    font=_BLACK_NORM, align=_LEFT,   border=_THIN)
            _cell(ws, row, 4, diff,  font=diff_font,   align=_CENTER, border=_THIN)

            # Row height
            max_lines = max(c1.count("\n"), c2.count("\n"), 0) + 1
            ws.row_dimensions[row].height = max(18, 15 * max_lines)
            row += 1

    # ── RECOMMENDATION ───────────────────────────────────────────────────────
    recommendation = str(data.get("recommendation") or "").strip()
    if recommendation:
        row += 1
        is_positive = "✅" in recommendation
        rec_fill = PatternFill("solid", fgColor="D1FAE5") if is_positive else PatternFill("solid", fgColor="FEF9C3")
        rec_font = Font(color="065F46" if is_positive else "92400E",
                        bold=True, name="Calibri", size=12)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _cell(ws, row, 1, recommendation,
              fill=rec_fill, font=rec_font, align=_CENTER, border=_THIN)
        ws.row_dimensions[row].height = 28

    # ── Save ─────────────────────────────────────────────────────────────────
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

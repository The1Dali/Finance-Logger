import sys
import csv
import os
from datetime import datetime

FINANCE_FILE = os.environ.get("FINANCE_FILE", "finance.csv")
LOG_FILE     = os.environ.get("LOG_FILE",     "log.csv")
EXPORT_FILE  = os.environ.get("EXPORT_FILE",  "finance.xlsx")

EXIT_OK            = 0
EXIT_BAD_ARGS      = 1
EXIT_INVALID_INPUT = 2
EXIT_NOT_FOUND     = 3
EXIT_FILE_ERROR    = 4


def print_help():
    print(
        "\nUsage: python finance.py <operation>\n"
        "\nOperations:\n"
        "  +  <name> <value>   Add <value> to an item called <name>. Creates it if it doesn't exist.\n"
        "  -  <name>           Remove the item called <name>.\n"
        "  clear               Wipe all entries from the finance file.\n"
        "  list                Print all current entries to the terminal.\n"
        "  export              Export finance and log data to a formatted .xlsx file.\n"
        "  --help, -h          Show this help message.\n"
        "\nEnvironment Variables:\n"
        f"  FINANCE_FILE        Path to the finance CSV file  (default: {FINANCE_FILE})\n"
        f"  LOG_FILE            Path to the log CSV file      (default: {LOG_FILE})\n"
        f"  EXPORT_FILE         Path for the exported .xlsx   (default: {EXPORT_FILE})\n"
        "\nExit Codes:\n"
        "  0  Success\n"
        "  1  Bad or missing arguments\n"
        "  2  Invalid input value\n"
        "  3  Item not found\n"
        "  4  File error\n"
    )


def main():
    try:
        with open(FINANCE_FILE, "r") as file:
            reader = csv.reader(file, delimiter="-")
            rows = [[cell.strip() for cell in row] for row in reader]
    except FileNotFoundError:
        rows = []
        try:
            with open(FINANCE_FILE, "w"):
                pass
        except OSError as e:
            print(f"Error: could not create '{FINANCE_FILE}': {e}", file=sys.stderr)
            sys.exit(EXIT_FILE_ERROR)

    try:
        match sys.argv[1]:
            case "--help" | "-h":
                print_help()
                sys.exit(EXIT_OK)

            case "+" if len(sys.argv) > 3:
                add(sys.argv[2], sys.argv[3], rows)

            case "+":
                print("Usage: python finance.py + <name> <value>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "-" if len(sys.argv) > 2:
                remove(sys.argv[2], rows)

            case "-":
                print("Usage: python finance.py - <name>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "clear":
                flash()

            case "list":
                list_entries(rows)

            case "export":
                export_xlsx(rows)

            case _:
                print(
                    "Unknown operation. Run 'python finance.py --help' for usage.",
                    file=sys.stderr
                )
                sys.exit(EXIT_BAD_ARGS)

    except IndexError:
        print(
            "No operation specified. Run 'python finance.py --help' for usage.",
            file=sys.stderr
        )
        sys.exit(EXIT_BAD_ARGS)

    sys.exit(EXIT_OK)


def log(action, name="", value=""):
    write_header = not os.path.exists(LOG_FILE)

    date_str   = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    name_fmt   = str(name).ljust(45)
    action_fmt = str(action).center(12)
    value_fmt  = (f"{float(value):.3f}" if value != "" else "").center(20)
    date_fmt   = date_str.rjust(19)

    header = [
        "Item".ljust(45),
        "Action".center(12),
        "Value".center(20),
        "Date".rjust(19),
    ]
    row = [name_fmt, action_fmt, value_fmt, date_fmt]

    try:
        with open(LOG_FILE, "a", newline="") as f:
            writer = csv.writer(f, delimiter="-")
            if write_header:
                writer.writerow(header)
            writer.writerow(row)
    except OSError as e:
        print(f"Warning: could not write to log file '{LOG_FILE}': {e}", file=sys.stderr)


def write(rows):
    try:
        with open(FINANCE_FILE, "w", newline="") as file:
            writer = csv.writer(file, delimiter="-")
            for row in rows:
                try:
                    pct = f"{float(row[2]):.3f}%"
                except (ValueError, IndexError):
                    pct = "0.000%"
                f_row = [
                    row[0].ljust(45),
                    (f"{float(row[1]):.3f}").center(20),
                    pct.rjust(20)
                ]
                writer.writerow(f_row)
    except OSError as e:
        print(f"Error: could not write to '{FINANCE_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)


def refresh(rows):
    total = 0
    total_index = -1

    for i in range(len(rows)):
        if rows[i][0].strip().upper() == "TOTAL":
            total_index = i
        else:
            total += float(rows[i][1])

    if total_index >= 0:
        rows[total_index][1] = str(total)
    elif len(rows) > 0:
        rows.append(["TOTAL", str(total), "0.0"])

    for i in range(len(rows)):
        try:
            rows[i][2] = str(float(rows[i][1]) / total * 100)
        except ZeroDivisionError:
            rows[i][2] = "0.0"

    rows.sort(key=lambda row: float(row[2]), reverse=True)
    write(rows)


def add(name, value, rows):
    try:
        value = float(value)
    except ValueError:
        print(f"Error: '{value}' is not a valid number.", file=sys.stderr)
        sys.exit(EXIT_INVALID_INPUT)

    found = False
    for i in range(len(rows)):
        if rows[i][0].strip().upper() == name.upper():
            old_value = float(rows[i][1])
            rows[i][1] = str(old_value + value)
            found = True
            print(f"Updated '{name}': {old_value:.3f} → {old_value + value:.3f}")
            break

    if not found:
        rows.append([name, str(value), "0.0"])
        print(f"Added new item '{name}' with value {value:.3f}")

    log("ADD", name, value)
    refresh(rows)


def remove(name, rows):
    for i in range(len(rows)):
        if rows[i][0].strip().upper() == name.upper():
            rows.pop(i)
            print(f"'{name}' was removed.")
            log("REMOVE", name)
            break
    else:
        print(f"'{name}' was not found.", file=sys.stderr)
        sys.exit(EXIT_NOT_FOUND)

    refresh(rows)


def flash():
    try:
        with open(FINANCE_FILE, "w"):
            pass
    except OSError as e:
        print(f"Error: could not clear '{FINANCE_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)
    log("CLEAR", "", "")
    print(f"'{FINANCE_FILE}' has been cleared.")


def list_entries(rows):
    if not rows:
        print("No entries found.")
        return

    print(f"\n{'Item':<45} {'Value':>15} {'Share':>10}")
    print("-" * 72)
    for row in rows:
        try:
            name  = row[0].strip()
            value = f"{float(row[1]):.3f}"
            pct   = f"{float(row[2]):.3f}%"
        except (ValueError, IndexError):
            continue
        label = f"[{name}]" if name.upper() == "TOTAL" else name
        print(f"{label:<45} {value:>15} {pct:>10}")
    print()


def export_xlsx(rows):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import (
            Font, PatternFill, Alignment, Border, Side, GradientFill
        )
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Error: openpyxl is required for export. Run: pip install openpyxl", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)

    wb = Workbook()

    # ── shared style helpers ──────────────────────────────────────────────────
    FONT_NAME   = "Arial"
    CLR_HEADER_BG  = "2E4057"  
    CLR_HEADER_FG  = "FFFFFF"   
    CLR_TOTAL_BG   = "D9E1F2"   
    CLR_TOTAL_FG   = "1F3864"  
    CLR_ALT_ROW    = "F2F5FB"   
    CLR_WHITE      = "FFFFFF"
    CLR_ACCENT     = "4472C4"   

    thin  = Side(style="thin",   color=CLR_ACCENT)
    thick = Side(style="medium", color=CLR_ACCENT)

    def header_style(cell, text, align="left"):
        cell.value     = text
        cell.font      = Font(name=FONT_NAME, bold=True, color=CLR_HEADER_FG, size=11)
        cell.fill      = PatternFill("solid", fgColor=CLR_HEADER_BG)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = Border(bottom=thick)

    def total_style(cell, align="left"):
        cell.font      = Font(name=FONT_NAME, bold=True, color=CLR_TOTAL_FG, size=10)
        cell.fill      = PatternFill("solid", fgColor=CLR_TOTAL_BG)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = Border(top=thick, bottom=thick)

    def body_style(cell, align="left", alt=False):
        cell.font      = Font(name=FONT_NAME, size=10)
        cell.fill      = PatternFill("solid", fgColor=CLR_ALT_ROW if alt else CLR_WHITE)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = Border(
            left=Side(style="hair", color="CCCCCC"),
            right=Side(style="hair", color="CCCCCC"),
            bottom=Side(style="hair", color="CCCCCC")
        )

    # ── Finance sheet ─────────────────────────────────────────────────────────
    fs = wb.active
    fs.title = "Finance"
    fs.freeze_panes = "A2"   
    fs.row_dimensions[1].height = 28

    header_style(fs["A1"], "Item")
    header_style(fs["B1"], "Value",  align="right")
    header_style(fs["C1"], "Share %", align="right")

    fs.column_dimensions["A"].width = 38
    fs.column_dimensions["B"].width = 18
    fs.column_dimensions["C"].width = 14

    data_rows  = [r for r in rows if r[0].strip().upper() != "TOTAL"]
    total_rows = [r for r in rows if r[0].strip().upper() == "TOTAL"]
    ordered    = data_rows + total_rows

    total_excel_row = len(data_rows) + 2

    for idx, row in enumerate(ordered):
        excel_row  = idx + 2
        is_total   = row[0].strip().upper() == "TOTAL"
        alt        = (idx % 2 == 1) and not is_total

        try:
            val = float(row[1])
        except ValueError:
            val = 0.0

        name_cell  = fs.cell(row=excel_row, column=1, value=row[0].strip())
        value_cell = fs.cell(row=excel_row, column=2, value=val)
        pct_cell   = fs.cell(row=excel_row, column=3)

        if is_total:
            pct_cell.value  = None
            total_style(name_cell,  align="left")
            total_style(value_cell, align="right")
            total_style(pct_cell,   align="right")
            value_cell.value = f"=SUM(B2:B{total_excel_row - 1})"
        else:
            pct_cell.value = f"=IFERROR(B{excel_row}/B{total_excel_row}, 0)"
            body_style(name_cell,  align="left",  alt=alt)
            body_style(value_cell, align="right", alt=alt)
            body_style(pct_cell,   align="right", alt=alt)

        value_cell.number_format = '#,##0.000'
        pct_cell.number_format   = '0.00%'

    # ── Log sheet ─────────────────────────────────────────────────────────────
    ls = wb.create_sheet(title="Log")
    ls.freeze_panes = "A2"
    ls.row_dimensions[1].height = 28

    header_style(ls["A1"], "Item")
    header_style(ls["B1"], "Action",  align="center")
    header_style(ls["C1"], "Value",   align="right")
    header_style(ls["D1"], "Date",    align="center")

    ls.column_dimensions["A"].width = 38
    ls.column_dimensions["B"].width = 12
    ls.column_dimensions["C"].width = 18
    ls.column_dimensions["D"].width = 22

    ACTION_COLORS = {
        "ADD":    "E2EFDA",   
        "REMOVE": "FCE4D6",  
        "CLEAR":  "FFF2CC",  
    }

    log_rows = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", newline="") as f:
                reader = csv.reader(f, delimiter="-")
                next(reader, None)   
                log_rows = list(reader)
        except OSError:
            pass

    for idx, row in enumerate(log_rows):
        excel_row = idx + 2
        alt       = idx % 2 == 1

        try:
            name   = row[0].strip() if len(row) > 0 else ""
            action = row[1].strip() if len(row) > 1 else ""
            value  = row[2].strip() if len(row) > 2 else ""
            date   = row[3].strip() if len(row) > 3 else ""
        except IndexError:
            continue

        action_bg = ACTION_COLORS.get(action.upper(), CLR_WHITE)

        val_num = None
        try:
            val_num = float(value) if value else None
        except ValueError:
            pass

        name_cell   = ls.cell(row=excel_row, column=1, value=name)
        action_cell = ls.cell(row=excel_row, column=2, value=action)
        value_cell  = ls.cell(row=excel_row, column=3, value=val_num)
        date_cell   = ls.cell(row=excel_row, column=4, value=date)

        body_style(name_cell,   align="left",   alt=alt)
        body_style(date_cell,   align="center", alt=alt)
        body_style(value_cell,  align="right",  alt=alt)

        action_cell.font      = Font(name=FONT_NAME, bold=True, size=10,
                                     color=CLR_TOTAL_FG)
        action_cell.fill      = PatternFill("solid", fgColor=action_bg)
        action_cell.alignment = Alignment(horizontal="center", vertical="center")
        action_cell.border    = Border(
            left=Side(style="hair",  color="CCCCCC"),
            right=Side(style="hair", color="CCCCCC"),
            bottom=Side(style="hair", color="CCCCCC")
        )

        if val_num is not None:
            value_cell.number_format = '#,##0.000'

    # ── save ─────────────────────────────────────────────────────────────────
    try:
        wb.save(EXPORT_FILE)
        print(f"Exported to '{EXPORT_FILE}' — Finance sheet ({len(data_rows)} items) + Log sheet ({len(log_rows)} entries).")
    except OSError as e:
        print(f"Error: could not save '{EXPORT_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)


main()
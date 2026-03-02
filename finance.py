import sys
import csv
import os
import subprocess
import shutil
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
                print("Unknown operation. Run 'python finance.py --help' for usage.", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)
    except IndexError:
        print("No operation specified. Run 'python finance.py --help' for usage.", file=sys.stderr)
        sys.exit(EXIT_BAD_ARGS)

    sys.exit(EXIT_OK)


def log(action, name="", value=""):
    write_header = not os.path.exists(LOG_FILE)
    date_str   = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    name_fmt   = str(name).ljust(45)
    action_fmt = str(action).center(12)
    value_fmt  = (f"{float(value):.3f}" if value != "" else "").center(20)
    date_fmt   = date_str.rjust(19)
    header = ["Item".ljust(45), "Action".center(12), "Value".center(20), "Date".rjust(19)]
    row    = [name_fmt, action_fmt, value_fmt, date_fmt]
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
            print(f"Updated '{name}': {old_value:.3f} -> {old_value + value:.3f}")
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
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("Error: openpyxl is required for export. Run: pip install openpyxl", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)

    wb = Workbook()

    FONT         = "Arial"
    TITLE_BG     = "1B2A6B"
    TITLE_FG     = "FFFFFF"
    HDR_BG       = "2E5FA3"
    HDR_FG       = "FFFFFF"
    BODY_WHITE   = "FFFFFF"
    BODY_ALT     = "EBF1FA"
    TOTAL_BG     = "BDD0EE"
    TOTAL_FG     = "0D1B5E"
    BORDER_CLR   = "B8CCE4"
    TAB_FIN      = "1B2A6B"
    TAB_LOG      = "2E5FA3"
    ADD_FG       = "155724"
    ADD_BG       = "D4EDDA"
    REMOVE_FG    = "721C24"
    REMOVE_BG    = "F8D7DA"
    CLEAR_FG     = "856404"
    CLEAR_BG     = "FFF3CD"

    def solid(color):
        return PatternFill("solid", fgColor=color)

    def hborder():
        s = Side(style="hair", color=BORDER_CLR)
        return Border(left=s, right=s, top=s, bottom=s)

    def apply_title(cell, text):
        cell.value     = text
        cell.font      = Font(name=FONT, bold=True, size=15, color=TITLE_FG)
        cell.fill      = solid(TITLE_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def apply_header(cell, text, align="left"):
        cell.value     = text
        cell.font      = Font(name=FONT, bold=True, size=10, color=HDR_FG)
        cell.fill      = solid(HDR_BG)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = Border(bottom=Side(style="medium", color=TITLE_BG))

    def apply_total(cell, align="left"):
        cell.font      = Font(name=FONT, bold=True, size=10, color=TOTAL_FG)
        cell.fill      = solid(TOTAL_BG)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = Border(
            top=Side(style="medium", color=TITLE_BG),
            bottom=Side(style="medium", color=TITLE_BG),
            left=Side(style="hair", color=BORDER_CLR),
            right=Side(style="hair", color=BORDER_CLR),
        )

    def apply_body(cell, align="left", alt=False):
        cell.font      = Font(name=FONT, size=10, color="1A1A2E")
        cell.fill      = solid(BODY_ALT if alt else BODY_WHITE)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = hborder()

    def apply_badge(cell, action):
        badges = {
            "ADD":    (ADD_FG,    ADD_BG),
            "REMOVE": (REMOVE_FG, REMOVE_BG),
            "CLEAR":  (CLEAR_FG,  CLEAR_BG),
        }
        fg, bg = badges.get(action.upper(), ("333333", BODY_WHITE))
        cell.font      = Font(name=FONT, bold=True, size=9, color=fg)
        cell.fill      = solid(bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = hborder()

    fs = wb.active
    fs.title = "Finance"
    fs.sheet_properties.tabColor = TAB_FIN
    fs.freeze_panes = "A3"
    fs.sheet_view.showGridLines = False
    fs.row_dimensions[1].height = 38
    fs.row_dimensions[2].height = 22
    fs.column_dimensions["A"].width = 36
    fs.column_dimensions["B"].width = 18
    fs.column_dimensions["C"].width = 14

    fs.merge_cells("A1:C1")
    apply_title(fs["A1"], "IEEE Logger")
    apply_header(fs["A2"], "Item",    align="left")
    apply_header(fs["B2"], "Amount",  align="right")
    apply_header(fs["C2"], "Share %", align="right")

    data_rows       = [r for r in rows if r[0].strip().upper() != "TOTAL"]
    total_rows      = [r for r in rows if r[0].strip().upper() == "TOTAL"]
    ordered         = data_rows + total_rows
    total_excel_row = len(data_rows) + 3

    for idx, row in enumerate(ordered):
        excel_row = idx + 3
        is_total  = row[0].strip().upper() == "TOTAL"
        alt       = (idx % 2 == 1) and not is_total
        fs.row_dimensions[excel_row].height = 20
        try:
            val = float(row[1])
        except ValueError:
            val = 0.0
        name_cell  = fs.cell(row=excel_row, column=1, value=row[0].strip())
        value_cell = fs.cell(row=excel_row, column=2, value=val)
        pct_cell   = fs.cell(row=excel_row, column=3)
        if is_total:
            value_cell.value = f"=SUM(B3:B{total_excel_row - 1})"
            pct_cell.value   = None
            apply_total(name_cell,  align="left")
            apply_total(value_cell, align="right")
            apply_total(pct_cell,   align="right")
        else:
            pct_cell.value = f"=IFERROR(B{excel_row}/B{total_excel_row},0)"
            apply_body(name_cell,  align="left",  alt=alt)
            apply_body(value_cell, align="right", alt=alt)
            apply_body(pct_cell,   align="right", alt=alt)
        value_cell.number_format = "#,##0.000"
        pct_cell.number_format   = "0.00%"

    ls = wb.create_sheet(title="Log")
    ls.sheet_properties.tabColor = TAB_LOG
    ls.freeze_panes = "A3"
    ls.sheet_view.showGridLines = False
    ls.row_dimensions[1].height = 38
    ls.row_dimensions[2].height = 22
    ls.column_dimensions["A"].width = 36
    ls.column_dimensions["B"].width = 11
    ls.column_dimensions["C"].width = 16
    ls.column_dimensions["D"].width = 22

    ls.merge_cells("A1:D1")
    apply_title(ls["A1"], "ACTIVITY LOG")
    apply_header(ls["A2"], "Item",   align="left")
    apply_header(ls["B2"], "Action", align="center")
    apply_header(ls["C2"], "Amount", align="right")
    apply_header(ls["D2"], "Date",   align="center")

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
        excel_row = idx + 3
        alt       = idx % 2 == 1
        ls.row_dimensions[excel_row].height = 20
        try:
            name   = row[0].strip() if len(row) > 0 else ""
            action = row[1].strip() if len(row) > 1 else ""
            value  = row[2].strip() if len(row) > 2 else ""
            date   = row[3].strip() if len(row) > 3 else ""
        except IndexError:
            continue
        val_num = None
        try:
            val_num = float(value) if value else None
        except ValueError:
            pass
        name_cell   = ls.cell(row=excel_row, column=1, value=name)
        action_cell = ls.cell(row=excel_row, column=2, value=action)
        value_cell  = ls.cell(row=excel_row, column=3, value=val_num)
        date_cell   = ls.cell(row=excel_row, column=4, value=date)
        apply_body(name_cell,  align="left",   alt=alt)
        apply_body(value_cell, align="right",  alt=alt)
        apply_body(date_cell,  align="center", alt=alt)
        apply_badge(action_cell, action)
        if val_num is not None:
            value_cell.number_format = "#,##0.000"

    try:
        wb.save(EXPORT_FILE)
    except OSError as e:
        print(f"Error: could not save '{EXPORT_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)

    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo:
        abs_export = os.path.abspath(EXPORT_FILE)
        export_dir = os.path.dirname(abs_export) or "."
        result = subprocess.run(
            [lo, "--headless", "--convert-to", "xlsx", "--outdir", export_dir, abs_export],
            capture_output=True
        )
        if result.returncode != 0:
            print("Warning: LibreOffice conversion step failed.", file=sys.stderr)

    print(f"Exported to '{EXPORT_FILE}' - Finance sheet ({len(data_rows)} items) + Log sheet ({len(log_rows)} entries).")


main()
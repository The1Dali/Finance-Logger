import sys
import csv
import os
import subprocess
import shutil
from datetime import datetime

FINANCE_FILE = os.environ.get("FINANCE_FILE", "finance.csv")
CONFIG_FILE  = os.environ.get("CONFIG_FILE",  "config.csv")
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
        "  income <description> <amount>    Add an income entry.\n"
        "  expense <description> <amount>   Add an expense entry.\n"
        "  remove income <id>               Remove an income entry by ID.\n"
        "  remove expense <id>              Remove an expense entry by ID.\n"
        "  balance <amount>                 Set the opening balance.\n"
        "  list                             Print a summary to the terminal.\n"
        "  export                           Export to a formatted .xlsx file.\n"
        "  clear                            Wipe all entries.\n"
        "  --help, -h                       Show this help message.\n"
        "\nEnvironment Variables:\n"
        f"  FINANCE_FILE   Path to the finance CSV  (default: {FINANCE_FILE})\n"
        f"  CONFIG_FILE    Path to the config CSV   (default: {CONFIG_FILE})\n"
        f"  LOG_FILE       Path to the log CSV      (default: {LOG_FILE})\n"
        f"  EXPORT_FILE    Path for the .xlsx       (default: {EXPORT_FILE})\n"
        "\nExit Codes:\n"
        "  0  Success\n"
        "  1  Bad or missing arguments\n"
        "  2  Invalid input value\n"
        "  3  Entry not found\n"
        "  4  File error\n"
    )


def main():
    rows = load_finance()

    try:
        match sys.argv[1]:
            case "--help" | "-h":
                print_help()
                sys.exit(EXIT_OK)

            case "income" if len(sys.argv) > 3:
                add_entry("INCOME", sys.argv[2], sys.argv[3], rows)

            case "income":
                print("Usage: python finance.py income <description> <amount>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "expense" if len(sys.argv) > 3:
                add_entry("EXPENSE", sys.argv[2], sys.argv[3], rows)

            case "expense":
                print("Usage: python finance.py expense <description> <amount>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "remove" if len(sys.argv) > 3:
                remove_entry(sys.argv[2], sys.argv[3], rows)

            case "remove":
                print("Usage: python finance.py remove <income|expense> <id>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "balance" if len(sys.argv) > 2:
                set_balance(sys.argv[2])

            case "balance":
                print("Usage: python finance.py balance <amount>", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "list":
                list_entries(rows)

            case "export":
                export_xlsx(rows)

            case "clear":
                clear_all()

            case _:
                print("Unknown operation. Run 'python finance.py --help' for usage.", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

    except IndexError:
        print("No operation specified. Run 'python finance.py --help' for usage.", file=sys.stderr)
        sys.exit(EXIT_BAD_ARGS)

    sys.exit(EXIT_OK)


def load_finance():
    try:
        with open(FINANCE_FILE, "r", newline="") as f:
            reader = csv.DictReader(f)
            return list(reader)
    except FileNotFoundError:
        try:
            with open(FINANCE_FILE, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["type", "description", "amount", "date"])
                writer.writeheader()
        except OSError as e:
            print(f"Error: could not create '{FINANCE_FILE}': {e}", file=sys.stderr)
            sys.exit(EXIT_FILE_ERROR)
        return []


def save_finance(rows):
    try:
        with open(FINANCE_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["type", "description", "amount", "date"])
            writer.writeheader()
            writer.writerows(rows)
    except OSError as e:
        print(f"Error: could not write to '{FINANCE_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)


def load_balance():
    try:
        with open(CONFIG_FILE, "r", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("key") == "opening_balance":
                    return float(row["value"])
    except (FileNotFoundError, ValueError):
        pass
    return 0.0


def save_balance(amount):
    try:
        with open(CONFIG_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["key", "value"])
            writer.writeheader()
            writer.writerow({"key": "opening_balance", "value": str(amount)})
    except OSError as e:
        print(f"Error: could not write to '{CONFIG_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)


def log_event(action, entry_type="", description="", amount=""):
    write_header = not os.path.exists(LOG_FILE)
    date_str     = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    try:
        with open(LOG_FILE, "a", newline="") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(["date", "action", "type", "description", "amount"])
            writer.writerow([date_str, action, entry_type, description, amount])
    except OSError as e:
        print(f"Warning: could not write to log file '{LOG_FILE}': {e}", file=sys.stderr)


def add_entry(entry_type, description, amount_str, rows):
    try:
        amount = float(amount_str)
    except ValueError:
        print(f"Error: '{amount_str}' is not a valid number.", file=sys.stderr)
        sys.exit(EXIT_INVALID_INPUT)

    date_str = datetime.now().strftime("%d/%m/%Y")
    rows.append({
        "type":        entry_type,
        "description": description,
        "amount":      str(amount),
        "date":        date_str,
    })
    save_finance(rows)
    log_event("ADD", entry_type, description, amount)
    print(f"Added {entry_type.lower()} '{description}': {amount:.3f} TND")


def remove_entry(entry_type, id_str, rows):
    entry_type = entry_type.upper()
    if entry_type not in ("INCOME", "EXPENSE"):
        print(f"Error: type must be 'income' or 'expense', got '{entry_type}'.", file=sys.stderr)
        sys.exit(EXIT_BAD_ARGS)

    try:
        target_id = int(id_str)
    except ValueError:
        print(f"Error: '{id_str}' is not a valid ID.", file=sys.stderr)
        sys.exit(EXIT_INVALID_INPUT)

    typed_rows = [r for r in rows if r["type"].upper() == entry_type]
    if target_id < 1 or target_id > len(typed_rows):
        print(f"Error: {entry_type.lower()} entry #{target_id} not found.", file=sys.stderr)
        sys.exit(EXIT_NOT_FOUND)

    target = typed_rows[target_id - 1]
    rows.remove(target)
    save_finance(rows)
    log_event("REMOVE", entry_type, target["description"], target["amount"])
    print(f"Removed {entry_type.lower()} #{target_id} '{target['description']}'.")


def set_balance(amount_str):
    try:
        amount = float(amount_str)
    except ValueError:
        print(f"Error: '{amount_str}' is not a valid number.", file=sys.stderr)
        sys.exit(EXIT_INVALID_INPUT)
    save_balance(amount)
    log_event("BALANCE", "", "Opening balance set", amount)
    print(f"Opening balance set to {amount:.3f} TND")


def list_entries(rows):
    opening   = load_balance()
    income    = [r for r in rows if r["type"].upper() == "INCOME"]
    expenses  = [r for r in rows if r["type"].upper() == "EXPENSE"]
    total_in  = sum(float(r["amount"]) for r in income)
    total_exp = sum(float(r["amount"]) for r in expenses)
    profit    = total_in - total_exp
    net       = opening + profit

    print(f"\n{'SUMMARY':=<45}")
    print(f"  {'Opening Balance':<30} {opening:>10.3f} TND")
    print(f"  {'Total Income':<30} {total_in:>10.3f} TND")
    print(f"  {'Total Expenses':<30} {total_exp:>10.3f} TND")
    print(f"  {'Profit':<30} {profit:>10.3f} TND")
    print(f"  {'Net Balance':<30} {net:>10.3f} TND")

    if income:
        print(f"\n{'INCOME':=<45}")
        print(f"  {'#':<5} {'Description':<28} {'Amount':>10}  {'Date'}")
        print(f"  {'-'*60}")
        for i, r in enumerate(income, 1):
            print(f"  {i:<5} {r['description']:<28} {float(r['amount']):>10.3f}  {r['date']}")

    if expenses:
        print(f"\n{'EXPENSES':=<45}")
        print(f"  {'#':<5} {'Description':<28} {'Amount':>10}  {'Date'}")
        print(f"  {'-'*60}")
        for i, r in enumerate(expenses, 1):
            print(f"  {i:<5} {r['description']:<28} {float(r['amount']):>10.3f}  {r['date']}")
    print()


def clear_all():
    try:
        with open(FINANCE_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["type", "description", "amount", "date"])
            writer.writeheader()
    except OSError as e:
        print(f"Error: could not clear '{FINANCE_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)
    log_event("CLEAR")
    print(f"All entries cleared.")


def export_xlsx(rows):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("Error: openpyxl is required. Run: pip install openpyxl", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)

    wb = Workbook()

    FONT       = "Arial"
    WHITE      = "FFFFFF"
    BODY_ALT   = "F4F6FB"
    BORDER_CLR = "D0D9EC"

    BLUE_DARK  = "1A56A0"
    BLUE_MID   = "2E75C8"
    BLUE_LIGHT = "D6E4F7"
    BLUE_FG    = "0D2B5E"

    GREEN_DARK  = "1A7A3A"
    GREEN_MID   = "28A745"
    GREEN_LIGHT = "D4EDDA"
    GREEN_FG    = "0B3D1E"

    RED_DARK  = "B52525"
    RED_MID   = "E03535"
    RED_LIGHT = "F8D7DA"
    RED_FG    = "5C0A0A"

    GREY_HDR = "C8D4E8"

    def solid(color):
        return PatternFill("solid", fgColor=color)

    def border(color=BORDER_CLR, bottom_color=None, bottom_weight="hair"):
        s = Side(style="hair", color=color)
        b = Side(style=bottom_weight, color=bottom_color or color)
        return Border(left=s, right=s, top=s, bottom=b)

    def title_cell(cell, text, bg, fg=WHITE):
        cell.value     = text
        cell.font      = Font(name=FONT, bold=True, size=14, color=fg)
        cell.fill      = solid(bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = Border(bottom=Side(style="medium", color=bg))

    def header_cell(cell, text, bg, fg=WHITE, align="center"):
        cell.value     = text
        cell.font      = Font(name=FONT, bold=True, size=9, color=fg)
        cell.fill      = solid(bg)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = border(bg, bottom_color=bg, bottom_weight="thin")

    def summary_label(cell, text, bg, fg):
        cell.value     = text
        cell.font      = Font(name=FONT, bold=True, size=10, color=fg)
        cell.fill      = solid(bg)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border    = border(BORDER_CLR)

    def summary_value(cell, bg, fg):
        cell.font      = Font(name=FONT, bold=True, size=10, color=fg)
        cell.fill      = solid(bg)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border    = border(BORDER_CLR)
        cell.number_format = '#,##0.000 "TND"'

    def body_cell(cell, value, align="left", alt=False, fmt=None):
        cell.value     = value
        cell.font      = Font(name=FONT, size=10, color="1A1A2E")
        cell.fill      = solid(BODY_ALT if alt else WHITE)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = border()
        if fmt:
            cell.number_format = fmt

    income_rows  = [r for r in rows if r["type"].upper() == "INCOME"]
    expense_rows = [r for r in rows if r["type"].upper() == "EXPENSE"]
    opening      = load_balance()

    n_inc  = len(income_rows)
    n_exp  = len(expense_rows)

    inc_data_start  = 4
    inc_data_end    = inc_data_start + n_inc - 1 if n_inc else inc_data_start
    exp_data_start  = 4
    exp_data_end    = exp_data_start + n_exp - 1 if n_exp else exp_data_start

    inc_sum_ref  = f"'Income'!C{inc_data_start}:C{inc_data_end}" if n_inc else "'Income'!C4:C4"
    exp_sum_ref  = f"'Expenses'!C{exp_data_start}:C{exp_data_end}" if n_exp else "'Expenses'!C4:C4"

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ss = wb.active
    ss.title = "Summary"
    ss.sheet_properties.tabColor = BLUE_DARK
    ss.sheet_view.showGridLines  = False
    ss.column_dimensions["A"].width = 28
    ss.column_dimensions["B"].width = 20

    ss.row_dimensions[1].height = 38
    ss.merge_cells("A1:B1")
    title_cell(ss["A1"], "SUMMARY", BLUE_DARK)

    ss.row_dimensions[2].height = 6

    summary_items = [
        ("Opening Balance (TND)", str(opening),          BLUE_LIGHT,  BLUE_FG,   BLUE_DARK),
        ("Total Income (TND)",    f"=SUM({inc_sum_ref})", GREEN_LIGHT, GREEN_FG,  GREEN_DARK),
        ("Total Expenses (TND)",  f"=SUM({exp_sum_ref})", RED_LIGHT,   RED_FG,    RED_DARK),
        ("Profit (TND)",          "=B4-B5",               BLUE_LIGHT,  BLUE_FG,   BLUE_DARK),
        ("Net Balance (TND)",     "=B3+B6",               BLUE_MID,    WHITE,     BLUE_DARK),
    ]

    for i, (label, formula, bg, fg, border_col) in enumerate(summary_items):
        row = i + 3
        ss.row_dimensions[row].height = 22

        lc = ss.cell(row=row, column=1)
        lc.value     = label
        lc.font      = Font(name=FONT, bold=True, size=10, color=fg)
        lc.fill      = solid(bg)
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border    = Border(
            left=Side(style="medium", color=border_col),
            top=Side(style="hair", color=BORDER_CLR),
            bottom=Side(style="hair", color=BORDER_CLR),
            right=Side(style="hair", color=BORDER_CLR),
        )

        vc = ss.cell(row=row, column=2)
        if isinstance(formula, str) and formula.startswith("="):
            vc.value = formula
        else:
            vc.value = float(formula)
        vc.font         = Font(name=FONT, bold=(i == 4), size=10, color=fg)
        vc.fill         = solid(bg)
        vc.alignment    = Alignment(horizontal="right", vertical="center")
        vc.number_format = '#,##0.000 "TND"'
        vc.border       = Border(
            right=Side(style="medium", color=border_col),
            top=Side(style="hair", color=BORDER_CLR),
            bottom=Side(style="hair", color=BORDER_CLR),
            left=Side(style="hair", color=BORDER_CLR),
        )

    # ── Income sheet ──────────────────────────────────────────────────────────
    ws_inc = wb.create_sheet("Income")
    ws_inc.sheet_properties.tabColor = GREEN_DARK
    ws_inc.sheet_view.showGridLines  = False
    ws_inc.column_dimensions["A"].width = 6
    ws_inc.column_dimensions["B"].width = 36
    ws_inc.column_dimensions["C"].width = 18
    ws_inc.column_dimensions["D"].width = 16

    ws_inc.row_dimensions[1].height = 38
    ws_inc.row_dimensions[2].height = 6
    ws_inc.row_dimensions[3].height = 20
    ws_inc.merge_cells("A1:D1")
    title_cell(ws_inc["A1"], "INCOME", GREEN_DARK)
    ws_inc.freeze_panes = "A4"

    for col, (text, align) in enumerate([("ID","center"),("Description","left"),("Amount (TND)","right"),("Date","center")], 1):
        header_cell(ws_inc.cell(row=3, column=col), text, GREEN_MID, align=align)

    for i, r in enumerate(income_rows):
        row = i + inc_data_start
        alt = i % 2 == 1
        ws_inc.row_dimensions[row].height = 20
        body_cell(ws_inc.cell(row=row, column=1), f"=ROW()-{inc_data_start - 1}", align="center", alt=alt)
        body_cell(ws_inc.cell(row=row, column=2), r["description"],              align="left",   alt=alt)
        body_cell(ws_inc.cell(row=row, column=3), float(r["amount"]),            align="right",  alt=alt, fmt='#,##0.000 "TND"')
        body_cell(ws_inc.cell(row=row, column=4), r["date"],                     align="center", alt=alt)

    # ── Expenses sheet ────────────────────────────────────────────────────────
    ws_exp = wb.create_sheet("Expenses")
    ws_exp.sheet_properties.tabColor = RED_DARK
    ws_exp.sheet_view.showGridLines  = False
    ws_exp.column_dimensions["A"].width = 6
    ws_exp.column_dimensions["B"].width = 36
    ws_exp.column_dimensions["C"].width = 18
    ws_exp.column_dimensions["D"].width = 16

    ws_exp.row_dimensions[1].height = 38
    ws_exp.row_dimensions[2].height = 6
    ws_exp.row_dimensions[3].height = 20
    ws_exp.merge_cells("A1:D1")
    title_cell(ws_exp["A1"], "EXPENSES", RED_DARK)
    ws_exp.freeze_panes = "A4"

    for col, (text, align) in enumerate([("ID","center"),("Description","left"),("Amount (TND)","right"),("Date","center")], 1):
        header_cell(ws_exp.cell(row=3, column=col), text, RED_MID, align=align)

    for i, r in enumerate(expense_rows):
        row = i + exp_data_start
        alt = i % 2 == 1
        ws_exp.row_dimensions[row].height = 20
        body_cell(ws_exp.cell(row=row, column=1), f"=ROW()-{exp_data_start - 1}", align="center", alt=alt)
        body_cell(ws_exp.cell(row=row, column=2), r["description"],               align="left",   alt=alt)
        body_cell(ws_exp.cell(row=row, column=3), float(r["amount"]),             align="right",  alt=alt, fmt='#,##0.000 "TND"')
        body_cell(ws_exp.cell(row=row, column=4), r["date"],                      align="center", alt=alt)

    # ── Log sheet ─────────────────────────────────────────────────────────────
    ws_log = wb.create_sheet("Log")
    ws_log.sheet_properties.tabColor = "555555"
    ws_log.sheet_view.showGridLines  = False
    ws_log.column_dimensions["A"].width = 20
    ws_log.column_dimensions["B"].width = 10
    ws_log.column_dimensions["C"].width = 10
    ws_log.column_dimensions["D"].width = 30
    ws_log.column_dimensions["E"].width = 16

    ws_log.row_dimensions[1].height = 38
    ws_log.row_dimensions[2].height = 6
    ws_log.row_dimensions[3].height = 20
    ws_log.merge_cells("A1:E1")
    title_cell(ws_log["A1"], "ACTIVITY LOG", "2C2C2C")
    ws_log.freeze_panes = "A4"

    HDR_GREY = "4A4A4A"
    for col, (text, align) in enumerate([("Date","center"),("Action","center"),("Type","center"),("Description","left"),("Amount","right")], 1):
        header_cell(ws_log.cell(row=3, column=col), text, HDR_GREY, align=align)

    log_rows = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", newline="") as f:
                reader = csv.DictReader(f)
                log_rows = list(reader)
        except OSError:
            pass

    ACTION_BG = {"ADD": "D4EDDA", "REMOVE": "F8D7DA", "CLEAR": "FFF3CD", "BALANCE": "D6E4F7"}
    ACTION_FG = {"ADD": "155724", "REMOVE": "721C24", "CLEAR": "856404", "BALANCE": "0D2B5E"}

    for i, r in enumerate(log_rows):
        row = i + 4
        alt = i % 2 == 1
        ws_log.row_dimensions[row].height = 20
        action = r.get("action", "").upper()
        abg = ACTION_BG.get(action, WHITE)
        afg = ACTION_FG.get(action, "333333")

        body_cell(ws_log.cell(row=row, column=1), r.get("date",""),        align="center", alt=alt)
        ac = ws_log.cell(row=row, column=2, value=action)
        ac.font      = Font(name=FONT, bold=True, size=9, color=afg)
        ac.fill      = solid(abg)
        ac.alignment = Alignment(horizontal="center", vertical="center")
        ac.border    = border()
        body_cell(ws_log.cell(row=row, column=3), r.get("type",""),        align="center", alt=alt)
        body_cell(ws_log.cell(row=row, column=4), r.get("description",""), align="left",   alt=alt)
        amt = r.get("amount", "")
        try:
            body_cell(ws_log.cell(row=row, column=5), float(amt), align="right", alt=alt, fmt='#,##0.000 "TND"')
        except (ValueError, TypeError):
            body_cell(ws_log.cell(row=row, column=5), "", align="right", alt=alt)

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

    print(f"Exported '{EXPORT_FILE}' — Summary | Income ({n_inc} entries) | Expenses ({n_exp} entries) | Log ({len(log_rows)} entries).")


main()
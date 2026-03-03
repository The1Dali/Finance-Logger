import sys
import csv
import os
import subprocess
import shutil
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

FINANCE_FILE    = os.environ.get("FINANCE_FILE",        "finance.csv")
CONFIG_FILE     = os.environ.get("CONFIG_FILE",         "config.csv")
LOG_FILE        = os.environ.get("LOG_FILE",            "log.csv")
EXPORT_FILE     = os.environ.get("EXPORT_FILE",         "finance.xlsx")
GSHEET_ID       = os.environ.get("GSHEET_ID",           "")
GSHEET_CREDS    = os.environ.get("GSHEET_CREDS",        "credentials.json")
GSHEET_TAB      = os.environ.get("GSHEET_TAB",          "Sheet1")

FIELDS = ["type", "description", "amount", "date", "notes"]

EXIT_OK            = 0
EXIT_BAD_ARGS      = 1
EXIT_INVALID_INPUT = 2
EXIT_NOT_FOUND     = 3
EXIT_FILE_ERROR    = 4


def print_help():
    print(
        "\nUsage: python finance.py <operation>\n"
        "\nOperations:\n"
        "  income <description> <amount> [notes]     Add an income entry.\n"
        "  expense <description> <amount> [notes]    Add an expense entry.\n"
        "  remove income <id>                        Remove an income entry by ID.\n"
        "  remove expense <id>                       Remove an expense entry by ID.\n"
        "  balance <amount>                          Set the opening balance.\n"
        "  list                                      Print a summary to the terminal.\n"
        "  export                                    Export to a formatted .xlsx file.\n"
        "  import [spreadsheet_id]                   Pull new entries from Google Sheets.\n"
        "  clear                                     Wipe all entries.\n"
        "  --help, -h                                Show this help message.\n"
        "\nGoogle Sheets Import:\n"
        "  Requires a Google Cloud service account credentials JSON file.\n"
        "  Your sheet must have a header row with columns: type, description, amount, date, notes\n"
        "  New rows are merged in -- existing entries are never duplicated.\n"
        "\nEnvironment Variables:\n"
        f"  FINANCE_FILE     Path to the finance CSV      (default: {FINANCE_FILE})\n"
        f"  CONFIG_FILE      Path to the config CSV       (default: {CONFIG_FILE})\n"
        f"  LOG_FILE         Path to the log CSV          (default: {LOG_FILE})\n"
        f"  EXPORT_FILE      Path for the .xlsx           (default: {EXPORT_FILE})\n"
        f"  GSHEET_ID        Google Spreadsheet ID        (default: {GSHEET_ID or 'not set'})\n"
        f"  GSHEET_CREDS     Path to service account JSON (default: {GSHEET_CREDS})\n"
        f"  GSHEET_TAB       Sheet tab name to read from  (default: {GSHEET_TAB})\n"
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
                notes = sys.argv[4] if len(sys.argv) > 4 else ""
                add_entry("INCOME", sys.argv[2], sys.argv[3], rows, notes=notes)

            case "income":
                print("Usage: python finance.py income <description> <amount> [notes]", file=sys.stderr)
                sys.exit(EXIT_BAD_ARGS)

            case "expense" if len(sys.argv) > 3:
                notes = sys.argv[4] if len(sys.argv) > 4 else ""
                add_entry("EXPENSE", sys.argv[2], sys.argv[3], rows, notes=notes)

            case "expense":
                print("Usage: python finance.py expense <description> <amount> [notes]", file=sys.stderr)
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

            case "import":
                sheet_id = sys.argv[2] if len(sys.argv) > 2 else GSHEET_ID
                if not sheet_id:
                    print("Error: provide a spreadsheet ID or set the GSHEET_ID environment variable.", file=sys.stderr)
                    sys.exit(EXIT_BAD_ARGS)
                import_from_gsheet(sheet_id, rows)

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
            rows = list(reader)
            for r in rows:
                if "notes" not in r:
                    r["notes"] = ""
            return rows
    except FileNotFoundError:
        try:
            with open(FINANCE_FILE, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=FIELDS)
                writer.writeheader()
        except OSError as e:
            print(f"Error: could not create '{FINANCE_FILE}': {e}", file=sys.stderr)
            sys.exit(EXIT_FILE_ERROR)
        return []


def save_finance(rows):
    try:
        with open(FINANCE_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDS)
            writer.writeheader()
            for r in rows:
                writer.writerow({k: r.get(k, "") for k in FIELDS})
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
    date_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    try:
        with open(LOG_FILE, "a", newline="") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(["date", "action", "type", "description", "amount"])
            writer.writerow([date_str, action, entry_type, description, amount])
    except OSError as e:
        print(f"Warning: could not write to log file '{LOG_FILE}': {e}", file=sys.stderr)


def add_entry(entry_type, description, amount_str, rows, notes="", date_str=None):
    try:
        amount = float(amount_str)
    except ValueError:
        print(f"Error: '{amount_str}' is not a valid number.", file=sys.stderr)
        sys.exit(EXIT_INVALID_INPUT)

    if date_str is None:
        date_str = datetime.now().strftime("%d/%m/%Y")

    rows.append({
        "type":        entry_type,
        "description": description,
        "amount":      str(amount),
        "date":        date_str,
        "notes":       notes,
    })
    save_finance(rows)
    log_event("ADD", entry_type, description, amount)
    note_suffix = f"  [{notes}]" if notes else ""
    print(f"Added {entry_type.lower()} '{description}': {amount:.3f} TND{note_suffix}")


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
    opening  = load_balance()
    income   = [r for r in rows if r["type"].upper() == "INCOME"]
    expenses = [r for r in rows if r["type"].upper() == "EXPENSE"]
    total_in  = sum(float(r["amount"]) for r in income)
    total_exp = sum(float(r["amount"]) for r in expenses)
    profit    = total_in - total_exp
    net       = opening + profit

    print(f"\n{'SUMMARY':=<55}")
    print(f"  {'Opening Balance':<30} {opening:>10.3f} TND")
    print(f"  {'Total Income':<30} {total_in:>10.3f} TND")
    print(f"  {'Total Expenses':<30} {total_exp:>10.3f} TND")
    print(f"  {'Profit':<30} {profit:>10.3f} TND")
    print(f"  {'Net Balance':<30} {net:>10.3f} TND")

    if income:
        print(f"\n{'INCOME':=<55}")
        print(f"  {'#':<5} {'Description':<25} {'Amount':>10}  {'Date':<12}  Notes")
        print(f"  {'-'*70}")
        for i, r in enumerate(income, 1):
            print(f"  {i:<5} {r['description']:<25} {float(r['amount']):>10.3f}  {r['date']:<12}  {r.get('notes','')}")

    if expenses:
        print(f"\n{'EXPENSES':=<55}")
        print(f"  {'#':<5} {'Description':<25} {'Amount':>10}  {'Date':<12}  Notes")
        print(f"  {'-'*70}")
        for i, r in enumerate(expenses, 1):
            print(f"  {i:<5} {r['description']:<25} {float(r['amount']):>10.3f}  {r['date']:<12}  {r.get('notes','')}")
    print()


def clear_all():
    try:
        with open(FINANCE_FILE, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDS)
            writer.writeheader()
    except OSError as e:
        print(f"Error: could not clear '{FINANCE_FILE}': {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)
    log_event("CLEAR")
    print("All entries cleared.")


def import_from_gsheet(sheet_id, existing_rows):
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError:
        print(
            "Error: Google API libraries are required for import.\n"
            "Run: pip install google-api-python-client google-auth",
            file=sys.stderr
        )
        sys.exit(EXIT_FILE_ERROR)

    if not os.path.exists(GSHEET_CREDS):
        print(
            f"Error: credentials file '{GSHEET_CREDS}' not found.\n\n"
            "To set up Google Sheets access:\n"
            "  1. Go to https://console.cloud.google.com\n"
            "  2. Create a project, then go to APIs & Services > Library\n"
            "  3. Enable the 'Google Sheets API'\n"
            "  4. Go to IAM & Admin > Service Accounts and create a service account\n"
            "  5. Click the account > Keys > Add Key > Create new key (JSON)\n"
            "  6. Save the downloaded file as 'credentials.json' next to finance.py\n"
            "     (or set GSHEET_CREDS to its path)\n"
            "  7. Open your Google Sheet and share it with the service account email\n"
            "     (looks like: name@project.iam.gserviceaccount.com) as Viewer\n\n"
            "Your sheet must have a header row with these column names:\n"
            "  type | description | amount | date | notes",
            file=sys.stderr
        )
        sys.exit(EXIT_FILE_ERROR)

    print(f"Connecting to Google Sheets (spreadsheet: {sheet_id})...")

    try:
        creds = service_account.Credentials.from_service_account_file(
            GSHEET_CREDS,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        service = build("sheets", "v4", credentials=creds)
        
        spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        if not sheets:
            print("No sheets found in this spreadsheet.")
            return
            
        print(f"Found {len(sheets)} tab(s): {[s['properties']['title'] for s in sheets]}")
        
    except Exception as e:
        print(f"Error: could not connect to Google Sheets: {e}", file=sys.stderr)
        sys.exit(EXIT_FILE_ERROR)

    total_added = 0
    total_skipped = 0
    total_invalid = 0
    
    for sheet in sheets:
        tab_name = sheet['properties']['title']
        if tab_name.lower() in ['template', 'readme', 'instructions']:
            print(f"\nSkipping tab: {tab_name}")
            continue
            
        print(f"\nProcessing tab: {tab_name}")
        
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=f"'{tab_name}'!A:E",  
            ).execute()
            raw_rows = result.get("values", [])
        except Exception as e:
            print(f"  Warning: could not read tab '{tab_name}': {e}")
            continue

        if not raw_rows:
            print(f"  Tab is empty -- nothing to import.")
            continue

        header = [h.strip().lower() for h in raw_rows[0]]
        missing = {"type", "description", "amount", "date"} - set(header)
        if missing:
            print(f"  Skipping -- missing required columns: {', '.join(sorted(missing))}")
            continue

        def col(row_vals, name):
            try:
                idx = header.index(name)
                return row_vals[idx].strip() if idx < len(row_vals) else ""
            except ValueError:
                return ""

        existing_keys = {
            (r["type"].upper(), r["description"].strip(), r["amount"].strip(), r["date"].strip())
            for r in existing_rows
        }

        added = skipped = invalid = 0

        for raw in raw_rows[1:]:
            if not any(c.strip() for c in raw):
                continue

            entry_type = col(raw, "type").upper()
            description = col(raw, "description")
            amount_str = col(raw, "amount")
            date_str = col(raw, "date")
            notes = col(raw, "notes") if "notes" in header else ""

            if entry_type not in ("INCOME", "EXPENSE"):
                print(f"    Skipped -- unknown type '{entry_type}': {description}")
                invalid += 1
                continue

            try:
                amount = float(amount_str)
            except ValueError:
                print(f"    Skipped -- invalid amount '{amount_str}': {description}")
                invalid += 1
                continue

            key = (entry_type, description, str(amount), date_str)
            if key in existing_keys:
                skipped += 1
                continue

            existing_rows.append({
                "type": entry_type,
                "description": description,
                "amount": str(amount),
                "date": date_str,
                "notes": notes,
            })
            existing_keys.add(key)
            log_event("IMPORT", entry_type, description, amount)
            added += 1
            print(f"    + {entry_type:<8} {description:<30} {amount:.3f} TND")

        print(f"  Tab complete: {added} added, {skipped} already existed, {invalid} invalid rows skipped.")
        total_added += added
        total_skipped += skipped
        total_invalid += invalid

    if total_added > 0:
        save_finance(existing_rows)

    print(f"\n{'='*50}")
    print(f"Import complete across all tabs:")
    print(f"  {total_added} added | {total_skipped} already existed | {total_invalid} invalid")


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

    def solid(color):
        return PatternFill("solid", fgColor=color)

    def hborder(color=BORDER_CLR, bottom_color=None, bottom_weight="hair"):
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
        cell.border    = hborder(bg, bottom_color=bg, bottom_weight="thin")

    def body_cell(cell, value, align="left", alt=False, fmt=None):
        cell.value     = value
        cell.font      = Font(name=FONT, size=10, color="1A1A2E")
        cell.fill      = solid(BODY_ALT if alt else WHITE)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = hborder()
        if fmt:
            cell.number_format = fmt

    income_rows  = [r for r in rows if r["type"].upper() == "INCOME"]
    expense_rows = [r for r in rows if r["type"].upper() == "EXPENSE"]
    opening      = load_balance()

    n_inc = len(income_rows)
    n_exp = len(expense_rows)

    inc_data_start = 4
    inc_data_end   = inc_data_start + n_inc - 1 if n_inc else inc_data_start
    exp_data_start = 4
    exp_data_end   = exp_data_start + n_exp - 1 if n_exp else exp_data_start

    inc_sum_ref = f"'Income'!C{inc_data_start}:C{inc_data_end}" if n_inc else "'Income'!C4:C4"
    exp_sum_ref = f"'Expenses'!C{exp_data_start}:C{exp_data_end}" if n_exp else "'Expenses'!C4:C4"

    ss = wb.active
    ss.title = "Summary"
    ss.sheet_properties.tabColor = BLUE_DARK
    ss.sheet_view.showGridLines  = False
    ss.column_dimensions["A"].width = 36
    ss.column_dimensions["B"].width = 24
    ss.column_dimensions["C"].width = 22

    def section_header(row_num, text, bg, fg=WHITE):
        ss.merge_cells(f"A{row_num}:C{row_num}")
        c = ss.cell(row=row_num, column=1, value=text)
        c.font      = Font(name=FONT, bold=True, size=9, color=fg)
        c.fill      = solid(bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border    = Border(left=Side(style="medium", color=bg), bottom=Side(style="thin", color=bg))
        ss.row_dimensions[row_num].height = 16

    def summary_row(row_num, label, value_formula, context, bg, fg, border_col, bold_val=False, pct_fmt=False):
        ss.row_dimensions[row_num].height = 24

        lc = ss.cell(row=row_num, column=1, value=label)
        lc.font      = Font(name=FONT, bold=True, size=10, color=fg)
        lc.fill      = solid(bg)
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        lc.border    = Border(
            left=Side(style="medium", color=border_col),
            top=Side(style="hair", color=BORDER_CLR),
            bottom=Side(style="hair", color=BORDER_CLR),
            right=Side(style="hair", color=BORDER_CLR),
        )

        vc = ss.cell(row=row_num, column=2)
        if value_formula is None:
            vc.value = ""
        elif isinstance(value_formula, str) and value_formula.startswith("="):
            vc.value = value_formula
        else:
            vc.value = float(value_formula)
        vc.font          = Font(name=FONT, bold=bold_val, size=10, color=fg)
        vc.fill          = solid(bg)
        vc.alignment     = Alignment(horizontal="right", vertical="center")
        vc.number_format = "0.0%" if pct_fmt else '#,##0.000 "TND"'
        vc.border        = Border(
            top=Side(style="hair", color=BORDER_CLR),
            bottom=Side(style="hair", color=BORDER_CLR),
            left=Side(style="hair", color=BORDER_CLR),
            right=Side(style="hair", color=BORDER_CLR),
        )

        cc = ss.cell(row=row_num, column=3)
        cc.value     = context
        cc.font      = Font(name=FONT, size=9, color=fg)
        cc.fill      = solid(bg)
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc.border    = Border(
            right=Side(style="medium", color=border_col),
            top=Side(style="hair", color=BORDER_CLR),
            bottom=Side(style="hair", color=BORDER_CLR),
            left=Side(style="hair", color=BORDER_CLR),
        )

    def spacer_row(row_num):
        ss.row_dimensions[row_num].height = 8
        for col in range(1, 4):
            ss.cell(row=row_num, column=col).fill = solid("F0F4FB")

    export_date = datetime.now().strftime("%d/%m/%Y %H:%M")
    inc_label   = f"{n_inc} entr{'y' if n_inc == 1 else 'ies'}"
    exp_label   = f"{n_exp} entr{'y' if n_exp == 1 else 'ies'}"

    ss.row_dimensions[1].height = 44
    ss.merge_cells("A1:C1")
    title_cell(ss["A1"], "FINANCIAL SUMMARY", BLUE_DARK)

    spacer_row(2)

    section_header(3, "  BALANCE OVERVIEW", BLUE_MID)
    summary_row(4,  "Opening Balance",  str(opening),            "Starting funds",                                  BLUE_LIGHT,  BLUE_FG,  BLUE_DARK)
    summary_row(5,  "Total Income",     f"=SUM({inc_sum_ref})",  inc_label,                                         GREEN_LIGHT, GREEN_FG, GREEN_DARK)
    summary_row(6,  "Total Expenses",   f"=SUM({exp_sum_ref})",  exp_label,                                         RED_LIGHT,   RED_FG,   RED_DARK)

    spacer_row(7)

    section_header(8, "  RESULTS", BLUE_MID)
    summary_row(9,  "Gross Profit",   "=B5-B6",            '=IF(B9>=0,"Surplus \u25b2","Deficit \u25bc")',          BLUE_LIGHT, BLUE_FG, BLUE_DARK)
    summary_row(10, "Expense Ratio",  "=IFERROR(B6/B5,0)", '=IFERROR(TEXT(B6/B5,"0.0%")&" of income","N/A")',       BLUE_LIGHT, BLUE_FG, BLUE_DARK, pct_fmt=True)
    summary_row(11, "Total Profit",   "=B5-B6",            '=IF(B11>=0,"Surplus \u25b2","Deficit \u25bc")',          BLUE_LIGHT, BLUE_FG, BLUE_DARK)
    summary_row(12, "Net Balance",    "=B4+B11",           '=IF(B12>=0,"Positive \u2714","Negative \u2716")',        BLUE_MID,   WHITE,   BLUE_DARK, bold_val=True)

    spacer_row(13)

    section_header(14, "  METADATA", "7A90B8")
    summary_row(15, "Total Entries",  None, f"{n_inc + n_exp} total ({n_inc} income, {n_exp} expense)", "EEF2FA", "3A4A6B", "7A90B8")
    ss.cell(row=15, column=2).value         = n_inc + n_exp
    ss.cell(row=15, column=2).number_format = "General"
    ss.cell(row=15, column=2).font          = Font(name=FONT, size=10, color="3A4A6B")
    summary_row(16, "Last Exported",  None, export_date,                                                "EEF2FA", "3A4A6B", "7A90B8")
    ss.cell(row=16, column=2).value         = ""
    ss.cell(row=16, column=2).number_format = "General"

    ws_inc = wb.create_sheet("Income")
    ws_inc.sheet_properties.tabColor = GREEN_DARK
    ws_inc.sheet_view.showGridLines  = False
    ws_inc.column_dimensions["A"].width = 6
    ws_inc.column_dimensions["B"].width = 34
    ws_inc.column_dimensions["C"].width = 18
    ws_inc.column_dimensions["D"].width = 14
    ws_inc.column_dimensions["E"].width = 30

    ws_inc.row_dimensions[1].height = 38
    ws_inc.row_dimensions[2].height = 6
    ws_inc.row_dimensions[3].height = 20
    ws_inc.merge_cells("A1:E1")
    title_cell(ws_inc["A1"], "INCOME", GREEN_DARK)
    ws_inc.freeze_panes = "A4"

    for col, (text, align) in enumerate([
        ("ID", "center"), ("Description", "left"), ("Amount (TND)", "right"),
        ("Date", "center"), ("Notes", "left"),
    ], 1):
        header_cell(ws_inc.cell(row=3, column=col), text, GREEN_MID, align=align)

    for i, r in enumerate(income_rows):
        row = i + inc_data_start
        alt = i % 2 == 1
        ws_inc.row_dimensions[row].height = 20
        body_cell(ws_inc.cell(row=row, column=1), f"=ROW()-{inc_data_start - 1}", align="center", alt=alt)
        body_cell(ws_inc.cell(row=row, column=2), r["description"],               align="left",   alt=alt)
        body_cell(ws_inc.cell(row=row, column=3), float(r["amount"]),             align="right",  alt=alt, fmt='#,##0.000 "TND"')
        body_cell(ws_inc.cell(row=row, column=4), r["date"],                      align="center", alt=alt)
        body_cell(ws_inc.cell(row=row, column=5), r.get("notes", ""),             align="left",   alt=alt)

    ws_exp = wb.create_sheet("Expenses")
    ws_exp.sheet_properties.tabColor = RED_DARK
    ws_exp.sheet_view.showGridLines  = False
    ws_exp.column_dimensions["A"].width = 6
    ws_exp.column_dimensions["B"].width = 34
    ws_exp.column_dimensions["C"].width = 18
    ws_exp.column_dimensions["D"].width = 14
    ws_exp.column_dimensions["E"].width = 30

    ws_exp.row_dimensions[1].height = 38
    ws_exp.row_dimensions[2].height = 6
    ws_exp.row_dimensions[3].height = 20
    ws_exp.merge_cells("A1:E1")
    title_cell(ws_exp["A1"], "EXPENSES", RED_DARK)
    ws_exp.freeze_panes = "A4"

    for col, (text, align) in enumerate([
        ("ID", "center"), ("Description", "left"), ("Amount (TND)", "right"),
        ("Date", "center"), ("Notes", "left"),
    ], 1):
        header_cell(ws_exp.cell(row=3, column=col), text, RED_MID, align=align)

    for i, r in enumerate(expense_rows):
        row = i + exp_data_start
        alt = i % 2 == 1
        ws_exp.row_dimensions[row].height = 20
        body_cell(ws_exp.cell(row=row, column=1), f"=ROW()-{exp_data_start - 1}", align="center", alt=alt)
        body_cell(ws_exp.cell(row=row, column=2), r["description"],               align="left",   alt=alt)
        body_cell(ws_exp.cell(row=row, column=3), float(r["amount"]),             align="right",  alt=alt, fmt='#,##0.000 "TND"')
        body_cell(ws_exp.cell(row=row, column=4), r["date"],                      align="center", alt=alt)
        body_cell(ws_exp.cell(row=row, column=5), r.get("notes", ""),             align="left",   alt=alt)

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

    for col, (text, align) in enumerate([
        ("Date", "center"), ("Action", "center"), ("Type", "center"),
        ("Description", "left"), ("Amount", "right"),
    ], 1):
        header_cell(ws_log.cell(row=3, column=col), text, "4A4A4A", align=align)

    log_rows = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", newline="") as f:
                reader = csv.DictReader(f)
                log_rows = list(reader)
        except OSError:
            pass

    ACTION_BG = {"ADD": "D4EDDA", "REMOVE": "F8D7DA", "CLEAR": "FFF3CD", "BALANCE": "D6E4F7", "IMPORT": "E8D5F5"}
    ACTION_FG = {"ADD": "155724", "REMOVE": "721C24", "CLEAR": "856404", "BALANCE": "0D2B5E", "IMPORT": "4A0E72"}

    for i, r in enumerate(log_rows):
        row = i + 4
        alt = i % 2 == 1
        ws_log.row_dimensions[row].height = 20
        action = r.get("action", "").upper()
        abg = ACTION_BG.get(action, WHITE)
        afg = ACTION_FG.get(action, "333333")

        body_cell(ws_log.cell(row=row, column=1), r.get("date", ""),        align="center", alt=alt)
        ac = ws_log.cell(row=row, column=2, value=action)
        ac.font      = Font(name=FONT, bold=True, size=9, color=afg)
        ac.fill      = solid(abg)
        ac.alignment = Alignment(horizontal="center", vertical="center")
        ac.border    = hborder()
        body_cell(ws_log.cell(row=row, column=3), r.get("type", ""),        align="center", alt=alt)
        body_cell(ws_log.cell(row=row, column=4), r.get("description", ""), align="left",   alt=alt)
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
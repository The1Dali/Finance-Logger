import sys
import csv
import os
from datetime import datetime

FINANCE_FILE = os.environ.get("FINANCE_FILE", "finance.csv")
LOG_FILE = os.environ.get("LOG_FILE", "log.csv")

EXIT_OK           = 0
EXIT_BAD_ARGS     = 1
EXIT_INVALID_INPUT = 2
EXIT_NOT_FOUND    = 3
EXIT_FILE_ERROR   = 4


def print_help():
    print(
        "\nUsage: python finance.py <operation>\n"
        "\nOperations:\n"
        "  +  <name> <value>   Add <value> to an item called <name>. Creates it if it doesn't exist.\n"
        "  -  <name>           Remove the item called <name>.\n"
        "  clear               Wipe all entries from the finance file.\n"
        "  list                Print all current entries to the terminal.\n"
        "  --help, -h          Show this help message.\n"
        "\nEnvironment Variables:\n"
        f"  FINANCE_FILE        Path to the finance CSV file (default: {FINANCE_FILE})\n"
        f"  LOG_FILE            Path to the log CSV file     (default: {LOG_FILE})\n"
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

    date_str   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  
    action_fmt = str(action).ljust(8)
    name_fmt   = str(name).ljust(45)
    value_fmt  = (f"{float(value):.3f}" if value != "" else "").rjust(20)

    header = [
        "Date               ",   
        "Action  ",              
        "Item".ljust(45),      
        "Value".rjust(20),       
    ]

    row = [date_str, action_fmt, name_fmt, value_fmt]

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


main()
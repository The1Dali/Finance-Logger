import sys
import csv
import os
from datetime import datetime


def main():
    # Fix #9: close the file immediately after reading rows
    try:
        with open("finance.csv", "r") as file:
            reader = csv.reader(file, delimiter="-")
            rows = [[cell.strip() for cell in row] for row in reader]
    except FileNotFoundError:
        # Fix #1: create the file if it doesn't exist yet
        rows = []
        with open("finance.csv", "w"):
            pass

    try:
        match sys.argv[1]:
            case "*.":
                flash()
            case "+" if len(sys.argv) > 3:
                add(sys.argv[2], sys.argv[3], rows)
            case "+":
                # Fix #8: targeted error message for missing arguments
                print("Usage: python finance.py + <name> <value>")
                sys.exit(1)
            case "-" if len(sys.argv) > 2:
                remove(sys.argv[2], rows)
            case "-":
                # Fix #8: targeted error message for missing arguments
                print("Usage: python finance.py - <name>")
                sys.exit(1)
            case _:
                print("Correct usage: python finance.py {operation}")
                sys.exit(1)
    except IndexError:
        print("Operation not specified")
        sys.exit(1)

    sys.exit(0)


def log(action, name="", value=""):
    # Fix #10: use os.path.exists instead of opening the file to check
    write_header = not os.path.exists("log.csv")

    with open("log.csv", "a", newline="") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["Date", "Action", "Item", "Value"])
        writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), action, name, value])


def write(rows):
    with open("finance.csv", "w", newline="") as file:
        writer = csv.writer(file, delimiter="-")
        for row in rows:
            # Fix #7: safely fall back to 0.000% if percentage is missing or empty
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


def refresh(rows):
    # Fix #4: renamed 'sum' to 'total' to avoid shadowing the built-in
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

    # Fix #3: sort in place so the list is actually mutated, not just a local rebind
    rows.sort(key=lambda row: float(row[2]), reverse=True)
    write(rows)


def add(name, value, rows):
    # Fix #6: validate that value is numeric at the entry point with a clear error message
    try:
        value = float(value)
    except ValueError:
        print(f"Error: '{value}' is not a valid number.")
        sys.exit(1)

    found = False
    for i in range(len(rows)):
        if rows[i][0].strip().upper() == name.upper():
            rows[i][1] = str(float(rows[i][1]) + value)
            found = True
            # Fix #2: break after first match, consistent with remove()
            break

    if not found:
        # Fix #7: initialize percentage to "0.0" instead of "" to avoid formatting crash
        rows.append([name, str(value), "0.0"])

    log("ADD", name, value)
    refresh(rows)


def remove(name, rows):
    for i in range(len(rows)):
        if rows[i][0].strip().upper() == name.upper():
            rows.pop(i)
            print("Element removed")
            log("REMOVE", name)
            break
    else:
        print("Element not found")

    refresh(rows)


def flash():
    # Fix #5: log the flash event without ever deleting anything from log.csv
    with open("finance.csv", "w"):
        pass
    log("FLASH", "", "")
    print("Flashed")


main()
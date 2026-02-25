import sys
import csv
from datetime import datetime

def main():
    with open("finance.csv", "r") as file:
        reader = csv.reader(file, delimiter = "-")
        rows = [[cell.strip() for cell in row] for row in reader]
        try:
            match sys.argv[1]:
                case "*.":
                    flash()
                case "+" if len(sys.argv) > 3:  
                    add(sys.argv[2], sys.argv[3], rows)
                case "-" if len(sys.argv) > 2:
                    remove(sys.argv[2], rows)
                case _:
                    print("Correct usage: python finance.py {operation}")
                    sys.exit(1)
        except IndexError:
            print("Operation not specified")
            sys.exit(1)
    sys.exit(0)

def log(action, name, value=""):
    file_exists = False
    try:
        with open("log.csv", "r") as f:
            file_exists = any(True for _ in f)
    except FileNotFoundError:
        pass

    with open("log.csv", "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["Date", "Action", "Item", "Value"])
        writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), action, name, value])

def write(rows):
    with open("finance.csv", "w") as file:
        writer = csv.writer(file, delimiter = "-")
        for row in rows:
            f_row = [row[0].ljust(45), (f"{float(row[1]):.3f}").center(20), (f"{float(row[2]):.3f}%").rjust(20)]
            writer.writerow(f_row)

def refresh(rows):
    sum = 0
    total_index = -1
    for i in range(len(rows)):
        if rows[i][0].upper() == "TOTAL":
            total_index = i
        else:
            sum += float(rows[i][1])
    if total_index >= 0:
        rows[total_index][1] = str(sum)
    elif len(rows) > 0:
        rows.append(["TOTAL", str(sum), ""])
    for i in range(len(rows)):
        try:
            rows[i][2] = str(float(rows[i][1]) / sum * 100)
        except ZeroDivisionError:
            continue
    rows = sorted(rows, key = lambda row: float(row[2]), reverse = True)
    write(rows)
    
def add(name, value, rows):
    found = False
    for i in range(len(rows)):
        if rows[i][0].upper() == name.upper():
            rows[i][1] = str(float(rows[i][1]) + float(value))
            found = True
    if not(found):
        rows.append([name, value, ""])
    log("ADD", name, value)
    refresh(rows)

def remove(name, rows):
    for i in range(len(rows)):
        if rows[i][0].upper() == name.upper():
            rows.pop(i)
            print("Element removed")
            log("REMOVE", name)
            break
    else:
        print("Element not found")
    refresh(rows)

def flash():
    with open("finance.csv", "w") as file:
        print("Flashed")

main()
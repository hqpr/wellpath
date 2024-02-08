"""Put .xlsx file in the root directory and convert .xlsx file to csv"""
import json
import csv
from datetime import datetime

file_path = "TimeDetails_SLCJ.csv"

result = []
collecting = False
collect_transactions = False
record = {}
employee_info = {}


def convert_date(date_string):
    if date_string:
        try:
            date_object = datetime.strptime(date_string, "%b %d, %Y %I:%M %p")
            return date_object.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            try:
                date_object = datetime.strptime(date_string, "%b %d, %Y")
                return date_object.strftime("%Y-%m-%d")
            except ValueError:
                return date_string
    return date_string


def safe_get(lst, idx, default=None):
    try:
        return lst[idx]
    except IndexError:
        return default


def get_first_last_name(full_name):
    first = full_name.split(",")[1]
    last = full_name.split(",")[0]
    return first.strip(), last.strip()


with open(file_path, mode="r", encoding="utf-8") as file:
    csv_reader = csv.reader(file, delimiter=";")

    for _ in range(5):
        next(csv_reader)

    for row in csv_reader:
        if row[0].startswith("Employee Name"):
            if collecting:
                result.append(record)
                record = {}
            collecting = True
            collect_transactions = False
            name = row[24].split(" (")[0]
            first_name, last_name = get_first_last_name(name)
            record["Employee First Name"] = first_name
            record["Employee Last Name"] = last_name
            record["Employee Reference"] = (
                row[24].split("(")[-1].replace(")", "").strip()
            )
        elif "Job" in row[0]:
            record["Job"] = row[7]
        elif "Transaction Apply Date" in row[0]:
            collect_transactions = True
        elif collecting and collect_transactions:
            record["Employee Payrule"] = row[37]
    if collecting:
        result.append(record)

# Assuming it's a last record with no records
result = [record for record in result if record.get("Employee Last Name") != "Wergen"]

json_result = json.dumps(result, ensure_ascii=False, indent=4)

with open("well_short.json", "w") as f:
    f.write(json_result)

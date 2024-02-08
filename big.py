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
            if employee_info:
                for transaction in transactions:
                    record = employee_info.copy()
                    record.update(transaction)
                    result.append(record)
                employee_info = {}

            name = row[24].split(" (")[0]
            first_name, last_name = get_first_last_name(name)
            employee_info["Employee First Name"] = first_name
            employee_info["Employee Last Name"] = last_name
            employee_info["Employee Reference"] = (
                row[24].split("(")[-1].replace(")", "").strip()
            )
            transactions = []
        elif "Job" in row[0]:
            employee_info["Job"] = row[7]
            location = row[28].split("/")
            employee_info["Organization"] = safe_get(location, 0)
            employee_info["Area"] = safe_get(location, 1)
            employee_info["Region"] = safe_get(location, 2)
            employee_info["Site"] = safe_get(location, 3)
            employee_info["Department"] = safe_get(location, 4)
        elif "Transaction Apply Date" in row[0]:
            collect_transactions = True
        elif collect_transactions:
            transaction_info = {
                "Transaction Apply Date": convert_date(row[0]),
                "Transaction Apply To": row[6],
                "Transaction Type": row[11],
                "Transaction In Exceptions": row[15],
                "Transaction Out Exceptions": row[18],
                "Hours": row[22],
                "Money": row[25],
                "Days": row[29],
                "Transaction Start Date/Time": convert_date(row[32]),
                "Transaction End Date/Time": convert_date(row[34]),
                "Employee Payrule": row[37],
            }
            transactions.append(transaction_info)

    if employee_info and transactions:
        for transaction in transactions:
            record = employee_info.copy()
            record.update(transaction)
            result.append(record)

# Assuming it's a last record with no records
result = [record for record in result if record.get("Employee Last Name") != "Wergen"]

json_result = json.dumps(result, ensure_ascii=False, indent=4)

with open("well.json", "w") as f:
    f.write(json_result)

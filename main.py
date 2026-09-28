from pathlib import Path
import csv
from openpyxl import load_workbook
from openpyxl import Workbook
from dataclasses import dataclass
import re

BASE_FOLDER = Path(__file__).resolve().parent
INPUT_FOLDER = BASE_FOLDER / "input"
OUTPUT_FOLDER = BASE_FOLDER / "output"
REPORTS_FOLDER = BASE_FOLDER / "reports"

REQUIRED_FIELDS = ("name", "email", "amount")
EMAIL_PATTERN = re.compile(
    r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
)

@dataclass
class RecordResult:
    record: dict
    status: str
    source_file: str | None = None
    reason: str | None = None

def find_input_files():
    if not INPUT_FOLDER.exists():
        return []
        
    return sorted(
        file_path
        for file_path in INPUT_FOLDER.iterdir()
        if file_path.is_file()
        and file_path.suffix.lower() in {".csv", ".xlsx"}
    )

def read_csv_file(file_path):
    records = []

    with open(file_path, "r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        
        if reader.fieldnames is None:
            return []
        
        for row in reader:
            records.append(
                RecordResult(
                    record=row,
                    status="unprocessed",
                    source_file=file_path.name
                )
            )

    return records
   
def read_xlsx_file(file_path):
    workbook = load_workbook(file_path, data_only=True)
    worksheet = workbook.active

    rows = list(worksheet.iter_rows(values_only=True))

    if not rows:
        workbook.close()
        return []

    headers = rows[0]
    records = []

    for row in rows[1:]:
        if all(value is None for value in row):
            continue
            
        record = dict(zip(headers, row))

        records.append(
            RecordResult(
                record=record,
                status="unprocessed",
                source_file=file_path.name
            )
        )

    workbook.close()

    return records

def read_files():
    files = find_input_files()
    all_records = []
       
    for file_path in files:
        extension = file_path.suffix.lower()
        
        if extension == ".csv":
            records = read_csv_file(file_path)
        elif extension == ".xlsx":
            records = read_xlsx_file(file_path)
        else:
            continue
            
        all_records.extend(records)
    
    return all_records
    
def normalize_columns(records):
    for result in records:
        normalized_record = {}

        for header, value in result.record.items():
            if not isinstance(header, str):
                result.status = "failed"
                result.reason = "Invalid header"
                break

            normalized_header = header.strip().lower().replace(" ", "_")
            
            if not normalized_header:
                result.status = "failed"
                result.reason = "Empty header"
                break
                
            normalized_record[normalized_header] = value

        if result.status != "failed":
            result.record = normalized_record
            result.status = "normalized"

    return records
    
def clean_records(records):
    for result in records:
        if result.status == "failed":
            continue
        
        for field in REQUIRED_FIELDS:
            value = result.record.get(field)

            if value is None or str(value).strip() == "":
                result.status = "failed"
                result.reason = f"Missing required field: {field}"
                break
        
        if result.status == "failed":
            continue
        
        cleaned_record = {}

        for header, value in result.record.items():
            cleaned_value = value

            if isinstance(value, str):
                cleaned_value = value.strip()

                if header == "email":
                    cleaned_value = cleaned_value.lower()

                elif header == "state":
                    cleaned_value = cleaned_value.upper()

            if header == "amount":
                try:
                    cleaned_value = float(cleaned_value)
                except (ValueError, TypeError):
                    result.status = "failed"
                    result.reason = "Invalid amount"
                    break

            cleaned_record[header] = cleaned_value

        if result.status != "failed":
            result.record = cleaned_record
            result.status = "valid"

    return records

def remove_duplicates(records):
    seen_emails = set()
        
    for result in records:
        if result.status == "failed":
            continue
            
        email = result.record.get("email")
        
        if not isinstance(email, str) or not EMAIL_PATTERN.fullmatch(email):
            result.status = "failed"
            result.reason = "Invalid email"
            continue
        
        if email in seen_emails:
            result.status = "duplicate"
            result.reason = "Duplicate email"
            continue
            
        seen_emails.add(email)
        
    return records

def write_cleaned_excel(records):
    valid_records = [
        result for result in records
        if result.status == "valid"
    ]

    if not valid_records:
        print("No valid records to write.")
        return

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Cleaned Data"

    headers = []

    for result in valid_records:
        for header in result.record:
            if header not in headers:
                headers.append(header)

    worksheet.append(headers)

    for result in valid_records:
        row = []

        for header in headers:
            row.append(result.record.get(header))

        worksheet.append(row)

    output_path = OUTPUT_FOLDER / "cleaned_data.xlsx"
    workbook.save(output_path)
    workbook.close()

    print("Cleaned Excel file created successfully.")
    
def write_error_report(records):
    error_records = [
        result for result in records
        if result.status in {"failed", "duplicate"}
    ]
    
    if not error_records:
        print("No errors or duplicates found during operation")
        return
    
    REPORTS_FOLDER.mkdir(parents=True, exist_ok=True)
    
    report_path = REPORTS_FOLDER / "error_report.csv"
    
    with open(report_path, "w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        
        writer.writerow([
            "source_file",
            "status",
            "reason",
            "name",
            "email",
            "amount"
        ])
        
        for result in error_records:
            writer.writerow([
                result.source_file,
                result.status,
                result.reason,
                result.record.get("name"),
                result.record.get("email"),
                result.record.get("amount")
            ])
            
    print("Error report created successfully.")
    
def main():
    records = read_files()
    
    if not records:
        print("No records found.")
        return
        
    records = normalize_columns(records)
    records = clean_records(records)
    records = remove_duplicates(records)
    
    write_cleaned_excel(records)
    write_error_report(records)

if __name__ == "__main__":
    main()
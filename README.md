# Multi-File CSV/XLSX Data Cleaner

## Description

A Python command-line utility that reads customer-style data from multiple CSV and Excel files, normalizes and validates the records, removes duplicate email entries, writes cleaned data to Excel, and produces a separate CSV error report for failed or duplicate records.

## Problem

Business data often arrives from multiple spreadsheets with inconsistent formatting, mixed capitalization, malformed values, duplicate records, and slightly different column naming conventions.

Cleaning these files manually is repetitive and makes it easy to miss bad records.

This program combines supported input files into one processing workflow, normalizes their structure, validates required fields, detects duplicates, and separates clean output from records that need review.

## Features

- Reads multiple `.csv` and `.xlsx` files from one input folder
- Combines records from all supported files into one workflow
- Tracks the source file for every record
- Normalizes column names
- Rejects invalid or empty headers
- Validates required fields
- Trims surrounding whitespace
- Converts email addresses to lowercase
- Converts state values to uppercase
- Converts `amount` values to numbers
- Validates email format
- Detects duplicate email addresses across all input files
- Preserves optional columns found in valid records
- Writes valid records to Excel
- Writes failed and duplicate records to a separate CSV report
- Continues processing when individual records fail validation

## Required Fields

Each record must contain:

- `name`
- `email`
- `amount`

Column names are normalized before validation, so a header such as:

```text
Email
```

or:

```text
 email 
```

will normalize to:

```text
email
```

Spaces inside headers are converted to underscores.

For example:

```text
Customer State
```

becomes:

```text
customer_state
```

## Supported Input Files

The program reads:

- `.csv`
- `.xlsx`

Only files directly inside the `input` folder are processed.

For Excel files, the active worksheet is used.

## Project Structure

```text
multi-file-data-cleaner/
├── input/
│   ├── customers_a.csv
│   └── customers_b.xlsx
├── output/              # created automatically
│   └── cleaned_data.xlsx
├── reports/             # created automatically
│   └── error_report.csv
├── main.py
├── README.md
├── requirements.txt
└── .gitignore
```

The `output` and `reports` folders are generated at runtime and should not be committed.

## Requirements

- Python 3.10 or newer
- `openpyxl`

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Running the Program

1. Clone or download the repository.
2. Place one or more `.csv` or `.xlsx` files in the `input` folder.
3. Open a terminal in the project directory.
4. Install dependencies if needed:

```bash
pip install -r requirements.txt
```

5. Run:

```bash
python main.py
```

If valid records are found, the cleaned workbook is created at:

```text
output/cleaned_data.xlsx
```

If failed or duplicate records are found, the error report is created at:

```text
reports/error_report.csv
```

## Processing Workflow

The program processes records in several stages:

1. Discover supported input files.
2. Read CSV and XLSX records.
3. Store source-file information for each record.
4. Normalize column names.
5. Validate required fields.
6. Clean string values.
7. Convert `amount` values to floating-point numbers.
8. Validate email addresses.
9. Detect duplicate emails across all input files.
10. Write valid records to Excel.
11. Write failed and duplicate records to the error report.

## Column Normalization

Column names are:

- stripped of surrounding whitespace
- converted to lowercase
- changed from spaces to underscores

Examples:

```text
Name        -> name
Email       -> email
Amount      -> amount
State       -> state
Customer ID -> customer_id
```

Records with non-text or empty headers are marked as failed.

## Data Cleaning

For string values, surrounding whitespace is removed.

Additional field-specific cleaning includes:

- `email` values are converted to lowercase
- `state` values are converted to uppercase
- `amount` values are converted to `float`

Records with invalid amount values are marked as failed.

## Email Validation

The program uses a regular expression to perform basic practical email validation.

The pattern checks that an email contains:

- text before `@`
- an `@` symbol
- text after `@`
- a dot and domain portion
- no whitespace in the address

The validation is intentionally lightweight rather than attempting full RFC email validation.

## Duplicate Handling

Email addresses are used as the unique identifier for duplicate detection.

After email normalization and validation:

- the first occurrence of an email remains valid
- later records with the same email are marked as `duplicate`
- duplicates are excluded from the cleaned Excel file
- duplicate records appear in the error report with the reason `Duplicate email`

Duplicate detection works across all input files in the same run.

## Cleaned Excel Output

Only records with status `valid` are written to:

```text
output/cleaned_data.xlsx
```

The worksheet is named:

```text
Cleaned Data
```

The output preserves all columns encountered in valid records, including optional fields beyond the required fields.

This prevents additional client-specific data from being discarded during cleaning.

## Error Report

Failed and duplicate records are written to:

```text
reports/error_report.csv
```

The report contains:

```csv
source_file,status,reason,name,email,amount
```

Possible reasons include:

- Invalid header
- Empty header
- Missing required field
- Invalid amount
- Invalid email
- Duplicate email

## Error Handling

A bad record does not stop the rest of the dataset from being processed.

Each record carries its own processing status and optional failure reason through the workflow using the `RecordResult` dataclass.

Possible statuses include:

- `unprocessed`
- `normalized`
- `valid`
- `failed`
- `duplicate`

## Design Decisions

### Record Tracking with a Dataclass

`RecordResult` keeps the raw or cleaned record together with:

- processing status
- source filename
- failure or duplicate reason

This makes it possible to process valid data while still preserving enough context to explain rejected records.

### Separate Processing Stages

File reading, header normalization, record cleaning, duplicate detection, output generation, and error reporting are handled by separate functions.

This keeps each stage focused on one responsibility and makes the workflow easier to understand and maintain.

### Cross-File Duplicate Detection

All records are combined before duplicate detection.

This allows duplicate email addresses to be identified even when they appear in different input files.

### Preserve Optional Columns

Valid records may contain additional fields beyond `name`, `email`, and `amount`.

The Excel output collects every column found across valid records so those optional fields are preserved instead of discarded.

## Current Limitations

The current version intentionally:

- Processes only `.csv` and `.xlsx` files
- Processes only files directly inside the `input` folder
- Uses the active worksheet from Excel files
- Uses email as the duplicate identifier
- Uses lightweight email validation rather than full RFC validation
- Converts `amount` values to Python floating-point numbers
- Replaces generated output files on each run

These behaviors define the current project scope.

## Testing

The project should be tested with combinations such as:

- valid CSV records
- valid XLSX records
- mixed CSV and XLSX input
- missing required fields
- blank required values
- invalid amount values
- invalid email addresses
- duplicate emails within one file
- duplicate emails across different files
- extra optional columns
- uppercase/lowercase email differences
- state values requiring normalization
- whitespace around headers and values
- empty spreadsheet rows
- malformed or empty headers

## What I Learned

This project provides practical experience with:

- Python file processing
- CSV parsing
- Excel reading and writing with `openpyxl`
- Dataclasses
- Regular expressions
- Data validation
- Data normalization
- Duplicate detection
- Error reporting
- Multi-file processing
- Preserving optional schema fields
- Separating a data-processing pipeline into clear stages

## Author

Created as a practical Python data-cleaning and spreadsheet-automation portfolio project.

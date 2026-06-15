# Test Data Generator

A Python utility that reads a database schema from an Excel workbook and generates realistic test data along with ready-to-use SQL INSERT and DELETE scripts.

## Output

Produces a timestamped Excel file: `HMS_Test_Data_<YYYYMMDD_HHMMSS>.xlsx`

For each table three sheets are created:

| Sheet | Description | Header colour |
|---|---|---|
| `<Table>_testdata` | Generated data rows | Blue |
| `<Table>_insertscript` | SQL INSERT statements | Green |
| `<Table>_deletescript` | SQL DELETE statements | Orange |

## Requirements

- Python 3.8+
- `openpyxl >= 3.1.0`
- `faker >= 26.0.0`

## Environment Setup

It is recommended to use a virtual environment to keep dependencies isolated.

**Create and activate the virtual environment:**
Termainal > New
# Create
py -m venv .venv
# Activate (Windows PowerShell)
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
# Run the main script
py generate_test_data.py
deactivate 
## Usage

1. **Prepare the schema workbook** — create `HMS_Table_Schema.xlsx` in the project folder with one or more sheets. Each sheet must have the following columns (column positions are auto-detected by header name, case-insensitive):

   | Column | Description |
   |---|---|
   | `schema_name` | Database schema name |
   | `table_name` | Table name |
   | `column_name` | Column name |
   | `data_type` | SQL data type (`int`, `varchar`, `date`, `decimal`, `bigint`, …) |
   | `max_length` | Maximum length (numeric) |
   | `precision` | Numeric precision (optional) |
   | `scale` | Numeric scale (optional) |
   | `rule` | Generation rule (optional — see Rule Syntax below) |

2. **Run the script:**

   ```bash
   py generate_test_data.py
   ```

3. Open the generated `HMS_Test_Data_*.xlsx` file.

## Getting the Schema from SQL Server

Use the provided SQL script (`schema.script to fetch.sql`) to query `sys.columns` and export the results into `HMS_Table_Schema.xlsx`. Update the `WHERE` clause to target your schema and tables:

```sql
WHERE s.name = 'your_schema_name'
AND t.name IN ('Table1', 'Table2', ...)
```

## Rule Syntax

Rules are placed in the `Rule` column of the schema workbook and control how values are generated for a column.

| Rule | Example | Behaviour |
|---|---|---|
| *(blank)* | | Smart generation based on data type and column name |
| `oneof:A,B,C` | `oneof:Y,N` | Cartesian product — generates one row per combination of values |
| `NULL` (inside oneof list) | `oneof:Y,NULL` | Inserts SQL `NULL` |
| `startswith:<prefix>` | `startswith:9033` | Value begins with the given prefix |
| `startswith:<prefix>xxx` | `startswith:DBLQAxxxxxxx` | Each `x` is replaced by one random digit |
| `range:min,max` | `range:1,100` | Generates boundary values: minimum, midpoint, and maximum |
| *(plain text)* | `ACTIVE` | Fixed value — always inserts exactly that text |

## Project Structure

```
datagenerator/
├── generate_test_data.py       # Main script
├── requirements.txt            # Python dependencies
├── schema.script to fetch.sql  # SQL Server query to build HMS_Table_Schema.xlsx
└── README.md
```

## Input File Name

The script looks specifically for `HMS_Table_Schema.xlsx` in the same folder as [generate_test_data.py](generate_test_data.py).

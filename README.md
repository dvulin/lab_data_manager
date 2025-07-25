# ExcelParser

A lightweight utility for extracting both single-cell values and multi-row tables from Excel worksheets into a structured Python dictionary, with HTML and JSON export.

---

## Features

* **Load any sheet**
  Load by name or index, with optional header row handling.

* **Flexible range parsing**

  * Excel-style (e.g. `B2,E11` or `A2:B8`)
  * Numeric lists (`"1,2,5"`)
  * Numeric slices (`"4:8"`)

* **Value extraction**
  Pull arbitrary cell ranges (non-table data) into `sheet_values`.

* **Table extraction**

  * Automatically detect header rows above your data
  * Store table bodies in `sheet_tables`
  * Preserve raw header rows in `sheet_table_headers`
  * Track captions in `table_titles`

* **In-memory workbook dict**
  Everything ends up in one `wb_dictionary` keyed by sheet name/index:

  ```python
  {
    "Sheet1": {
      "sheet_values":         [ DataFrame, … ],
      "sheet_tables":         [ DataFrame, … ],
      "sheet_table_headers":  [ DataFrame, … ],
      "table_titles":         [ "Title1", … ]
    },
    …
  }
  ```

* **Export**

  * `to_html(output_file="wb_dictionary.html")` renders the full dict as nested HTML
  * `save_json(output_file="wb_dictionary.json")` dumps the raw dict to JSON

---

## 📦 Installation

```bash
pip install pandas openpyxl jinja2
```

Copy `utilities.py` into your project and import:

```python
from modules.utilities import ExcelParser
```

---

## ⚙️ Quickstart

```python
from pathlib import Path
from modules.utilities import ExcelParser

# 1) Initialize & load
wb_file = Path("Gas-condensate wellstream template.xlsx")
parser  = ExcelParser(wb_file, sheet_name="Term.Exp.")
parser.load_excel(header=None)

# 2) Extract single-cell values
parser.isTable = False
parser.extract_value_data(rows="2,11", columns="2,5")

# 3) Extract a table with 2 header rows above it
parser.isTable = True
parser.extract_value_data(
    rows="6:9",
    columns="2:4",
    title="Term.exp",
    header=2
)

# 4) Inspect or export
print(parser.wb_dictionary)           # in-memory dict
parser.to_html("output.html")         # HTML file
parser.save_json("output.json")       # JSON dump
```

---

## 📝 API Reference

### `ExcelParser(file_path, sheet_name=0)`

* **file\_path**: `Path` or `str` to `.xlsx`
* **sheet\_name**: sheet name or index

### Properties

* `isTable` — `bool`, toggle between value vs. table mode
* `wb_dictionary` — read-only dict of all extracted data

### Methods

* `load_excel(header=None)`
* `extract_value_data(rows, columns, title=None, header=None)`
* `to_html(output_file="wb_dictionary.html")`
* `save_json(output_file="wb_dictionary.json")`

---

> **Tip:** Batch your `extract_value_data` calls and then export once at the end for a fully-populated `wb_dictionary`.

from jinja2 import Template
import pandas as pd
import ast
import re
import json

class ExcelParser:
    def __init__(self, file_path, sheet_name=0):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.excel = None
        self.data = None
        self.subset = None
        self.table_data_header = None
        self._isTable = False
        self._wb_dictionary = {}
        self._pending_table_title = None
        self._cols_idx = None

    @property
    def isTable(self):
        """Getter for isTable property."""
        return self._isTable

    @isTable.setter
    def isTable(self, value: bool):
        """Setter for isTable property."""
        if not isinstance(value, bool):
            raise ValueError("isTable must be a boolean value.")
        self._isTable = value

    @property
    def wb_dictionary(self):
        """Getter for wb_dictionary property."""
        return self._wb_dictionary

    def load_excel(self, header=None):
        """Load Excel file into self.excel."""
        self.excel = pd.read_excel(
            self.file_path,
            sheet_name=self.sheet_name,
            engine="openpyxl",
            header=header,
        )

    def parse_range(self, range_input) -> slice | list:
        """Parse input to slice or list for iloc indexing."""
        if isinstance(range_input, (slice, list)):
            return range_input
        elif isinstance(range_input, str):
            # Excel-style: e.g., "B2,E11" or "A2:B8"
            if re.match(r'^[A-Z]+\d+(,[A-Z]+\d+)*$', range_input) or re.match(r'^[A-Z]+\d+:[A-Z]+\d+$', range_input):
                cells = range_input.split(":") if ":" in range_input else range_input.split(",")
                if ":" in range_input:
                    start_col, start_row = re.match(r'([A-Z]+)(\d+)', cells[0]).groups()
                    end_col, end_row = re.match(r'([A-Z]+)(\d+)', cells[1]).groups()
                    col_start = sum((ord(c) - 65) * (26 ** i) for i, c in enumerate(reversed(start_col)))
                    col_end = sum((ord(c) - 65) * (26 ** i) for i, c in enumerate(reversed(end_col))) + 1
                    row_start, row_end = int(start_row) - 1, int(end_row) - 1
                    return slice(row_start, row_end)
                else:
                    indices = []
                    for cell in cells:
                        col, row = re.match(r'([A-Z]+)(\d+)', cell).groups()
                        col_idx = sum((ord(c) - 65) * (26 ** i) for i, c in enumerate(reversed(col)))
                        indices.append((int(row) - 1, col_idx))
                    rows, cols = zip(*indices) if indices else ([], [])
                    return list(set(rows)) if range_input == cells[0] else list(set(cols))
            # Numeric list: e.g., "1,10"
            elif "," in range_input:
                try:
                    indices = [int(i) - 1 for i in range_input.split(",")]
                    return indices
                except ValueError:
                    raise ValueError(f"Invalid list format: {range_input}. Use '1,2,3' or '[1,2,3]'.")
            # Numeric slice: e.g., "4:8"
            else:
                try:
                    start, end = map(int, range_input.split(":"))
                    return slice(start - 1, end - 1)  # 1-based to 0-based
                except ValueError:
                    raise ValueError(f"Invalid slice format: {range_input}. Use 'start:end'.")
        else:
            raise ValueError(f"Unsupported range type: {type(range_input)}. Use slice, list, or string ('start:end', '1,2,3', or 'B2,E11').")

    def update_workbook(self):
        """Append extracted data to wb_dictionary for the current sheet."""
        if self.sheet_name not in self._wb_dictionary:
            self._wb_dictionary[self.sheet_name] = {
                "sheet_values": [],
                "sheet_tables": [],
                "sheet_table_headers": [],
                "table_titles": []
            }
        if self.subset is not None:
            if self.isTable:
                if self._pending_table_title is None or self.table_data_header is None:
                    raise ValueError("Table data requires a title and header. Call data_header with a title.")
                # Set first row of header as columns for table
                if not self.table_data_header.empty:
                    header_first_row = self.table_data_header.iloc[0].tolist()
                    self.subset.columns = header_first_row
                self._wb_dictionary[self.sheet_name]["sheet_tables"].append(self.subset)
                self._wb_dictionary[self.sheet_name]["sheet_table_headers"].append(self.table_data_header)
                self._wb_dictionary[self.sheet_name]["table_titles"].append(self._pending_table_title)
                self._pending_table_title = None
            else:
                self._wb_dictionary[self.sheet_name]["sheet_values"].append(self.subset)

    def extract_value_data(self, rows, columns, title: str = None, header=None):
        """Extract subset of DataFrame and append to wb_dictionary."""
        rows_idx = self.parse_range(rows)
        cols_idx = self.parse_range(columns)
        self._cols_idx = cols_idx if not isinstance(cols_idx, slice) else list(range(len(self.excel.columns))[cols_idx])
        self.subset = self.excel.iloc[rows_idx, cols_idx]
        self.subset = self.subset.dropna(how="all").reset_index(drop=True)
        if self.isTable:
            if title is None:
                raise ValueError("Title required for table data extraction.")
            self._pending_table_title = title
            # pull header rows from the ORIGINAL sheet indices
            if header is not None:
                if isinstance(rows_idx, slice):
                    start_row = rows_idx.start
                elif isinstance(rows_idx, list):
                    start_row = min(rows_idx)
                else:
                    raise ValueError("Unsupported rows index for header extraction.")
                header_rows = slice(start_row - header, start_row)
                self.table_data_header = (
                    self.excel
                        .iloc[header_rows, self._cols_idx]
                        .dropna(how="all")
                        .reset_index(drop=True)
                )
            else:
                self.table_data_header = pd.DataFrame()
            self.update_workbook()
        else:
            # non-table: just append the value subset
            self.update_workbook()

    def render_value(self, value):
        if isinstance(value, dict):
            return self.render_dict(value)
        elif isinstance(value, list):
            return "<ul>" + "".join(f"<li>{self.render_value(item)}</li>" for item in value) + "</ul>"
        elif isinstance(value, pd.DataFrame):
            return value.to_html(border=1, index=False)
        else:
            return str(value)

    def render_dict(self, data):
        html = "<ul>"
        for key, value in data.items():
            html += f"<li><strong>{key}</strong>: {self.render_value(value)}</li>"
        html += "</ul>"
        return html

    def to_html(self, output_file="wb_dictionary.html"):
        """
        Render the entire wb_dictionary to an HTML file.
        """
        html_content = self.render_dict(self._wb_dictionary)
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)
    
    def save_json(self, output_file="wb_dictionary.json"):
        """
        Dump the entire wb_dictionary to a JSON file.
        """
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(self._wb_dictionary, f, indent=2, default=str)
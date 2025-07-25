import pandas as pd
import json
from pathlib import Path
from modules import utilities as ut


#
# logic of the info sheet should be separated
# TODO: extract data from info sheet as set of key : [value, note-or-unit] pairs.
#
#previous (non OOP) logic for reading info sheet:
## filter sample info
# my_excel.sheet_name="Info"
# my_excel.load_excel()
# sample_info = my_excel.extract_value_data(rows = "1:11", 
#                                           columns = "0:2")
# field_info = my_excel.extract_value_data(rows = "1:12", 
#                                          columns = "3:6")
# wb[my_excel.sheet_name] = {
#                             "sample_info" : sample_info, 
#                             "field_info" : field_info
#                                 }

wb_file = Path(r"G:/My Drive/_studenti/2025 Jurica Kovacic/excel_templates/Gas-condensate WS template/Gas-condensate wellstream template.xlsx")
my_excel = ut.ExcelParser(wb_file, sheet_name="Term.Exp.")
my_excel.load_excel(header=None)

# Append value_data
my_excel.isTable = False
my_excel.extract_value_data(rows="2,11", columns="2,5")  # Rows 2,11; cols B,E

# Append table_data
my_excel.isTable = True
my_excel.extract_value_data(rows="6:9", columns="2:4",
                            title="Term.exp", 
                            header = 2
                            )

wb_dictionary = my_excel.wb_dictionary
my_excel.to_html(output_file = "dict_output.html")
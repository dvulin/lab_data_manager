import pandas as pd
import json
from pathlib import Path

def dict_to_html(my_dict):
    from jinja2 import Template
    template = Template("<ul>{% for key, value in data.items() %}<li>{{ key }}: {{ value }}</li>{% endfor %}</ul>")
    html_content = template.render(data=my_dict)
    with open("dictionary.html", "w") as text_file:
        text_file.write(html_content)
    return html_content

# Workbook file
#wb_file = r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx"
wb_file = Path(r"G:/My Drive/_studenti/2025 Jurica Kovacic/excel_templates/Gas-condensate WS template/Gas-condensate wellstream template.xlsx")

wb = {}

### 

# filter sample info
sheet="Info"
df = pd.read_excel(wb_file, sheet_name=sheet)
sample_info = df.iloc[1:11, 0:2].dropna()                       # tablica tipova value_data
field_info = df.iloc[1:12, 3:6].dropna(axis=0, how = 'all')     # tablica tipa value_data
wb[sheet] = {
            "sample_info" : sample_info, 
            "field_info" : field_info
            }

sheet="Term.Exp."
df = pd.read_excel(wb_file, sheet_name=sheet)
value_data = df.iloc[[1, 10],[1,4]]

table_title = sheet
table_data_header = df.iloc[2:4,1:3]
table_data = df.iloc[4:8,1:3]
wb[sheet] = {
            "sheet_values" : [value_data],
            "table_titles" : [table_title],
            "sheet_table_headers" : [table_data_header],
            "sheet_tables" : [table_data]
            }

sheet="CCE 1"
df = pd.read_excel(wb_file, sheet_name=sheet)
value_data = df.iloc[[1, 10],[1,4]]


print (dict_to_html(wb))





# subset = df.iloc[1:15, 0:6]

# # Ukloni sve `NaN` vrijednosti iz redova (ali ne briše red)
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# # Term.Exp.
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Term.Exp."
# )

# subset = df.iloc[1:11, 0:6]

# # Ukloni sve `NaN` vrijednosti iz redova (ali ne briše red)
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# # CCE1
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="CCE 1"
# )

# subset = df.iloc[1:44, 0:16]

# # Pretvori sve vrijednosti u string (uključuje datetime, NaN, float itd.)
# subset = subset.astype(str)
# subset = subset.replace("nan", "")
# # Pretvori u listu redova
# data = subset.values.tolist()
# # Ukloni prazne redove (ako su svi elementi prazni)
# data = [row for row in data if any(cell.strip() for cell in row)]
# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #CCE Flash
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="CCE Flash"
# )

# subset = df.iloc[0:25, 0:12]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Flash Gas Comp
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Flash Gas Comp"
# )
# subset = df.iloc[0:35, 0:5]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Sep Gas Comp
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Sep Gas Comp"
# )
# subset = df.iloc[0:35, 0:5]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Sep Liquid 
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Sep Liquid"
# )
# subset = df.iloc[0:47, 0:17]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Well stream 
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Well stream"
# )
# subset = df.iloc[0:47, 0:13]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Report Comp 
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Report Comp"
# )
# subset = df.iloc[0:60, 0:14]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)

# #Report Volumetric 
# df = pd.read_excel(
#     r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/Gas-condensate wellstream template.xlsx",
#     sheet_name="Report Volumetric"
# )
# subset = df.iloc[0:13, 0:4]
# data = [[cell for cell in row if pd.notna(cell)] for row in subset.values.tolist()]
# data = [row for row in data if row]  # izbaci potpuno prazne redove

# json_data = json.dumps(data, indent=4, ensure_ascii=False)
# print(json_data)
import pandas as pd
import json
from pathlib import Path
from modules import utilities as ut

wb_file = Path(r"C:/Users/Lenovo/Desktop/Faks/Diplomski/excel_templates/Gas-condensateWStemplate/GUI/Gas-condensate wellstream template.xlsx")

wb_dictionary = {}

#Sheet1: Info

info = ut.ExcelParser(wb_file, sheet_name="Info")
info.load_excel(header=None)

#Append value_data
info.isTable = False
info.extract_value_data(rows= "3,5,7,8,10,11", columns="1,2", title="sample_info")
info.extract_value_data(rows= "3,5,7,8,11,12,13,15", columns="4,5,6", title="field_info")
                        
wb_dictionary["Info"] = info.wb_dictionary["Info"]


#Sheet2: Term.Exp.

term_exp = ut.ExcelParser(wb_file, sheet_name="Term.Exp.")
term_exp.load_excel(header=None)

# Append value_data
term_exp.isTable = False
term_exp.extract_value_data(rows="2,11", columns="2,5")  # Rows 2,11; cols B,E

# Append table_data
term_exp.isTable = True
term_exp.extract_value_data(rows="6:9", columns="2:4",
                            title="Term.exp", 
                            header = 2
                            )

wb_dictionary["Term.Exp."] = term_exp.wb_dictionary["Term.Exp."]

#Sheet3: CCE 1

cce1 = ut.ExcelParser(wb_file, sheet_name="CCE 1")
cce1.load_excel(header=None)

#Append value_data
cce1.isTable = False
cce1.extract_value_data(rows="1,3", columns="4,5,6,8,9", title="Osnovni podaci o CCE 1 testu")

#Append table_data

cce1.isTable = True
cce1.extract_value_data(rows="8:24", columns="1,2,3,4,10,11,12,13",
                        title="Pcor,Vcor",
                        header = 2
                        )
cce1.extract_value_data(rows="8:13", columns="15,16",
                        title="Compressibility calc",
                        header = 2
                        )
cce1.extract_value_data(rows="42:45", columns="2:6",
                        title="Računsko određivanje Pb-a",
                        header = 2
                        )

wb_dictionary["CCE 1"] = cce1.wb_dictionary["CCE 1"]

#Sheet4: CCE Flash


cceflash = ut.ExcelParser(wb_file, sheet_name="CCE Flash")
cceflash.load_excel(header=None)

#Append value data
cceflash.isTable = False
cceflash.extract_value_data(rows="1,2", columns="4,5,6", title="Temperatura ćelije")
cceflash.extract_value_data(rows="12,13,15,17,19,21,23,25", columns="10,11,12", title="Flash test podaci")

#Append table data

cceflash.isTable = True
cceflash.extract_value_data(rows="10:26", columns="1:7",
                            title="Izbacivanje plina",
                            header = 2 
                            )

wb_dictionary["CCE Flash"] = cceflash.wb_dictionary["CCE Flash"]

#Sheet5: Flash Gas Comp

flashgascomp = ut.ExcelParser(wb_file, sheet_name="Flash Gas Comp")
flashgascomp.load_excel(header=None)

#Append value data
flashgascomp.isTable = False
flashgascomp.extract_value_data(rows="28,29,30,32,33,34,35", columns="3,4,5")

#Append table data
flashgascomp.isTable = True
flashgascomp.extract_value_data(rows="8:27", columns="3:6",
                                title="Kromatografska analiza plina [ISO6975]",
                                header = 2 
                                )

wb_dictionary["Flash Gas Comp"] = flashgascomp.wb_dictionary["Flash Gas Comp"]

#Sheet6 : Sep Gas Comp

sepgascomp = ut.ExcelParser(wb_file, sheet_name="Sep Gas Comp")
sepgascomp.load_excel(header=None)

#Append value data
sepgascomp.isTable = False
sepgascomp.extract_value_data(rows="28,29,30,32,33,34,35", columns="3,4,5")

#Append table data
sepgascomp.isTable = True
sepgascomp.extract_value_data(rows="8:27", columns="3:6",
                              title="Kromatografska analiza plina [ISO6975]",
                              header = 2 
                              )

wb_dictionary["Sep Gas Comp"] = sepgascomp.wb_dictionary["Sep Gas Comp"]

#Sheet7 : Sep Liquid

sepliquid = ut.ExcelParser(wb_file, sheet_name="Sep Liquid")
sepliquid.load_excel(header=None)

#Append value data
sepliquid.isTable = False
sepliquid.extract_value_data(rows="2,3,4,7,8,11,12,13,15,16,18,19", columns="1,2")

#Append table data
sepliquid.isTable = True
sepliquid.extract_value_data(rows="2:47", columns="4,5,6,7,8,9,12,13,15,16,17",
                             title="Stock Tank Liquid + Flash Gas (Sep Liquid or Downhole Sample)",
                             header = 1 
                             )

wb_dictionary["Sep Liquid"] = sepliquid.wb_dictionary["Sep Liquid"]

#Sheet8 : Well stream

wellstream = ut.ExcelParser(wb_file, sheet_name="Well stream")
wellstream.load_excel(header=None)

#Append value data
wellstream.isTable = False
wellstream.extract_value_data(rows="2,3,4,5,6,8,10,11,13,15,17,19,20", columns="1,2")

#Append table data
wellstream.isTable = True
wellstream.extract_value_data(rows="2:47", columns="4,5,6,8,9,11,12,13",
                              title="Well stream composition",
                              header = 1 
                              )

wb_dictionary["Well stream"] = wellstream.wb_dictionary["Well stream"]

#Sheet9: Report Comp

reportcomp = ut.ExcelParser(wb_file, sheet_name="Report Comp")
reportcomp.load_excel(header=None)

#Append value data
reportcomp.isTable = False
reportcomp.extract_value_data(rows="26,27", columns="1,2", title="C7+ properties")
reportcomp.extract_value_data(rows="59,60", columns="1,2", title="C10+ properties")

#Append table data
reportcomp.isTable = True
reportcomp.extract_value_data(rows="1:21", columns="1:7",
                              title="Do C7+ frakcije",
                              header = 2
                              )
reportcomp.extract_value_data(rows="30:53", columns="1:7",
                              title="Do C10+ frakcije",
                              header = 2 
                              )
reportcomp.extract_value_data(rows="1:50", columns="9:15",
                              title="Do C40+ frakcije",
                              header = 2
                              )

wb_dictionary["Report Comp"] = reportcomp.wb_dictionary["Report Comp"]

#Sheet10: Report Volumetric

reportvolumetric = ut.ExcelParser(wb_file, sheet_name="Report Volumetric")
reportvolumetric.load_excel(header=None)

#Append value data
reportvolumetric.isTable = False
reportvolumetric.extract_value_data(rows="2:14", columns="1:5")

wb_dictionary["Report Volumetric"] = reportvolumetric.wb_dictionary["Report Volumetric"]





# ========== Zajednički HTML export ==========
def render_value(value):
    if isinstance(value, dict):
        return render_dict(value)
    elif isinstance(value, list):
        return "<ul>" + "".join(f"<li>{render_value(item)}</li>" for item in value) + "</ul>"
    elif isinstance(value, pd.DataFrame):
        return value.fillna("").to_html(border=1, index=False)
    else:
        return str(value)

def render_dict(data):
    html = "<ul>"
    for key, value in data.items():
        html += f"<li><strong>{key}</strong>: {render_value(value)}</li>"
    html += "</ul>"
    return html

html_content = "<html><body><h1>Svi sheetovi</h1>"
for sheet, data in wb_dictionary.items():
    html_content += f"<h2>{sheet}</h2>{render_dict(data)}<hr>"
html_content += "</body></html>"

with open("dict_output.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✅ Podaci svih listova spremljeni u dict_output.html")



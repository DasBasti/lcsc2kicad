###
# LCSC Part Database parser
###
from xlrd import open_workbook
import os

columns = [
    "LCSC",
    "FirstCategory",
    "SecondCategory",
    "MFRPart",
    "Package",
    "SolderJoint",
    "Manufacturer",
    "LibraryType",
    "Description",
    "Datasheet",
    "Price",
    "Stock"
]

data = {}

### Load file into buffer
def load_data(file):
    book = open_workbook(file_contents=file, on_demand=True)
    for name in book.sheet_names():
        if name == 'JLCPCB SMT Parts Library':
            sheet = book.sheet_by_name(name)
            names = []
            for rc in range(sheet.nrows):
                if rc == 0: # skip header
                    continue
                cat = sheet.cell_value(rowx=rc, colx=1)
                if cat not in data:  # library files
                    data[cat] = {}
                
                grp = sheet.cell_value(rowx=rc, colx=2)
                grp_field = grp.replace("/", "")
                if grp_field not in data[cat]:  # library files
                    data[cat][grp_field] = []

                fields = {"field": grp_field}
                for c,f in enumerate(sheet.row(rc)):
                    value = f.value
                    value = value.replace("\"", "")
                    if columns[c] == "MFRPart": # part name
                        if value in names:
                            value += "_"
                        names.append(value)
                        
                        fields["MFRPart_sanitized"] = sanitise_name(value)
                    if columns[c] == "Description":
                        value = value.replace(grp, "").strip()
                    fields[columns[c]] = value
                data[cat][grp_field].append(fields)
        book.unload_sheet(name)


def sanitise_name(name, spaces=True):
    if spaces:
        name = name.replace(" ","")
    return name.replace("&", "").replace("(", "").replace(")","").replace(",","").replace("/", "")

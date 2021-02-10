import table
import os
import glob
import requests
from xlrd import open_workbook

dcm_templates = {}
lib_templates = {}

if not os.path.exists("export"):
    os.mkdir("export")

print("load: DCM template")
with open("templates/template.dcm") as f:
    dcm_template = f.read()

for filepath in glob.iglob("templates/*.lib"):
    print("load: "+filepath)
    with open(filepath) as f:
        lib_templates[os.path.basename(filepath).split('.')[0]] = f.read()

print("fetch file from JLCPCB")
url = 'https://jlcpcb.com/componentSearch/uploadComponentInfo'
r = requests.get(url)

print("parsing data. This might take a while!")

table.load_data(r.content)

parts = 0
for lib in table.data:
    lib_parts = 0
    print("Library: "+lib)
    file_pre = (lib.replace(" ","_").replace("&","").replace("__","_"))
    dcm_str = "EESchema-DOCLIB  Version 2.0\n"
    lib_str = "EESchema-LIBRARY Version 2.4\n"
    lib_str += "#encoding utf-8\n"
    for grp in table.data[lib]:
        print("Group: "+grp)
        for part in table.data[lib][grp]:
            template = table.sanitise_name(part['SecondCategory'], spaces=False)
            if template in lib_templates:
                parts += 1
                lib_parts += 1
                lib_str += lib_templates[template].format(**part)
                dcm_str += dcm_template.format(**part)

    lib_str += "#\n#End Library"
    dcm_str += "#\n#End Doc Library"

    if(lib_parts > 0):
        with open("export/LCSC_"+file_pre+".lib", "w") as l:
            l.write(lib_str)
        with open("export/LCSC_"+file_pre+".dcm", "w") as l:
            l.write(dcm_str)

print("Total: {0} parts".format(parts))
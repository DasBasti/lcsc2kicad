import table
import os
import sys
import glob
import requests
import json
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
        content = f.read()
        lines = content.splitlines()
        header = {}
        if lines[0].startswith("#!"): # parse header information
            header = json.loads(lines[0][2:])
        header['content'] = "\n".join(lines[1:])
        lib_templates[os.path.basename(filepath).split('.')[0]] = header

if len(sys.argv) == 1:
    print("fetch file from JLCPCB")
    url = 'https://jlcpcb.com/componentSearch/uploadComponentInfo'
    r = requests.get(url)
    db = r.content
else:
    print("load from: ", sys.argv[1])
    with open(sys.argv[1], "rb") as db_file:
        db = db_file.read()


print("parsing data. This might take a while!")

table.load_data(db)

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
            # Test for MPN Template
            template = table.sanitise_name(part['MFRPart'], spaces=False)
            if template in lib_templates:
                if lib_templates[template]['pins'] == part['SolderJoint']:
                    parts += 1
                    lib_parts += 1
                    lib_str += lib_templates[template]['content'].format(**part)
                    dcm_str += dcm_template.format(**part)
                else:
                    print("Pins missmatch", part['SolderJoint'], template)
            else: # Test for generic Group Template
                template = table.sanitise_name(part['SecondCategory'], spaces=False)
                if template in lib_templates:
                    if lib_templates[template]['pins'] == part['SolderJoint']:
                        parts += 1
                        lib_parts += 1
                        lib_str += lib_templates[template]['content'].format(**part)
                        dcm_str += dcm_template.format(**part)
                    else:
                        print("Group Pins missmatch", part['SolderJoint'], part['MFRPart'])

    lib_str += "#\n#End Library"
    dcm_str += "#\n#End Doc Library"

    if(lib_parts > 0):
        with open("export/LCSC_"+file_pre+".lib", "w") as l:
            l.write(lib_str)
        with open("export/LCSC_"+file_pre+".dcm", "w") as l:
            l.write(dcm_str)

print("Total: {0} parts".format(parts))
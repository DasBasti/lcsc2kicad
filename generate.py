import table
import os
import glob
from xlrd import open_workbook

dcm_templates = {}
lib_templates = {}

for filepath in glob.iglob("templates/*.dcm"):
    print("load: "+filepath)
    with open(filepath) as f:
        dcm_templates[os.path.basename(filepath).split('.')[0]] = f.read()

for filepath in glob.iglob("templates/*.lib"):
    print("load: "+filepath)
    with open(filepath) as f:
        lib_templates[os.path.basename(filepath).split('.')[0]] = f.read()


table.load_data("JLCPCB SMT Parts Library(20210203).xls")

parts = 0
for lib in table.data:
    print("Library: "+lib)
    file_pre = (lib.replace(" ","_").replace("&","").replace("__","_"))
    dcm_str = "EESchema-DOCLIB  Version 2.0\n"
    lib_str = "EESchema-LIBRARY Version 2.4\n"
    lib_str +="#encoding utf-8\n"
    for grp in table.data[lib]:
        print("Group: "+grp)
        for part in table.data[lib][grp]:
            template = table.sanitise_name(part['SecondCategory'], spaces=False)
            if template in lib_templates:
                parts+=1
                lib_str += lib_templates[template].format(**part)
                dcm_str += dcm_templates[template].format(**part)

    lib_str += "#\n#End Library"
    dcm_str += "#\n#End Doc Library"

    with open("export/"+file_pre+".lib", "w") as l:
        l.write(lib_str)
    with open("export/"+file_pre+".dcm", "w") as l:
        l.write(dcm_str)

print("Total: {0} parts".format(parts))
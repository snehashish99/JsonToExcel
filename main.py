import xlwt
import json

import_filename = 'JSON_adv_seg.jsonc'
data = open(import_filename, 'r').read()
data=json.loads(data)

# Mapped column names with column numbers
columns={
    "name": 0,
    "category": 1, 
    "sub-category": 2, 
    "filter name": 3, 
    "actual name": 4, 
    "field type": 5, 
    "filter_title" :6, 
    "filter_values": 7
    }
rowCount = {
    "name": 1,
    "category": 1, 
    "sub-category": 1, 
    "filter name": 1, 
    "actual name": 1, 
    "field type": 1, 
    "filter_title" : 1, 
    "filter_values": 1,
}

# Create workbook and sheet
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Datas", cell_overwrite_ok=True)

# Fill the heading
for i in columns:
    sheet.write(0,columns[i],i)

# Variable to keep track of current row
curr_row=1 

for obj in data:
    sheet.write(curr_row, columns["name"], obj["name"])
    
    template=obj["template"]
    if(len(template["filters"])==0):
        categories=template["sub_cat"]
        for category in categories:
            sheet.write(curr_row, columns["category"], category["title"])
            if(len(category["filters"])==0):
                for sub_cat in category["sub_cat"]:
                    sheet.write(curr_row, columns["sub-category"], sub_cat["title"])
                    for filter in sub_cat["filters"]: 
                        sheet.write(curr_row, columns["filter name"], filter["filter_name"])
                        sheet.write(curr_row, columns["field type"], filter["field_type"])
                        sheet.write(curr_row, columns["filter_title"], filter["title"])                                               
                        for value in filter["options"][0]["values"]:
                            sheet.write(curr_row, columns["filter_values"], value)
                            curr_row+=1
            else:
                for filter in category["filters"]:
                    sheet.write(curr_row, columns["filter name"], filter["filter_name"])
                    sheet.write(curr_row, columns["field type"], filter["field_type"])
                    sheet.write(curr_row, columns["filter_title"], filter["title"])
                    for value in filter["options"][0]["values"]:
                        sheet.write(curr_row, columns["filter_values"], value)
                        curr_row+=1 
    else:
        for filter in template["filters"]:
            sheet.write(curr_row, columns["filter name"], filter["filter_name"])
            sheet.write(curr_row, columns["field type"], filter["field_type"])
            sheet.write(curr_row, columns["filter_title"], filter["title"])
            for value in filter["options"][0]["values"]:
                sheet.write(curr_row, columns["filter_values"], value)
                curr_row+=1 

workbook.save("output.xls")

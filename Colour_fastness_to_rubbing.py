from create_test import create_test
from docx.enum.section import WD_SECTION
from docx.shared import Inches
from global_func import set_col_widths
def Color_fr_test_template(doc,test_name,test_method,remarks,title,result,req):
    create_test(doc, test_name, test_method, remarks)
    keys_list = list(result.keys())
    for k in range(round(len(keys_list)/2)):
        if k%4==0 and k !=0:
            doc.add_section(WD_SECTION.NEW_PAGE)
        item_list = []
        num_col = 6
        if len(result) < 2:
            num_col =4
            item_list.append(result.pop(keys_list[0]))
        else:
            item_list.append(result.pop(keys_list[0]))
            item_list.append(result.pop(keys_list[1]))

        sample_heading = doc.add_paragraph()
        table = doc.add_table(rows=len(title), cols=num_col)
        table.style = 'Table Grid'
        # Add title list to the first cell of each row
        for row_index, item in enumerate(title):
            cell = table.cell(row_index, 0)
            cell.text = item

        for i in range(0,len(item_list)+1,2):
            temp = keys_list.pop(0)
            a = table.cell(0,i+1).merge(table.cell(0,i+2))
            a.text = temp
            table.cell(1,i+1).text = 'Length'
            table.cell(1,i+2).text = 'Width'

        for i in range(1, len(item_list) + 2, 2):
            temp = list(item_list.pop(0).values())[0]
            for j in range(2,4):
                table.cell(j,i).text = temp[j-2]
                table.cell(j, i+1).text = temp[j-2]
        table.cell(2,num_col-1).text = req[0]
        table.cell(3, num_col-1).text = req[1]
        widths = (Inches(2), Inches(1), Inches(1), Inches(1), Inches(1), Inches(1.5))
        set_col_widths(table, widths)
        for row in table.rows:
            row.cells[-1].width = Inches(1.5)
def Color_fr_test_data():

    title = ["Sample","",
        "Dry rubbing"
        ,"Wet rubbing"
    ]

    values = {
        'Sample A': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample B': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample C': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample D': {'Length':['4-5','3'],'Width':['4-5','3']},
        'Sample E': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample F': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample G': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample H': {'Length':['4-5','3'],'Width':['4-5','3']},
         'Sample I': {'Length':['4-5','3'],'Width':['4-5','3']},
    }
    test_name = 'Colour fastness to rubbing:'
    test_method = "DIN EN ISO 105-X12:2016"
    remarks  = "Analyzed by LC-MS"

    return test_name,test_method,remarks,title,values
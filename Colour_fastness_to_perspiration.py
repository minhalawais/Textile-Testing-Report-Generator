from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from create_test import create_test
from docx import Document
from docx.shared import Inches
from create_test import create_test
from global_func import set_col_widths

def color_fp_test_template(doc, test_name, test_method, remarks, title, result,req):
    create_test(doc, test_name, test_method, remarks)
    keys_list = list(result.keys())
    for k in range(round(len(keys_list)/2)):
        if k % 2 == 0 and k!=0:
            doc.add_section(WD_SECTION.NEW_PAGE)
            create_test(doc, test_name, test_method, remarks)
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
            table.cell(0,i+1).text = temp
            table.cell(0,i+2).text = temp
            table.cell(1,i+1).text = 'Acid'
            table.cell(1,i+2).text = 'Alkane'

        for i in range(1, len(item_list) + 2, 2):
            temp = list(item_list.pop(0).values())
            temp1 = temp[1]
            temp = temp[0]


            for j in range(2,11):
                table.cell(j,i).text = temp[j-2]
                table.cell(j, i+1).text = temp1[j-2]
        table.cell(2,num_col-1).text = req[0]
        table.cell(3, num_col-1).text = req[1]
        a = table.cell(4, num_col-1)
        for i in range(4, 11):
            b = table.cell(i, num_col-1)
            A = a.merge(b)
            a = A
        a.text = req[2]
        widths = (Inches(2), Inches(1), Inches(1), Inches(1), Inches(1), Inches(1.5), Inches(1.5), Inches(1.5))
        set_col_widths(table, widths)
        for row in table.rows:
            row.cells[-1].width = Inches(2)
def color_fp_test_data():
    title = ['Sample',"",'Color Change','Self-staining','Staining On:',' - Acetate',' - Cotton'	,' - Polyamide',' - Polyester',' - Acrylic	',' - Wool']
    values = {'Sample A': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']}
        , 'Sample B': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']},
              'Sample C': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']}
        , 'Sample D': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']},
              'Sample E': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']}
        , 'Sample F': {'Acid':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],'Alkane':['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']},
           'Sample G': {'Acid': ['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5'],
                         'Alkane': ['4-5', '4-5', '-', '4-5', '4-5', '4-5', '4-5', '4-5', '4-5']}
               }

    test_name = "Colour fastness to perspiration:"
    test_method = "DIN EN ISO 105-E04"
    remarks = ''
    return test_name,test_method,remarks,title,values





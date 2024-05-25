from create_test import create_test
from docx.enum.section import WD_SECTION
from docx.shared import Inches
from global_func import set_col_widths

def fabric_test_template(doc,test_name,test_method,remarks,title,values,req):
    create_test(doc, test_name, test_method, remarks)
    max_samples_per_table = 4

    num_tables = (len(values) + max_samples_per_table - 1) // max_samples_per_table

    for table_index in range(num_tables):
        start_sample = table_index * max_samples_per_table
        end_sample = start_sample + max_samples_per_table

        num_samples = min(max_samples_per_table, len(values) - start_sample)
        num_columns = num_samples + 2  # Including the last column

        table = doc.add_table(rows=4, cols=num_columns)
        table.style = 'Table Grid'

        # Populate the first row with the sample names
        for i, sample in enumerate(list(values.keys())[start_sample:end_sample]):
            cell = table.cell(0, i + 1)
            cell.text = sample

        # Populate the second row with the first set of values
        for i, data in enumerate(list(values.values())[start_sample:end_sample]):
            cell = table.cell(1, i + 1)
            cell.text = data[0]

        # Populate the third row with the second set of values
        for i, data in enumerate(list(values.values())[start_sample:end_sample]):
            cell = table.cell(2, i + 1)
            cell.text = data[1]

        # Populate the last column with the additional values
        column_index = num_columns - 1
        table.cell(0, column_index).text = req[0]
        table.cell(1, column_index).text = req[1]
        table.cell(2, column_index).text = req[2]

        # Populate the first column with the title
        for i, value in enumerate(title):
            cell = table.cell(i, 0)
            cell.text = value

        # Add a section break if there are more tables
        if (table_index + 1) %4 == 0 and table_index !=0:
            doc.add_section(WD_SECTION.NEW_PAGE)
        sample_heading = doc.add_paragraph()
        widths = (Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5))
        set_col_widths(table, widths)

def fabric_test_data():
    title = ['Sample','g/m²','Oz']

    values = {
        'Sample A': ['203.0','5.99'],
        'Sample B': ['203.0', '5.99'],
        'Sample C': ['203.0','5.99'],
        'Sample D': ['203.0', '5.99'],
        'Sample E': ['203.0','5.99'],
        'Sample F': ['203.0', '5.99'],
        'Sample G': ['203.0','5.99/'],
        'Sample H': ['203.0','5.99'],
        'Sample I': ['203.0','5.99'],
        'Sample J': ['203.0','5.99'],
        'Sample K': ['203.0','5.99'],
    }
    test_name = "Fabric weight:"
    test_method = "ISO 3801:1977"
    remarks = ""
    return test_name,test_method,remarks,title,values
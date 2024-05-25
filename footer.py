import win32com.client as win32
from docx.shared import Inches,Cm
import pythoncom

def dispatch():
    try:
        from win32com import client
        app = client.gencache.EnsureDispatch('Word.Application')
    except:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        app = client.gencache.EnsureDispatch('Word.Application')
    return app
def set_footer(doc=r"test_template.docx"):
    # Connect to an existing instance of Word
    pythoncom.CoInitialize()  # Initialize the COM library

    try:
        word = dispatch()
    except:
        from win32com import client
        word = client.gencache.EnsureDispatch('Word.Application')
    # Open an existing Word document
    doc = word.Documents.Open(doc)

    # Access the footer of the first section
    footer = doc.Sections(1).Footers(1)

    # Add a shape to the footer
    shape = footer.Shapes.AddPicture(r'Logo\footer.png')

    # Set the wrap format of the shape to floating
    shape.WrapFormat.Type = win32.constants.wdWrapSquare
    shape.WrapFormat.AllowOverlap = True
    shape.WrapFormat.Side = win32.constants.wdWrapBoth



    # Set the width of the shape to match the page width


# Set the width of the shape to match the page width in centimeters
    shape.Width = 620 # 567 twips = 1 cm

    # Save and close the document
    doc.Save()
    doc.Close()

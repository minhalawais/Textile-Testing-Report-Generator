# Project Name: Test Report Generator

## Overview

The Test Report Generator is a Python script that automates the creation of comprehensive test reports in DOCX format. It is designed to streamline the process of generating structured and formatted test reports based on provided data.

The main features of this project include:
- Dynamic generation of tables with specified column widths and data.
- Insertion of images, QR codes, and signatures into the document.
- Customizable styling for font, size, alignment, and table borders.
- Integration with external modules for QR code generation and global functions.

## Installation

1. Clone the repository to your local machine:
```sh
git clone https://github.com/minhalawais/Textile-Testing-Report-Generator
```

2. Navigate to the project directory:
```sh
cd Textile-Testing-Report-Generator
```

3. Install the required dependencies using pip:
```sh
pip install python-docx
```


5. Ensure that you have the necessary image files (e.g., logos, signatures) in the specified directories.

## Usage

1. Open the `generate_report.py` script in your preferred code editor.

2. Update the data dictionaries (`applicant_dict`, `buyer_dict`, `sample_dict`, etc.) with the relevant information for your test report.

3. Customize the script as needed for specific formatting, additional sections, or data manipulation.

4. Run the script to generate the test report document:
```sh
python generate_report.py
```


5. The script will create a DOCX file named `test_template.docx` in the project directory, containing the formatted test report.

## File Structure

- `generate_report.py`: Main Python script for generating the test report document.
- `generate_qrcode.py`: Module for generating QR codes with images.
- `global_func.py`: Module containing global functions and data.
- `images/`: Directory containing image files used in the document.
- `README.md`: Project documentation in Markdown format.

## Additional Notes

- Ensure that all required modules (`docx`, `generate_qrcode`, `global_func`) are installed and accessible.
- Customize the image files and data dictionaries according to your specific requirements.
- Test the generated DOCX file to verify formatting, image placements, and overall content alignment.
- Refer to the `README.md` file for detailed installation instructions and usage guidelines.
- For any issues or further customization needs, refer to the source code comments or contact minhalawais1@gmail.com

# DOCX to LaTeX Converter

# DOCX to LaTeX Converter

This project provides a Python script to convert a Microsoft Word document (.docx) into a LaTeX document. The script handles various elements such as text, headings, images, tables, and equations, and generates a zipped folder containing the LaTeX project, ready for upload to Overleaf.

## Features

- Extracts and converts text, headings, images, tables, and equations from a Word document.
- Saves all images from the Word document into a specified directory.
- Creates a LaTeX document with proper formatting.
- Generates a ZIP file containing the LaTeX project for easy upload to Overleaf.
- Displays progress bars for different stages of processing (paragraphs, tables, images, and equations).
- Automatically converts special characters to LaTeX format.
- Logs any errors encountered during the conversion process, especially missing images.

## Requirements

- Python 3.x
- `python-docx`
- `pylatex`
- `pylatexenc`
- `tqdm`

## Installation

1. Clone the repository or download the script file.
2. Install the required Python packages:
    ```bash
    pip install python-docx pylatex pylatexenc tqdm
    ```

## Usage

### Instructions for Running the Script:

1. **Place your Word document in a known location.**

2. **Run the Script from the Terminal:**
   Use the following command format to run the script:
   ```bash
   python main.py <path_to_docx_file> <output_directory>


## Example

`python docx_to_latex_converter.py /path/to/your/document.docx /path/to/output/directory`


## Output:
The script will process the DOCX file, generate a LaTeX document, and create a ZIP file in the specified output directory. Progress bars will show the processing stages for paragraphs, tables, images, and equations.

Example: 

docx_path = 'path/to/your/document.docx'
output_dir = 'path/to/output/directory'
converter = DocxToLatexConverter(docx_path, output_dir)
converter.convert()

## Script Details
Class: DocxToLatexConverter
Methods:
__init__(self, docx_path, output_dir): Initializes the converter with the DOCX file path and the output directory.
extract_images(self): Extracts images from the DOCX file and saves them in the images directory.
handle_paragraph(self, paragraph): Processes paragraphs and converts them to LaTeX format.
handle_table(self, table): Processes tables and converts them to LaTeX format.
handle_equation(self, equation): Processes equations and converts them to LaTeX format.
convert(self): Orchestrates the conversion process, handling text, images, tables, and equations.
create_zip(self): Creates a ZIP file of the LaTeX project.

## Output 
output_dir/
└── latex/
    ├── document.tex
    ├── images/
    │   ├── image0.png
    │   ├── image1.png
    │   └── ...
    └── latex_project.zip

- document.tex: The LaTeX document generated from the Word document.
- images/: Directory containing all images extracted from the Word document.
- latex_project.zip: ZIP file containing the LaTeX project for easy upload to Overleaf.
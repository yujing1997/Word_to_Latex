# DOCX to LaTeX Converter

This project provides a Python script to convert a Microsoft Word document (.docx) into a LaTeX document. The script handles various elements such as text, headings, images, tables, and equations, and generates a zipped folder containing the LaTeX project, ready for upload to Overleaf.

## Features

- Extracts and converts text, headings, images, tables, and equations from a Word document.
- Saves all images from the Word document into a specified directory.
- Creates a LaTeX document with proper formatting.
- Generates a ZIP file containing the LaTeX project for easy upload to Overleaf.
- Displays progress bars for different stages of processing (paragraphs, tables, images, and equations).

## Requirements

- Python 3.x
- `python-docx`
- `pylatex`
- `tqdm`

## Installation

1. Clone the repository or download the script file.
2. Install the required Python packages:
    ```bash
    pip install python-docx pylatex tqdm
    ```

## Usage

1. Place your Word document in a known location.

2. Update the `docx_path` and `output_dir` variables in the script with the path to your Word document and the desired output directory, respectively.

3. Run the script:
    ```bash
    python docx_to_latex_converter.py
    ```

## Example

```python
docx_path = 'path/to/your/document.docx'
output_dir = 'path/to/output/directory'
converter = DocxToLatexConverter(docx_path, output_dir)
converter.convert()

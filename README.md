# PDF to DOCX Converter

## Overview
This project provides a GUI-based tool to convert PDF files into DOCX format while preserving the layout and formatting. It supports batch processing, automatic hyperlink detection, and text alignment for better document presentation.

## Features
- **Drag and Drop GUI**: Easily select multiple PDF files for conversion.
- **Automatic Formatting**: Adjusts text alignment and converts detected hyperlinks.
- **Batch Processing**: Supports multiple PDF conversions in a single session.
- **User-friendly Interface**: Simple and intuitive design with message alerts.
- **Preserves Layout**: Maintains original document formatting as much as possible.

## Installation

### Prerequisites
Ensure you have Python installed on your system. Install the required dependencies using:
```sh
pip install pdfplumber pdf2docx python-docx tk
```

## Usage
1. Run the script:
   ```sh
   python convert.py
   ```
2. A GUI window will open.
3. Click the **Browse PDF Files** button to select multiple PDFs.
4. The tool will automatically convert and format the files.
5. A success message will appear with the saved file locations.

## Technologies Used
- **Python**: Core programming language.
- **pdfplumber**: Extracts text from PDFs.
- **pdf2docx**: Converts PDFs to DOCX.
- **python-docx**: Formats and modifies DOCX files.
- **Tkinter**: GUI interface for file selection and processing.

## Contributing
Feel free to fork this repository, create a new branch, and submit a pull request with your improvements!

## License
This project is licensed under the MIT License.


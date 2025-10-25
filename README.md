# Comment Pdf From Excel

A GUI-based utility to add comments from an Excel file to PDF documents. The tool searches for specific tags in PDF files and adds annotations with corresponding comments at those locations.

## Features
- **Fully GUI-based interface** - No command-line arguments needed
- **Batch processing** - Process multiple PDF files in a folder at once
- **Configurable annotation distance** - Set the spacing between tags and annotations via GUI
- **Same-folder output** - Annotated PDFs are saved in the same location as input files with "updated_" prefix
- **Customizable comment subject** - Set the subject for annotations
- **Excel-driven annotations** - Define tags and comments in a simple Excel spreadsheet

## Requirements
- Python 3.8+ recommended
- PyMuPDF (fitz)
- pandas
- openpyxl (for reading .xlsx files)
- tkinter (usually included with Python)

## Installation
1. Clone the repository:
```bash
git clone https://github.com/apkarthik1986/CommentPdfFromExcel.git
cd CommentPdfFromExcel
```

2. Install required dependencies:
```bash
pip install PyMuPDF pandas openpyxl
```

## Usage

### Excel File Format
Create an Excel file (.xlsx or .xls) with the following columns:
- **tag**: The text to search for in the PDF
- **comment**: The annotation text to add next to the tag

Example:

| tag     | comment                   |
|---------|---------------------------|
| REF-001 | This is a reference point |
| NOTE-A  | Important section         |

### Running the Application
Simply run the script:
```bash
python CommentPdf.py
```

The application will guide you through the following steps via GUI dialogs:

1. **Select Excel File** - Choose your Excel file containing tags and comments
2. **Select PDF Folder** - Choose the folder containing PDF files to annotate
3. **Enter Comment Subject** - Specify the subject for annotations (default: "Comment")
4. **Enter Annotation Distance** - Set the distance in points between tags and annotations (default: 10)

### Output
- Annotated PDF files are saved in the **same folder** as the input PDFs
- Output files have "_marked" suffix (e.g., `document.pdf` → `document_marked.pdf`)
- Original PDF files remain unchanged

## Configuration Options

### Annotation Distance
- Controls the horizontal spacing between the found tag and the annotation box
- Measured in points (1 point ≈ 1/72 inch)
- Default value: 10 points
- Can be left empty to use the default value
- Must be a non-negative integer

### Comment Subject
- Sets the subject field for all annotations
- Default value: "Comment"
- Can be customized for different annotation categories

## Example Workflow
1. Prepare an Excel file with your tags and comments
2. Run `python CommentPdf.py`
3. Select your Excel file when prompted
4. Select the folder containing your PDF files
5. Enter comment subject (or press Enter for default)
6. Enter annotation distance (or press Enter for default value of 10)
7. Wait for processing to complete
8. Find annotated PDFs with "_marked" suffix in the same folder

## Notes
- The tool searches for exact text matches of tags in PDF content
- Each occurrence of a tag will receive an annotation
- Annotations appear as yellow text boxes with dashed borders
- All PDFs in the selected folder will be processed automatically

## Troubleshooting
- **No PDF files found**: Ensure your folder contains .pdf files
- **Tag not found**: Verify the tag text exactly matches text in the PDF
- **Excel read error**: Ensure your Excel file has 'tag' and 'comment' columns
- **Permission error**: Ensure you have write permissions in the PDF folder

## Contributing
Contributions, issues, and feature requests are welcome. To contribute:
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Make your changes and add tests where applicable
4. Open a pull request describing your changes

## License
This project is open source. Please check the repository for license details.

## Contact / Support
For questions or help, open an issue on this repository or contact the maintainer: apkarthik1986

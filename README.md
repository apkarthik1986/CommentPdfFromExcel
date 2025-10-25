# CommentPdfFromExcel

A small utility to extract comments/notes from Excel workbooks (.xlsx) and generate a clean, paginated PDF report listing those comments with context (sheet name, cell address, author). This is useful for reviewers, audit trails, or sharing feedback from spreadsheets in a printable format.

## Features
- Extracts cell comments and notes from .xlsx files
- Includes comment context: sheet name, cell address, author, and timestamp (when available)
- Outputs a well-formatted, paginated PDF report
- Basic filtering and sorting options (e.g., by sheet, author, or date)
- Example scripts and sample workbooks included (if present in the repo)

## Requirements
- Python 3.8+ recommended
- pip for installing dependencies

If the repository includes a requirements.txt, install dependencies with:
```
pip install -r requirements.txt
```

Common libraries you might expect:
- openpyxl (for reading .xlsx files)
- reportlab or fpdf2 (for PDF generation)
- pillow (optional, for handling images)

## Installation
1. Clone the repository:
```
git clone https://github.com/apkarthik1986/CommentPdfFromExcel.git
cd CommentPdfFromExcel
```
2. Install dependencies:
```
pip install -r requirements.txt
```
(If there's no requirements.txt, install the needed packages individually, e.g. `pip install openpyxl reportlab`.)

## Usage
Run the provided script to generate a PDF from an Excel file. Example:
```
python extract_comments.py input.xlsx -o comments.pdf
```
Notes:
- Replace `extract_comments.py` with the actual script/module name if different.
- Use `-h` or `--help` (e.g., `python extract_comments.py -h`) to view available CLI options such as filtering, sorting, or page templates.
- If the repository exposes a package entry point or a different CLI, follow that pattern instead.

Example command with filters (example flags; adjust to the repo's actual CLI):
```
python extract_comments.py input.xlsx --sheet "Sheet1" --author "Alice" -o comments.pdf
```

## File structure (example)
- extract_comments.py      # Main script that reads .xlsx and writes a PDF
- README.md               # This file
- requirements.txt        # Python dependencies (if applicable)
- examples/               # Example workbooks and generated PDFs
- tests/                  # Unit tests (if present)
- docs/                   # Additional documentation (if present)

Adjust this section to match the repository's actual layout.

## Examples
Include sample input files and generated PDFs under `examples/` so users can quickly test functionality:
- examples/sample.xlsx -> examples/sample_comments.pdf

## Configuration
If the tool supports a config file (YAML/JSON) or environment variables, document the keys and examples here. Example:
```
# config.yml
output:
  page_size: A4
  font: "Helvetica"
filters:
  include_empty_comments: false
```

## Contributing
Contributions, issues, and feature requests are welcome. To contribute:
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Make your changes and add tests where applicable
4. Open a pull request describing your changes

Please follow any contribution guidelines or templates present in the repo.

## Testing
If tests are included, run them with:
```
pytest
```
(or use the project’s preferred test runner/command)

## License
Specify the project license here (e.g., MIT). If a LICENSE file exists in this repository, reference it:
```
This project is licensed under the MIT License - see the LICENSE file for details.
```

## Troubleshooting & Tips
- If comments do not appear, verify whether they are stored as "notes" vs "comments" in the Excel file format—different Excel versions store them differently.
- Large workbooks may require more time/memory; try filtering by sheet or author to reduce output size.
- For non-ASCII authors or text, ensure the selected PDF font supports the characters you need.

## Contact / Support
For questions or help, open an issue on this repository or contact the maintainer: apkarthik1986

---

If you'd like any adjustments (badges, CI status, exact script names, or improved examples), tell me what to change and I'll update the README content. When you're ready for me to apply this change to the repository, approve this content and I will commit it.

# BISE Result Extractor

A Python-based utility for extracting and exporting BISE Faisalabad result details from the official results website. The project can fetch student results over a roll-number range, detect pass/supply/fail classifications, and export the findings to Excel and HTML reports.

This repository includes:

- `app.py` — bulk result extractor for roll-number ranges
- `finwor.py` — a small PDF search utility for locating names in a PDF gazette

## Features

- Fetches result pages for a numeric roll-number range
- Handles common website variations and no-result/invalid-roll states safely
- Parses student name, father name, total marks, notification, and academic status
- Classifies academic outcome as:
  - regular/pass
  - supply
  - failed
  - absent
  - withheld/RL
  - other/unknown
- Exports results to:
  - Excel workbook (`.xlsx`)
  - HTML dashboard (`.html`)
- Includes retry logic, concurrency limits, and parser fallback behavior for robustness

## Requirements

Python 3.10+ is recommended.

Install the required dependencies:

```bash
pip install aiohttp beautifulsoup4 openpyxl PyMuPDF
```

## Project Structure

```text
BISE Result Extractor/
├── app.py
├── finwor.py
├── README.md
└── output files generated at runtime
```

## Using the bulk result extractor

Run:

```bash
python app.py
```

The script will prompt for the roll-number range and generate output files in the working directory.

### Example flow

```text
Enter starting roll number: 100001
Enter ending roll number: 100050
```

The tool then:

1. validates the range,
2. downloads each result page,
3. parses the result details,
4. saves an Excel workbook with the results,
5. writes an HTML report for review.

## Using the PDF search utility

`finwor.py` searches a local PDF file for a specific student name and prints matching pages/lines around the result.

Edit the script variables as needed:

```python
pdf_path = Path.home() / "Desktop" / "Gazette 10th Class FIRST ANNUAL 2026.pdf"
search_name = "Tahreem Kashif"
```

Then run:

```bash
python finwor.py
```

## Important notes

- This project interacts with a public board website and should be used responsibly.
- Respect website rate limits and avoid excessive automated requests.
- The code intentionally avoids bypassing CAPTCHAs, authentication, or access controls.
- Result pages can change without notice; parser logic may require minor updates if the board changes its HTML structure.

## License

This repository does not currently include a license file. If you plan to publish or reuse it publicly, add an appropriate open-source license such as MIT or Apache 2.0.

## Contributing

Contributions are welcome. If you improve parsing logic, add support for new result-page layouts, or improve the report output, feel free to open a pull request.

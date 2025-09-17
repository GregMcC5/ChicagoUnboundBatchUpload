# Chicago Unbound Batch Tool

## Project Overview

**Chicago Unbound Batch Tool** is an application designed to streamline the annual batch upload of Law School faculty scholarly outputs to [Chicago Unbound](https://chicagounbound.uchicago.edu/), the University of Chicago Law School’s institutional repository. This tool automates the conversion of LawCites CSV exports into the XLS format required for submission, maps and cleans metadata, and improves data quality and efficiency. It features a simple GUIT for ease of use.

This tool was developed by Gregory McCollum, the D'Angelo Law Library's Metadata & Repository Librarian. It was developed in late 2024 and initiatlly deployed in late 2025.


## Features

- **Data Mapping:** Converts LawCites CSV reports to Digital Commons XLS batch upload spreadsheets, handling complex field mapping.
- **GUI Interface:** User-friendly graphical menu for selecting publication type and input/export paths.
- **Encoding Error Prevention**: Identifies non-ASCII characters for review.
- **Catalog Linking:** Automatically generates links to UChicago Library catalog holdings when bib numbers are present.
- **Link Validation:** Validates external resource links verifies HTTP status, and prompts user action if ambiguous.
- **Duplication Prevention:** Detects and isolates potential duplicate content using fuzzy matching.

## Python Library Requirements

  - pandas
  - openpyxl
  - pyexcel
  - fuzzywuzzy
  - requests
  - tkinter (for GUI)

## Usage

1. **Prepare Input:**
   - Export your LawCites data as a CSV file.

2. **Run the Tool:**
   ```bash
   python csv_to_xls_gui_w_preprocess.py

   - The GUI will prompt you to:
        - Select the publication type
        - Choose the LawCites CVS input file.
        - Specify output XLS file location.
        - Enable link validation (if desired)

3. **Review Outputs:**
    - The output XLS will be saved to your specified location.
    - Navigate to the Batch Upload interface in the the Chicago Unbound page corresponding to the item type you are uploading to.
    - Under the "Upload spreadsheet" option, click "Choose File" and select the specified XLS output file for upload.

    - A separate CSV file (review.csv) will list entries flagged for review (potential duplicates, encoding errors, etc).

## Contact

Please contact Gregory McCollum at gregmcc@uchicago.edu for support.


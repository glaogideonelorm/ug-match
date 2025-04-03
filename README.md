```markdown
# Name Matching Tool

## Overview

This tool is a Python script that performs fuzzy matching between two columns in an Excel file. It compares names from the **CandName** and **UGName** columns using fuzzy matching techniques powered by [FuzzyWuzzy](https://github.com/seatgeek/fuzzywuzzy) and the [Levenshtein distance](https://en.wikipedia.org/wiki/Levenshtein_distance) algorithm. The script outputs a new Excel file listing the name pairs along with a new column that indicates whether the names are considered a match ("Y") or not ("X") based on a similarity threshold.

## Features

- **Fuzzy Matching:** Utilizes `fuzz.token_sort_ratio` from FuzzyWuzzy to handle minor misspellings and omissions.
- **Customizable Threshold:** The default matching threshold is set to 80 (this can be adjusted in the source code).
- **Command-Line Arguments:** Allows users to specify input and output file names via command-line flags without editing the source code.
- **User-Friendly:** Simply place your Excel file in the project root directory and run the script.

## Dependencies

This project requires the following Python libraries:

- **pandas** – For reading and writing Excel files.
- **fuzzywuzzy** – For performing fuzzy string matching.
- **python-Levenshtein** – Enhances the performance of fuzzy matching.
- **openpyxl** – For writing output Excel files in XLSX format.
- **xlrd** – For reading input Excel files in XLS format.
- **argparse** – (Built-in) For parsing command-line arguments.

## Installation

1. **Clone or Download the Repository:**

   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo
   ```

2. **(Optional) Create a Virtual Environment:**

   ```bash
   python -m venv myenv
   source myenv/bin/activate   # On Windows: myenv\Scripts\activate
   ```

3. **Install the Required Packages:**

   Ensure you have a `requirements.txt` file in your project root with the following content:

   ```
   pandas
   fuzzywuzzy
   python-Levenshtein
   openpyxl
   xlrd
   ```

   Then install the dependencies with:

   ```bash
   pip install -r requirements.txt
   ```

## How It Works

- **Input:**  
  The script reads an Excel file (default: `NameCheck (11).xls`) which must contain two columns: **CandName** and **UGName**.

- **Fuzzy Matching:**  
  - The script preprocesses the names by trimming extra spaces and converting them to lowercase.
  - It then uses `fuzz.token_sort_ratio` to calculate a similarity score based on the Levenshtein distance.
  - If the score is equal to or greater than the threshold (default: 80), the names are marked as a match ("Y"); otherwise, they are marked as not matching ("X").

- **Output:**  
  A new Excel file (default: `NameCheck_Output.xlsx`) is generated. This file includes the original name pairs along with a new column labeled **Match** that indicates the fuzzy matching result.

## Usage

1. **Place Your Input File:**  
   Ensure your Excel file (e.g., `NameCheck (11).xls`) is located in the root directory of the project.

2. **Run the Script with Default File Names:**

   ```bash
   python index.py
   ```

   This command will process `NameCheck (11).xls` and output `NameCheck_Output.xlsx`.

3. **Specify Custom File Names Using Command-Line Flags:**

   ```bash
   python index.py --input myfile.xls --output result.xlsx
   ```

   **Command-Line Arguments:**

   - `--input`: (Optional) The input Excel file. Defaults to `NameCheck (11).xls`.
   - `--output`: (Optional) The output Excel file. Defaults to `NameCheck_Output.xlsx`.

## Troubleshooting

- **ModuleNotFoundError:**  
  If you encounter an error like `ModuleNotFoundError: No module named 'openpyxl'`, ensure that all dependencies are installed:
  ```bash
  pip install -r requirements.txt
  ```

- **File Not Found:**  
  Verify that your input file is in the correct directory or specify the correct path using the `--input` flag.

## Acknowledgements

- [FuzzyWuzzy](https://github.com/seatgeek/fuzzywuzzy)
- [Python-Levenshtein](https://github.com/ztane/python-Levenshtein)
- [pandas](https://github.com/pandas-dev/pandas)
- [openpyxl](https://openpyxl.readthedocs.io)
- [xlrd](https://xlrd.readthedocs.io)


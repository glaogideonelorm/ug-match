import argparse
import pandas as pd
from fuzzywuzzy import fuzz

def is_fuzzy_match(name1, name2, threshold=80):
    if pd.isnull(name1) or pd.isnull(name2):
        return False
    name1 = str(name1).strip().lower()
    name2 = str(name2).strip().lower()
    score = fuzz.token_sort_ratio(name1, name2)
    return score >= threshold

def process_file(input_file, output_file):
    df = pd.read_excel(input_file, engine="xlrd")
    df.columns = df.columns.str.strip()
    if 'CandName' not in df.columns or 'UGName' not in df.columns:
        raise ValueError("The input file must contain 'CandName' and 'UGName' columns.")
    df['Match'] = df.apply(lambda row: "Y" if is_fuzzy_match(row['CandName'], row['UGName']) else "X", axis=1)
    non_matches = df[df['Match'] == "X"]
    non_matches["Highlighted non matches"] = non_matches.apply(lambda row: f"{str(row['CandName']).strip()} | {str(row['UGName']).strip()} (X)", axis=1)
    non_matches_df = non_matches[["Highlighted non matches"]]
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
         df.to_excel(writer, index=False, sheet_name="Data")
         non_matches_df.to_excel(writer, index=False, sheet_name="Highlighted non matches")
    print(f"Output saved to {output_file}")

def main():
    parser = argparse.ArgumentParser(description="Process an Excel file to perform fuzzy matching on 'CandName' and 'UGName'.")
    parser.add_argument("--input", help="Name of the input Excel file (default: 'NameCheck (11).xls')", default="NameCheck (11).xls")
    parser.add_argument("--output", help="Name of the output Excel file (default: 'NameCheck_Output.xlsx')", default="NameCheck_Output.xlsx")
    args = parser.parse_args()
    process_file(args.input, args.output)

if __name__ == "__main__":
    main()

import argparse
import datetime
import numpy as np
import os
import pandas as pd
import pdf2image
import pdftotext
import pytesseract
import re
import xlsxwriter

def parse_pdf_to_fwf(folder, file):

    if file[:-4] != ".pdf":
        Exception("Input file should have extension .pdf")

    with open(os.path.join(folder, file), "rb") as in_pdf:
        pdf = pdftotext.PDF(in_pdf)

    text = pdf[0]
    if  len(text) ==  0:
        print("No text found; pdf file {} contains an image\nConverting to pdf with text...\n".format(file))
        text = get_text_from_image_pdf(folder, file)

    # Remove "|" characters (sometimes carried in from tesseract-generated pdfs)
    text = text.replace("|", "")

    # Split into separate lines
    split_text = text.split("\n")

    # Extract the year and month for which this report was generated and convert from financial calendar to date
    m = re.search(r"BALANCE\s+CONTROL\s+(?P<year>\d{4})/(?P<month>\d{2})", split_text[1])
    year = int(m["year"])
    month = int(m["month"]) + 3
    if month > 12:
        year += 1
        month -= 12
    date_str = datetime.date(year, month, 1).strftime("%b-%Y")

    # Write out a txt file containing the main table
    output_file = file[:-4] + ".txt"
    with open(os.path.join(folder, output_file), "w") as outfile:
        [outfile.write("%s\n" % l) for l in split_text[2:-1]]

    return output_file, date_str


def get_text_from_image_pdf(folder, file):

    output_pdf_file = os.path.join(folder, "text_output.pdf")

    # Convert the pdf to a png image
    pdf2image.convert_from_path(os.path.join(folder, file), fmt="png", single_file=True, output_folder=args.folder, output_file="tmp")

    # Generate a pdf with selectable text based on the image
    pdf = pytesseract.image_to_pdf_or_hocr(os.path.join(folder, "tmp.png"), extension="pdf")
    with open(output_pdf_file, "w+b") as f:
        f.write(pdf)

    # Extract the text from the pdf
    with open(output_pdf_file, "rb") as in_pdf:
        pdf = pdftotext.PDF(in_pdf)

    # Check that we've actually got some text out from the converted file
    text = pdf[0]
    if len(text) == 0:
        Exception("Could not extract text from converted pdf file")

    return text


def fwf_to_df(folder, file):

    df = pd.read_fwf(os.path.join(folder, file), colspecs="infer")

    return df


def check_and_clean_df(df):

    expected_columns = ["ELEMENT DETAILS", "BROUGHT FORWARD", "THIS PERIOD", "ADJUSTMENT", "CARRIED FORWARD"]
    financial_columns = ["BROUGHT FORWARD", "THIS PERIOD", "ADJUSTMENT", "CARRIED FORWARD"]

    # Change multiple consecutive spaces to a single space
    # (tesseract-generated pdfs often insert more spaces between words than intended, even though the overall formatting is generally fine)
    spaces_pattern = re.compile(r"\s+")
    df = df.rename(columns={col: spaces_pattern.subn(" ", col)[0] for col in list(df.columns)})

    # Check that the expected columns are present

    if len(list(df)) > len(expected_columns):
        # Quick check: See if there are too many columns because two-word col names have been split
        df = combine_column_check(df, expected_columns)
        if len(list(df)) > len(expected_columns):
            raise Exception("Number of columns is greater than expected:\n{}".format(list(df)))
    elif len(list(df)) < len(expected_columns):
        raise Exception("Number of columns is less than expected:\n{}".format(list(df)))

    for col, expcol in zip(list(df), expected_columns):
        if col != expcol:
            raise Exception("Expected column name to be {}, not {}".format(expcol, col))

    # Check first column (should be three-parts, parse and separate, remove total row and retain for final check)
    df["Code"], df["Details"], df["Further details"] = zip(*df["ELEMENT DETAILS"].apply(parse_element_details))

    # Remove any additional spaces from second - fifth columns (another potential issue in tesseract-generated pdfs)
    df[financial_columns] = df[financial_columns].replace({" ": ""}, regex=True)

    # Check second - fifth columns (should be numbers with max 2 dp, move trailing "-" character to start when negative)
    for col in financial_columns:
        df[col] = df[col].apply(parse_and_check_financial_columns)

    # Drop old element details column and move total row to a separate table
    df = df.drop(columns=["ELEMENT DETAILS"])
    total = df.loc[df["Further details"] == "TOTAL"]
    df.drop(df.loc[df["Further details"] == "TOTAL"].index, inplace=True)

    # Check totals
    for col in financial_columns:
        expected_total = 0.0 if pd.isna(total[col].iloc[0]) else total[col].iloc[0] # NaN values will sum to 0, so make sure TOTAL is in the same form if empty
        if df[col].sum().round(2) != expected_total:
            raise Exception("Total of column '{}' ({}) does not match value from original document ({})".format(col, df[col].sum().round(2), expected_total))

    return df

def combine_column_check(df, expected_columns):

    # If columns are empty, it can lead to difficulties identifying the column boundaries
    # See if any of the columns are empty, and if so, whether the column names have originated from a split at a space in the name

    empty_cols = [col for col in df.columns if df[col].isnull().all()]

    # See if any of the pairs of empty columns match the names of the expected columns
    for c1, c2 in zip(empty_cols[:-1], empty_cols[1:]):
        if c1 + " " + c2 in expected_columns:
            print("Columns '{}' and '{}' -> '{}' found in expected column names list".format(c1, c2, c1 + " " + c2))
            print("Removing empty column '{}' and combining with empty column '{}'...\n".format(c1, c2))
            df = df.drop(columns=c1)
            df = df.rename(columns={c2: c1 + " " + c2})
        break

    return df


def parse_element_details(elt_details):

    # This regex should catch:
    #   - code: (optional) four (or more) digit number (ELEMENT DETAILS in original output spreadsheet)
    #   - short: (optional) one to six characters (no spaces), bounded at end by at least two consecutive spaces (Details in original)
    #   - long: characters, possibly separated by single spaces (Further Details in original)

    m = re.match(r"(?P<code>\d{4}) +(?P<short>\S{1,6}(?=\s{2}))? +(?P<long>(\S+\s?)+)", elt_details)
    if m is None:
        if elt_details == "TOTAL":
            return np.nan, np.nan, "TOTAL"
        else:
            raise Exception("Cannot parse row '{}'".format(elt_details))
    else:
        return int(m["code"]), m["short"] if m["short"] is not None else np.nan, m["long"]


def parse_and_check_financial_columns(value):

    # All values in these columns should either refer to monetary values (two decimal places, optional -ve sign at end)
    # or be NaN

    # No need to check for minus sign if value is empty or already in numerical form
    if pd.isna(value) or isinstance(value, float):
        return value

    m = re.match(r"(?P<value>\d+.\d{2}){1}(?P<sign>-)?$", value)
    if m is None:
        raise Exception("Unexpected entry format in column: '{}'".format(value))
    elif m["sign"] is None:
        return float(m["value"])
    elif m["sign"] == "-":
        return -float(m["value"])
    else:
        raise Exception("Unexpected entry format in column: '{}'".format(value))


def combine_dfs(first, second, first_date, second_date):

    # Combine into a single dataframe and rename "THIS PERIOD" columns so that they are named by month
    combined_df = pd.merge(first, second, on=["Code", "Details", "Further details"], suffixes=[" 1", " 2"])
    combined_df = combined_df.rename(columns={"THIS PERIOD 1": first_date, "THIS PERIOD 2": second_date})

    return combined_df


def write_to_xlsx(df, first_date, second_date, folder, filename):

    # Construct our list of columns to output, and add formula for equality check
    columns = [{"header": "Code"},
               {"header": "Details"},
               {"header": "Further details", "total_string": "TOTAL"},
               {"header": first_date, "total_function": "sum"},
               {"header": second_date, "total_function": "sum"},
               {"header": "Equal", "formula": "EXACT([@[{}]],[@[{}]])".format(first_date, second_date)},
               {"header": "Checked by HR"},
               {"header": "Checked by Finance"}]

    # Get the columns that we will output from the data frame
    df = df[[c["header"] for c in columns[:5]]]

    # Create an Excel workbook with a table ready to take the contents of the df, and initialise headers
    workbook = xlsxwriter.Workbook(os.path.join(folder, filename))
    worksheet = workbook.add_worksheet()
    worksheet.add_table(0, 0, len(df.index) + 1, len(columns) - 1,
                        {"columns": columns,
                         "banded_rows": False,
                         "style": "Table Style Medium 9",
                         "total_row": True})

    # We'll need some of the columns to have particular formats
    format_code = workbook.add_format({"num_format": "0000"})
    format_financial = workbook.add_format({"num_format": "#,##0.00"})
    format_issue = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

    # Put empty strings where no data is available, then write the df to the Excel sheet
    df = df.fillna("")
    for index, row in df.iterrows():
        worksheet.write_row(index + 1, 0, list(row))

    # Apply formatting and make columns wider
    worksheet.set_column(0, 0, 8, format_code)
    worksheet.set_column(1, 1, 12)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 16, format_financial)
    worksheet.set_column(4, 4, 16, format_financial)
    worksheet.conditional_format(1, 5, 20, 5,
                                 {"type": "cell", "criteria": "!=", "value": "TRUE", "format": format_issue})
    worksheet.set_column(6, 7, 16)

    # Close workbook to finish
    workbook.close()


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", type=str, default=os.getcwd(), help="Path to directory containing pdf reports")
    parser.add_argument("--first", type=str, help="Name of the file containing the first month's data")
    parser.add_argument("--second", type=str, help="Name of the file containing the second month's data")
    parser.add_argument("--output", type=str, default="output.xlsx", help="Name of the output Excel file")
    args = parser.parse_args()

    first_fwf, first_date = parse_pdf_to_fwf(args.folder, args.first)
    first_df = fwf_to_df(args.folder, first_fwf)
    first_df = check_and_clean_df(first_df)

    second_fwf, second_date = parse_pdf_to_fwf(args.folder, args.second)
    second_df = fwf_to_df(args.folder, second_fwf)
    second_df = check_and_clean_df(second_df)

    # While testing, we often use the same file twice so first_date is equal to second_date - so add a fix here for now
    if first_date == second_date:
        second_date = second_date + " (rep)"
        print("\nRepeated dates found...")
        print("Second occurrence of {} set to {}".format(first_date, second_date))


    combined_df = combine_dfs(first_df, second_df, first_date, second_date)
    write_to_xlsx(combined_df, first_date, second_date, args.folder, args.output)

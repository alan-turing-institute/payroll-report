import argparse
import numpy as np
import os
import pandas as pd
import pdftotext
import re


def parse_pdf_to_fwf(folder, file):

    if file[:-4] != ".pdf":
        Exception("Input file should have extension .pdf")

    with open(os.path.join(folder, file), "rb") as in_pdf:
        pdf = pdftotext.PDF(in_pdf)

    text = pdf[0]
    split_text = text.split("\n")

    output_file = file[:-4] + ".txt"

    with open(os.path.join(folder, output_file), "w") as outfile:
        [outfile.write("%s\n" % l) for l in split_text[2:-1]]

    return output_file


def fwf_to_df(folder, file):

    df = pd.read_fwf(os.path.join(folder, file), colspecs="infer")

    return df


def check_and_clean_df(df):

    expected_columns = ["ELEMENT DETAILS", "BROUGHT FORWARD", "THIS PERIOD", "ADJUSTMENT", "CARRIED FORWARD"]
    financial_columns = ["BROUGHT FORWARD", "THIS PERIOD", "ADJUSTMENT", "CARRIED FORWARD"]

    # Check columns

    if len(list(df)) != len(expected_columns):
        raise Exception("Number of columns is not {}, as expected".format(len(expected_columns)))

    for col, expcol in zip(list(df), expected_columns):
        if col != expcol:
            raise Exception("Expected column name to be {}, not {}".format(expcol, col))

    # Check first column (should be three-parts, parse and separate, remove total row and retain for final check)
    df["Code"], df["Short"], df["Long"] = zip(*df["ELEMENT DETAILS"].apply(parse_element_details))

    # Check second - fifth columns (should be numbers with max 2 dp, move ending - to start for negative)
    for col in financial_columns:
        df[col] = df[col].apply(parse_and_check_financial_columns)

    # Drop old element details column and move total row to a separate table
    df = df.drop(columns=["ELEMENT DETAILS"])
    total = df.loc[df["Long"] == "TOTAL"]
    df.drop(df.loc[df["Long"] == "TOTAL"].index, inplace=True)

    # Check totals
    for col in financial_columns:
        if df[col].sum().round(2) != total[col].iloc[0]:
            raise Exception("Total of column '{}' does not match value from original document".format(col))

    return df


def parse_element_details(elt_details):

    # This regex should catch:
    #   - code: (optional) four (or more) digit number
    #   - short: (optional) one to six characters (no spaces), bounded at end by at least two consecutive spaces
    #   - long: characters, possibly separated by single spaces

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

    if pd.isna(value):
        return value

    m = re.match(r"(?P<value>\d+.\d{2}){1}(?P<sign>-)?", value)
    if m["sign"] is None:
        return float(m["value"])
    elif m["sign"] == "-":
        return -float(m["value"])
    else:
        raise Exception("Unexpected entry in column: '{}'".format(value)) 


def combine_dfs(first, second):

    # Combine into a single dataframe

    return pd.merge(first, second, on=["Code", "Short", "Long"], suffixes=[" 1", " 2"])


def write_to_xlsx(df, folder, filename):

    output_columns = ["Code", "Short", "Long", "THIS PERIOD 1", "THIS PERIOD 2"]

    output_df = df[output_columns]
    output_df = output_df.assign(Equal=False)
    output_df['Equal'] = np.where(((output_df["THIS PERIOD 1"] == output_df["THIS PERIOD 2"]) |
                                  (output_df["THIS PERIOD 2"].isnull() & output_df["THIS PERIOD 2"].isnull())),
                                  True, False)

    writer = pd.ExcelWriter(os.path.join(folder, filename), engine="xlsxwriter")
    output_df.to_excel(writer, sheet_name="Sheet1")

    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    format_code = workbook.add_format({"num_format": "0000"})
    format_financial = workbook.add_format({"num_format": "#,##0.00"})

    worksheet.set_column(1, 1, 9.17, format_code)
    worksheet.set_column(2, 2, 9.17)
    worksheet.set_column(3, 3, 22.5)
    worksheet.set_column(4, 4, 16.67, format_financial)
    worksheet.set_column(5, 5, 16.67, format_financial)

    writer.save()


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", type=str, default=os.getcwd(), help="Path to directory containing pdf reports")
    parser.add_argument("--first", type=str, help="Name of the file containing the first month's data")
    parser.add_argument("--second", type=str, help="Name of the file containing the second month's data")
    parser.add_argument("--output", type=str, default="output.xlsx", help="Name of the output Excel file")
    args = parser.parse_args()

    first_fwf = parse_pdf_to_fwf(args.folder, args.first)
    first_df = fwf_to_df(args.folder, first_fwf)
    first_df = check_and_clean_df(first_df)

    second_fwf = parse_pdf_to_fwf(args.folder, args.second)
    second_df = fwf_to_df(args.folder, first_fwf)
    second_df = check_and_clean_df(second_df)

    combined_df = combine_dfs(first_df, second_df)
    write_to_xlsx(combined_df, args.folder, args.output)

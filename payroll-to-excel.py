import argparse
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

    # Check columns

    if len(list(df)) != len(expected_columns):
        raise Exception("Number of columns is not {}, as expected".format(len(expected_columns)))

    for col, expcol in zip(list(df), expected_columns):
        if col != expcol:
            raise Exception("Expected column name to be {}, not {}".format(expcol, col))

    # Check first column (should be three-parts, parse and separate, remove total row and retain for final check)
    df["Code"], df["Short"], df["Long"] = zip(*df["ELEMENT DETAILS"].apply(parse_element_details))

    # Drop total row for now
    total = df.loc[df["Long"] == "TOTAL"]
    df.drop(df.loc[df["Long"] == "TOTAL"].index, inplace=True)

    print(df)

    # Check second column (should be numbers with max 2 dp, change NaN -> 0, move ending - to start for negative)

    # Check third column (should be numbers with max 2 dp, change NaN -> 0, move ending - to start for negative)

    # Check fourth column (should be numbers with max 2 dp, change NaN -> 0, move ending - to start for negative)

    # Check fifth column (should be numbers with max 2 dp, change NaN -> 0, move ending - to start for negative)

    # Check total


def parse_element_details(elt_details):

    # This regex should catch:
    #   - code: (optional) four (or more) digit number
    #   - short: (optional) one to six characters (no spaces), bounded at end by at least two consecutive spaces
    #   - long: characters, possibly separated by single spaces

    m = re.match(r"(?P<code>\d{4}) +(?P<short>\S{1,6}(?=\s{2}))? +(?P<long>(\S+\s?)+)", elt_details)
    if m is None:
        if elt_details == "TOTAL":
            return None, None, "TOTAL"
        else:
            raise Exception("Cannot parse row '{}'".format(elt_details))
    else:
        return m["code"], m["short"], m["long"]


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", type=str, default=os.getcwd(), help="Path to directory containing pdf reports")
    parser.add_argument("--first", type=str, help="Name of the file containing the first month's data")
    args = parser.parse_args()

    first_fwf = parse_pdf_to_fwf(args.folder, args.first)
    first_df = fwf_to_df(args.folder, first_fwf)
    first_df = check_and_clean_df(first_df)


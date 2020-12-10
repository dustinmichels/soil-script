import pandas as pd
import os
from datetime import date


KEY_SPREADSHEET = "POXC_FinalData.xlsx"
DATA_SHEET_NAME = "POX-C Calculation"
DATA_DIR = "data"
OUT_DIR = "output"


def load_and_merge(filename, key_df):
    # load source data
    filepath = os.path.join(DATA_DIR, filename)
    df = pd.read_excel(filepath, sheet_name=DATA_SHEET_NAME, header=0)
    df["source sheet"] = filename

    # merge on columns 'depth' & 'plot'
    merged_df = df.merge(key_df, on=["plot", "depth"], how="left")
    return merged_df


def save_outout(full_df):
    # create output folder if it doesn't exist
    try:
        os.mkdir(OUT_DIR)
    except FileExistsError:
        pass

    # save dataframe to file with today's date
    today_str = date.today().strftime("%m-%d-%Y")
    filename = f"output_{today_str}.xlsx"
    full_df.to_excel(os.path.join(OUT_DIR, filename), index=False)


def main():
    # load key/treatment data
    key_df = pd.read_excel(os.path.join(DATA_DIR, KEY_SPREADSHEET), header=0)
    key_df.columns = [x.lower() for x in key_df.columns]

    # iterate over files, merging with keys
    dfs = []
    for filename in os.listdir(DATA_DIR):
        if filename != KEY_SPREADSHEET:
            merged_df = load_and_merge(filename, key_df)
            dfs.append(merged_df)

    # concatenate merged dfs into single datafame
    full_df = pd.concat(dfs)

    # save output to file
    save_outout(full_df)


if __name__ == "__main__":
    main()

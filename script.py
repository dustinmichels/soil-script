import os
from datetime import date

import pandas as pd
from termcolor import colored, cprint

# CONGIG
KEY_SPREADSHEET = "POXC_FinalData.xlsx"
DATA_SHEET_NAME = "POX-C Calculation"
DATA_DIR = "data"
OUT_DIR = "output"


def load_and_merge(filename, key_df):
    """
    Open a data file and merge it with key dataframe
    """

    # load source data
    filepath = os.path.join(DATA_DIR, filename)
    df = pd.read_excel(filepath, sheet_name=DATA_SHEET_NAME, header=0)

    # add source sheet column
    df["source sheet"] = filename

    # merge on columns 'depth' & 'plot'
    merged_df = df.merge(key_df, on=["plot", "depth"], how="left")
    return merged_df


def save_outout(full_df):
    """
    Save complete DataFrame to excel file
    """

    # create output folder if it doesn't exist
    try:
        os.mkdir(OUT_DIR)
    except FileExistsError:
        pass

    # save dataframe to file with today's date
    today_str = date.today().strftime("%m-%d-%Y")
    # filename = f"output_{today_str}.xlsx"
    # full_df.to_excel(os.path.join(OUT_DIR, filename), index=False)
    filename = f"output_{today_str}.csv"
    full_df.to_csv(os.path.join(OUT_DIR, filename), index=False)

    # print info
    cprint("\nDone!", "green", attrs=["bold"])
    print(f"> Saved result to file '{os.path.join(OUT_DIR, filename)}'\n")


def get_input():
    """
    Prompt for input which would override default, global values
    """

    global DATA_DIR, OUT_DIR, KEY_SPREADSHEET, DATA_SHEET_NAME

    boldify = lambda x: colored(f"[{x}]", attrs=["bold"])

    data_dir = input(f"Path to data directory? {boldify(DATA_DIR)}: ")
    if data_dir:
        DATA_DIR = data_dir

    out_dir = input(f"Where to output files? {boldify(OUT_DIR)}: ")
    if out_dir:
        OUT_DIR = out_dir

    key_spreadsheet = input(
        f"Name of the treatment sheet? {boldify(KEY_SPREADSHEET)}: "
    )
    if key_spreadsheet:
        KEY_SPREADSHEET = key_spreadsheet

    data_sheet_name = input(
        f"Name of data tab within data sheets? {boldify(DATA_SHEET_NAME)}: "
    )
    if data_sheet_name:
        DATA_SHEET_NAME = data_sheet_name


def main():
    """
    Iterate over data files, merging with key data and concatenating into single DataFrame
    """

    # load key/treatment data
    key_df = pd.read_excel(os.path.join(DATA_DIR, KEY_SPREADSHEET), header=0)
    key_df.columns = [x.lower() for x in key_df.columns]

    # iterate over files, merging with keys
    dfs = []
    for filename in os.listdir(DATA_DIR):
        if filename != KEY_SPREADSHEET:
            try:
                merged_df = load_and_merge(filename, key_df)
                dfs.append(merged_df)
                print(f"> Succesfully handled {filename}")
            except:
                print(f"> Warning! Error with {filename}")
                pass

    # concatenate merged dfs into single datafame
    full_df = pd.concat(dfs)

    # save output to file
    save_outout(full_df)


if __name__ == "__main__":
    print("\n{:-^30}".format("CONFIG"))
    print("Press enter to use default values.\n")
    get_input()

    print("\n{:-^30}".format("RUNNING"))
    main()

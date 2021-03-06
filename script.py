import os
from datetime import date

import pandas as pd
from tabulate import tabulate
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

    # add source sheet as column
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

    # save as excel
    filename1 = f"output_{today_str}.xlsx"
    full_df.to_excel(os.path.join(OUT_DIR, filename1), index=False)

    # save as csv
    filename2 = f"output_{today_str}.csv"
    full_df.to_csv(os.path.join(OUT_DIR, filename2), index=False)

    # print info
    cprint("\nDone!", "green", attrs=["bold"])
    print(f"> Saved result to file '{os.path.join(OUT_DIR, filename1)}")
    print(f"> Saved result to file '{os.path.join(OUT_DIR, filename2)}'\n")


def data_validation(full_df, key_df):
    print("\n{:-^30}".format("CHECKING DATA"))

    # cleanup
    full_df["plot"] = full_df["plot"].astype(str)
    key_df["plot"] = key_df["plot"].astype(str)
    full_df = full_df.reset_index(drop=True)

    # look for missing keys
    actual_keys = list(key_df.set_index(["plot", "depth"]).index)
    found_keys = list(full_df.set_index(["plot", "depth"]).index)
    missing = []
    for key_tuple in actual_keys:
        if key_tuple not in found_keys:
            missing.append(key_tuple)
    if missing:
        print("\n\nYou do not have data for every key. Missing keys:\n")
        print(tabulate(missing, headers=["plot", "depth"], tablefmt="pipe"))

    # look for duplicates
    full_df_keys = full_df[["plot", "depth"]]
    duplicates = full_df_keys[full_df_keys.duplicated()]
    if not duplicates.empty:
        print("\n\nYou have some duplicated keys in your data. Duplicates:\n")
        print(duplicates.to_markdown())


def get_input():
    """
    Prompt for input which would override default, global values
    """

    global DATA_DIR, OUT_DIR, KEY_SPREADSHEET, DATA_SHEET_NAME

    print("\n{:-^30}".format("CONFIG"))
    print("Press enter to use default values.\n")

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
    print("\n{:-^30}".format("RUNNING"))

    # load key/treatment data
    key_df = pd.read_excel(os.path.join(DATA_DIR, KEY_SPREADSHEET), header=0)
    key_df.columns = [x.lower() for x in key_df.columns]

    # iterate over files, merging with keys
    dfs = []
    for filename in os.listdir(DATA_DIR):
        if (".xlsx" in filename) and (filename != KEY_SPREADSHEET):
            try:
                merged_df = load_and_merge(filename, key_df)
                dfs.append(merged_df)
                print(f"> Succesfully handled {filename}")
            except:
                print(f"> Warning! Error with {filename}")
                pass

    # concatenate merged dfs into single datafame
    full_df = pd.concat(dfs)

    # sort by plot
    full_df["plot"] = full_df["plot"].astype(str)
    full_df = full_df.sort_values(by="plot", axis=0)

    # save output to file
    save_outout(full_df)

    # validate data
    data_validation(full_df, key_df)


if __name__ == "__main__":
    # get_input()
    main()

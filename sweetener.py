#
# sweetener
# blood sugar data cleaning & export
#

import numpy as np
import pandas as pd
from datetime import datetime
from argparse import ArgumentParser

from openpyxl.utils.dataframe import dataframe_to_rows

def read_sugar_df(csv_path: str) -> pd.DataFrame:
    """Read the blood sugar data from the given CSV as a DataFrame"""
    # read sugar data csv export
    col_dtypes = {
        "Date": str,
        "Time": str,
        "Tags": str,
        "Blood Sugar Measurement (mmol/L)": np.float64,
        "Basal Injection Units": np.float64,
        "Insulin (Meal)": np.float64,
        "Insulin (Correction)": np.float64,
        "Meal Carbohydrates (Grams, Factor 1)": np.float64,
        "Meal Descriptions": str,
        "Note": str,
    }

    sugar_df = pd.read_csv(
        csv_path,
        usecols=list(col_dtypes.keys()),
        dtype=col_dtypes,
    )
    if not isinstance(sugar_df, pd.DataFrame):
        raise ValueError("Expected sugar_df to be a DataFrame")

    # parse date & time columns
    sugar_df["Date"] = pd.to_datetime(sugar_df["Date"])
    sugar_df["Time"] = pd.to_datetime(sugar_df["Time"]).apply((lambda dt: dt.time())) # type: ignore

    # filter data to only include entries from start date onwards
    if args.start_from is not None:
        sugar_df = sugar_df[sugar_df["Date"] >= args.start_from]

    return sugar_df


if __name__ == "__main__":
    # parse script command line args
    parser = ArgumentParser(description="Prepare blood sugar data for DDD evaluation")
    parser.add_argument("csv_path", help="Path to the blood sugar data CSV export.")
    parser.add_argument("--start-from", help="Starting date in format DD/MM/YYYY. "
                        "Data from the starting date onwards is included.",
                        type=(lambda date_str: datetime.strptime(date_str, "%d/%m/%Y")))
    args = parser.parse_args()

    # read blood sugar data
    sugar_df = read_sugar_df(args.csv_path)

    # add hypo / hyperglycemia features
    sugar_df["Hyperglycemia"] = sugar_df["Blood Sugar Measurement (mmol/L)"] > 10.0
    sugar_df["Hypoglycemia"] = sugar_df["Blood Sugar Measurement (mmol/L)"] < 4.0

    # compute summary statistics
    stats_df = sugar_df.describe().drop(
        ["25%", "75%"]
    )
    if stats_df is None:
        raise RuntimeError("Unexpected None returned from DataFrame.drop()")
    stats_df = stats_df.rename(index={"50%": "median"})

    # compute hyperglycemia & hypoglycemia statistics
    for prefix in ["Hyper", "Hypo"]:
        glycemia = f"{prefix}glycemia"
        count = sugar_df[sugar_df[glycemia]][glycemia].agg("count")
        stats_df.loc["count", glycemia] = count
        stats_df.loc["mean", glycemia] = count / len(sugar_df)

    # TODO: template excel spreadsheet

    print(stats_df)
